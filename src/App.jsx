import React, { useState, useEffect, useCallback, useRef } from "react";
import {
  Clock, Calendar, Users, FileSpreadsheet, Bell, Store, ClipboardList,
  History, Plus, X, Check, AlertTriangle, ChevronLeft, ChevronRight,
  UserCog, User, Trash2, Edit3, Printer, FileDown, Sunrise, Upload,
  MessageCircle, KeyRound, LogOut, Loader2, ShieldCheck
} from "lucide-react";
import * as XLSX from "xlsx";

/* ===========================================================
   ROTAWISE — cloud edition
   Backed by a live Supabase project (Postgres + Auth + RLS +
   an Edge Function for provisioning employee logins).
   Talks to Supabase over plain fetch() — no SDK needed.
=========================================================== */

const SUPABASE_URL = "https://ukddnulxifhixatpmofz.supabase.co";
const ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InVrZGRudWx4aWZoaXhhdHBtb2Z6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODc1OTI1MTEsImV4cCI6MjEwMzE2ODUxMX0.ab0Zkr4VWMWMU6vsadtXscvPSVpMb_tgdkCFGcBLQWg";
const STAFF_DOMAIN = "rotawise.staff";

const FONT_STYLE = `
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
:root{
  --bg:#12151A; --surface:#1A1F27; --surface2:#212832; --border:#2B323D;
  --amber:#E8A33D; --amber-soft:#F2C572; --text:#EDEBE4; --muted:#8B93A3;
  --ok:#52A97C; --bad:#DB5A54; --info:#5B9BD5;
}
.rw-root{ font-family:'Inter',system-ui,sans-serif; background:var(--bg); color:var(--text); }
.rw-display{ font-family:'Space Grotesk',system-ui,sans-serif; }
.rw-mono{ font-family:'IBM Plex Mono',ui-monospace,monospace; }
.rw-scrollbar::-webkit-scrollbar{ height:8px; width:8px; }
.rw-scrollbar::-webkit-scrollbar-thumb{ background:#333c48; border-radius:8px; }
.rw-punch{ position:relative; }
.rw-punch::before{
  content:''; position:absolute; left:0; top:0; bottom:0; width:1px;
  background-image:repeating-linear-gradient(to bottom, var(--border) 0 6px, transparent 6px 12px);
}
.rw-hole{ width:8px; height:8px; border-radius:50%; background:var(--bg); border:1px solid var(--border); }
@keyframes rw-fade{ from{opacity:0; transform:translateY(4px);} to{opacity:1; transform:translateY(0);} }
.rw-fade{ animation:rw-fade .18s ease-out; }
`;

/* ---------------- date/time helpers ---------------- */
const todayISO = () => new Date().toISOString().slice(0, 10);
const fmtDate = (iso) => new Date(iso + "T00:00:00").toLocaleDateString("en-GB", { weekday: "short", day: "2-digit", month: "short" });
const fmtDateShort = (iso) => new Date(iso + "T00:00:00").toLocaleDateString("en-GB", { day: "2-digit", month: "2-digit", year: "numeric" });
const addDays = (iso, n) => { const d = new Date(iso + "T00:00:00"); d.setDate(d.getDate() + n); return d.toISOString().slice(0, 10); };
const startOfWeek = (iso) => { const d = new Date(iso + "T00:00:00"); const day = (d.getDay() + 6) % 7; d.setDate(d.getDate() - day); return d.toISOString().slice(0, 10); };
const hoursBetween = (start, end, breakMins = 0) => {
  if (!start || !end) return 0;
  const [sh, sm] = start.split(":").map(Number);
  const [eh, em] = end.split(":").map(Number);
  let mins = (eh * 60 + em) - (sh * 60 + sm);
  if (mins < 0) mins += 24 * 60;
  mins -= Number(breakMins || 0);
  return Math.max(0, Math.round((mins / 60) * 100) / 100);
};
const timeShort = (t) => (t ? t.slice(0, 5) : t);
const download = (blob, filename) => {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; document.body.appendChild(a); a.click();
  a.remove(); URL.revokeObjectURL(url);
};
const exportWorkbook = (sheets, filename) => {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ name, rows }) => XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), name.slice(0, 31)));
  download(new Blob([XLSX.write(wb, { bookType: "xlsx", type: "array" })], { type: "application/octet-stream" }), filename);
};
const exportCSV = (rows, filename) => {
  const csv = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(rows));
  download(new Blob([csv], { type: "text/csv;charset=utf-8;" }), filename);
};

/* ---------------- Supabase REST/Auth helpers (fetch-based, no SDK) ---------------- */
async function sbFetch(path, { method = "GET", token, body, prefer } = {}) {
  const headers = { apikey: ANON_KEY, "Content-Type": "application/json" };
  headers.Authorization = `Bearer ${token || ANON_KEY}`;
  if (prefer) headers.Prefer = prefer;
  const res = await fetch(`${SUPABASE_URL}${path}`, { method, headers, body: body !== undefined ? JSON.stringify(body) : undefined });
  const text = await res.text();
  let json = null;
  try { json = text ? JSON.parse(text) : null; } catch { json = text; }
  if (!res.ok) {
    const msg = (json && (json.message || json.error_description || json.error || json.msg)) || res.statusText || "Request failed";
    throw new Error(msg);
  }
  return json;
}
const authLogin = (email, password) => sbFetch("/auth/v1/token?grant_type=password", { method: "POST", body: { email, password } });
const restGet = (table, query, token) => sbFetch(`/rest/v1/${table}?${query}`, { token });
const restInsert = (table, row, token) => sbFetch(`/rest/v1/${table}`, { method: "POST", token, body: row, prefer: "return=representation" });
const restUpdate = (table, id, patch, token) => sbFetch(`/rest/v1/${table}?id=eq.${id}`, { method: "PATCH", token, body: patch, prefer: "return=representation" });
const restDelete = (table, id, token) => sbFetch(`/rest/v1/${table}?id=eq.${id}`, { method: "DELETE", token });
const invokeFn = (name, body, token) => sbFetch(`/functions/v1/${name}`, { method: "POST", token, body });

const uid = () => Math.random().toString(36).slice(2, 10);

/* =====================================================================
   APP
===================================================================== */
export default function App() {
  const [session, setSession] = useState(null); // { token, user, employee }
  const [authError, setAuthError] = useState("");
  const [authBusy, setAuthBusy] = useState(false);
  const [tab, setTab] = useState("dashboard");
  const [now, setNow] = useState(new Date());
  const [toast, setToast] = useState(null);

  const [db, setDb] = useState({ stores: [], departments: [], employees: [], shifts: [], leaveRequests: [], lateRequests: [], auditLog: [] });
  const [dbLoading, setDbLoading] = useState(false);
  const pollRef = useRef(null);

  useEffect(() => { const t = setInterval(() => setNow(new Date()), 1000); return () => clearInterval(t); }, []);
  const showToast = (msg, kind = "ok") => { setToast({ msg, kind, id: uid() }); setTimeout(() => setToast((t) => (t && t.msg === msg ? null : t)), 3200); };

  const isManager = !!(session?.employee?.is_manager || session?.employee?.is_admin);

  const loadAll = useCallback(async (token) => {
    setDbLoading(true);
    try {
      const from = addDays(todayISO(), -21);
      const to = addDays(todayISO(), 42);
      const [stores, departments, employees, shifts, leaveRequests, lateRequests] = await Promise.all([
        restGet("stores", "select=*&order=name", token),
        restGet("departments", "select=*&order=name", token),
        restGet("employees", "select=*&order=name", token),
        restGet("shifts", `select=*&shift_date=gte.${from}&shift_date=lte.${to}&order=shift_date`, token),
        restGet("leave_requests", "select=*&order=created_at.desc", token),
        restGet("late_requests", "select=*&order=created_at.desc", token),
      ]);
      let auditLog = [];
      try { auditLog = await restGet("audit_log", "select=*&order=created_at.desc&limit=300", token); } catch { auditLog = []; }
      setDb({ stores, departments, employees, shifts, leaveRequests, lateRequests, auditLog });
    } catch (e) {
      showToast(`Sync error: ${e.message}`, "bad");
    } finally {
      setDbLoading(false);
    }
  }, []);

  useEffect(() => {
    if (!session) { clearInterval(pollRef.current); return; }
    loadAll(session.token);
    pollRef.current = setInterval(() => loadAll(session.token), 45000);
    return () => clearInterval(pollRef.current);
  }, [session, loadAll]);

  const doLogin = async ({ code, pin }) => {
    setAuthError(""); setAuthBusy(true);
    try {
      const email = `${code.trim().toLowerCase()}@${STAFF_DOMAIN}`;
      const tok = await authLogin(email, pin);
      const emps = await restGet("employees", `select=*&auth_user_id=eq.${tok.user.id}`, tok.access_token);
      if (!emps || !emps.length) throw new Error("No employee profile linked to this login yet — ask your manager to provision it.");
      setSession({ token: tok.access_token, user: tok.user, employee: emps[0] });
      setTab("dashboard");
    } catch (e) {
      setAuthError(e.message === "Invalid login credentials" ? "Incorrect employee code or PIN." : e.message);
    } finally {
      setAuthBusy(false);
    }
  };
  const logout = () => { setSession(null); setDb({ stores: [], departments: [], employees: [], shifts: [], leaveRequests: [], lateRequests: [], auditLog: [] }); };

  const logAudit = async (action, details) => {
    try {
      await restInsert("audit_log", { actor_employee_id: session.employee.id, actor_label: session.employee.name, action, details, store_id: session.employee.store_id }, session.token);
    } catch { /* non-fatal */ }
  };

  if (!session) {
    return <LoginScreen onLogin={doLogin} error={authError} busy={authBusy} now={now} />;
  }

  const NAV = [
    { id: "dashboard", label: "Dashboard", icon: Calendar, roles: ["all"] },
    { id: "rota", label: "Rota", icon: Clock, roles: ["all"] },
    { id: "timesheet", label: "Timesheet", icon: FileSpreadsheet, roles: ["all"] },
    { id: "leave", label: "Leave & Requests", icon: ClipboardList, roles: ["all"] },
    { id: "employees", label: "Employees", icon: Users, roles: ["manager"] },
    { id: "stores", label: "Stores", icon: Store, roles: ["manager"] },
    { id: "import", label: "Excel Import", icon: Upload, roles: ["manager"] },
    { id: "audit", label: "Audit Log", icon: History, roles: ["manager"] },
  ];
  const pendingCount = db.leaveRequests.filter((l) => l.status === "Pending").length + db.lateRequests.filter((l) => l.status === "Pending").length;

  return (
    <div className="rw-root min-h-screen">
      <style>{FONT_STYLE}</style>
      <Header now={now} session={session} logout={logout} pendingCount={isManager ? pendingCount : 0} syncing={dbLoading} />
      <div className="max-w-7xl mx-auto px-4 md:px-6">
        <nav className="flex gap-1 overflow-x-auto rw-scrollbar py-3 border-b border-[var(--border)]">
          {NAV.filter((n) => n.roles[0] === "all" || isManager).map((n) => {
            const Icon = n.icon; const active = tab === n.id;
            return (
              <button key={n.id} onClick={() => setTab(n.id)}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-full text-sm whitespace-nowrap transition-colors ${active ? "bg-[var(--amber)] text-[#1A1200] font-semibold" : "text-[var(--muted)] hover:text-[var(--text)] hover:bg-[var(--surface2)]"}`}>
                <Icon size={14} /> {n.label}
                {n.id === "leave" && isManager && pendingCount > 0 && <span className="ml-1 text-[10px] bg-[var(--bad)] text-white rounded-full px-1.5">{pendingCount}</span>}
              </button>
            );
          })}
        </nav>

        <main className="py-6 rw-fade" key={tab}>
          {tab === "dashboard" && <Dashboard db={db} session={session} isManager={isManager} setTab={setTab} />}
          {tab === "rota" && <RotaView db={db} setDb={setDb} session={session} isManager={isManager} logAudit={logAudit} showToast={showToast} reload={() => loadAll(session.token)} />}
          {tab === "timesheet" && <TimesheetView db={db} setDb={setDb} session={session} isManager={isManager} logAudit={logAudit} showToast={showToast} reload={() => loadAll(session.token)} />}
          {tab === "leave" && <LeaveView db={db} setDb={setDb} session={session} isManager={isManager} logAudit={logAudit} showToast={showToast} reload={() => loadAll(session.token)} />}
          {tab === "employees" && isManager && <EmployeesView db={db} session={session} logAudit={logAudit} showToast={showToast} reload={() => loadAll(session.token)} />}
          {tab === "stores" && isManager && <StoresView db={db} session={session} logAudit={logAudit} showToast={showToast} reload={() => loadAll(session.token)} />}
          {tab === "import" && isManager && <ExcelImportView db={db} session={session} logAudit={logAudit} showToast={showToast} reload={() => loadAll(session.token)} />}
          {tab === "audit" && isManager && <AuditView db={db} />}
        </main>
      </div>

      {toast && (
        <div className={`fixed bottom-5 right-5 rw-fade px-4 py-3 rounded-lg shadow-xl text-sm font-medium border z-50 ${toast.kind === "bad" ? "bg-[#2A1414] border-[var(--bad)] text-[#F3B4B0]" : "bg-[#12241C] border-[var(--ok)] text-[#B9E6C9]"}`}>
          {toast.msg}
        </div>
      )}
    </div>
  );
}

/* ---------------- Login ---------------- */
function LoginScreen({ onLogin, error, busy, now }) {
  const [code, setCode] = useState("");
  const [pin, setPin] = useState("");
  return (
    <div className="rw-root min-h-screen flex items-center justify-center p-4">
      <style>{FONT_STYLE}</style>
      <div className="w-full max-w-sm">
        <div className="flex items-center gap-3 mb-6 justify-center">
          <div className="w-10 h-10 rounded-lg bg-[var(--amber)] flex items-center justify-center rw-display font-bold text-[#1A1200]">R</div>
          <div>
            <div className="rw-display font-semibold text-lg leading-tight">Rotawise</div>
            <div className="text-[11px] text-[var(--muted)] leading-tight">Shift, attendance &amp; leave management</div>
          </div>
        </div>
        <div className="bg-[var(--surface)] border border-[var(--border)] rounded-xl p-5">
          <div className="rw-mono text-xs text-[var(--muted)] mb-4 text-center">{now.toLocaleTimeString("en-GB", { hour12: false })} · {now.toLocaleDateString("en-GB", { weekday: "long", day: "2-digit", month: "long" })}</div>
          <form onSubmit={(e) => { e.preventDefault(); onLogin({ code, pin }); }}>
            <Field label="Employee code">
              <input autoFocus required value={code} onChange={(e) => setCode(e.target.value)} className={inputCls} placeholder="e.g. emp1042 or admin" />
            </Field>
            <Field label="PIN">
              <input required type="password" inputMode="numeric" value={pin} onChange={(e) => setPin(e.target.value)} className={inputCls} placeholder="••••••" />
            </Field>
            {error && <div className="text-xs text-[var(--bad)] mb-3">{error}</div>}
            <button type="submit" disabled={busy} className="w-full flex items-center justify-center gap-2 bg-[var(--amber)] text-[#1A1200] font-semibold rounded-lg py-2.5 text-sm disabled:opacity-50">
              {busy ? <Loader2 size={14} className="animate-spin" /> : <ShieldCheck size={14} />} {busy ? "Signing in…" : "Sign in"}
            </button>
          </form>
        </div>
        <div className="text-center text-[11px] text-[var(--muted)] mt-4">Ask your store manager for your employee code and PIN.</div>
      </div>
    </div>
  );
}

/* ---------------- Header ---------------- */
function Header({ now, session, logout, pendingCount, syncing }) {
  const timeStr = now.toLocaleTimeString("en-GB", { hour12: false });
  const dateStr = now.toLocaleDateString("en-GB", { weekday: "long", day: "2-digit", month: "long", year: "numeric" });
  return (
    <header className="border-b border-[var(--border)] bg-[var(--surface)]">
      <div className="max-w-7xl mx-auto px-4 md:px-6 py-4 flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-lg bg-[var(--amber)] flex items-center justify-center rw-display font-bold text-[#1A1200]">R</div>
          <div>
            <div className="rw-display font-semibold text-lg leading-tight">Rotawise</div>
            <div className="text-[11px] text-[var(--muted)] leading-tight">{syncing ? "Syncing…" : "Live"}</div>
          </div>
        </div>
        <div className="flex items-center gap-2 rw-mono text-sm bg-[var(--surface2)] border border-[var(--border)] rounded-lg px-3 py-1.5">
          <Clock size={14} className="text-[var(--amber)]" /><span className="tabular-nums">{timeStr}</span>
          <span className="text-[var(--muted)] hidden sm:inline">· {dateStr}</span>
        </div>
        <div className="flex items-center gap-2">
          {pendingCount > 0 && (
            <div className="flex items-center gap-1 text-xs bg-[#2A2010] border border-[var(--amber)] text-[var(--amber-soft)] rounded-full px-2.5 py-1">
              <Bell size={12} /> {pendingCount} pending
            </div>
          )}
          <div className="flex items-center gap-1.5 text-xs bg-[var(--surface2)] border border-[var(--border)] rounded-full px-2.5 py-1.5">
            {session.employee.is_admin || session.employee.is_manager ? <UserCog size={12} className="text-[var(--amber)]" /> : <User size={12} />}
            {session.employee.name}
          </div>
          <button onClick={logout} className="p-1.5 rounded-lg bg-[var(--surface2)] border border-[var(--border)] text-[var(--muted)] hover:text-[var(--bad)]"><LogOut size={14} /></button>
        </div>
      </div>
    </header>
  );
}

/* ---------------- Shared bits ---------------- */
function Card({ children, className = "" }) { return <div className={`bg-[var(--surface)] border border-[var(--border)] rounded-xl p-4 ${className}`}>{children}</div>; }
function StatCard({ label, value, sub, tone = "text" }) {
  const toneMap = { text: "text-[var(--text)]", ok: "text-[var(--ok)]", bad: "text-[var(--bad)]", amber: "text-[var(--amber)]" };
  return (<Card><div className="text-xs text-[var(--muted)] mb-1">{label}</div><div className={`rw-display text-2xl font-semibold ${toneMap[tone]}`}>{value}</div>{sub && <div className="text-[11px] text-[var(--muted)] mt-1">{sub}</div>}</Card>);
}
function Badge({ children, tone = "muted" }) {
  const toneMap = { muted: "bg-[var(--surface2)] text-[var(--muted)] border-[var(--border)]", ok: "bg-[#12241C] text-[var(--ok)] border-[var(--ok)]", bad: "bg-[#2A1414] text-[var(--bad)] border-[var(--bad)]", amber: "bg-[#2A2010] text-[var(--amber-soft)] border-[var(--amber)]", info: "bg-[#131F2A] text-[var(--info)] border-[var(--info)]" };
  return <span className={`text-[10px] font-semibold uppercase tracking-wide border rounded-full px-2 py-0.5 ${toneMap[tone]}`}>{children}</span>;
}
function Btn({ children, onClick, tone = "default", size = "md", icon: Icon, disabled, type = "button" }) {
  const base = "inline-flex items-center gap-1.5 rounded-lg font-medium transition-colors disabled:opacity-40 disabled:cursor-not-allowed";
  const sizes = { sm: "text-xs px-2.5 py-1.5", md: "text-sm px-3.5 py-2" };
  const tones = { default: "bg-[var(--amber)] text-[#1A1200] hover:bg-[var(--amber-soft)]", ghost: "bg-[var(--surface2)] text-[var(--text)] border border-[var(--border)] hover:border-[var(--amber)]", danger: "bg-[#2A1414] text-[#F3B4B0] border border-[var(--bad)] hover:bg-[var(--bad)] hover:text-white", ok: "bg-[#12241C] text-[var(--ok)] border border-[var(--ok)] hover:bg-[var(--ok)] hover:text-white" };
  return <button type={type} onClick={onClick} disabled={disabled} className={`${base} ${sizes[size]} ${tones[tone]}`}>{Icon && <Icon size={size === "sm" ? 12 : 14} />} {children}</button>;
}
function Modal({ title, onClose, children, wide }) {
  return (
    <div className="fixed inset-0 bg-black/60 flex items-center justify-center p-4 z-50" onClick={onClose}>
      <div className={`bg-[var(--surface)] border border-[var(--border)] rounded-xl p-5 w-full ${wide ? "max-w-2xl" : "max-w-md"} max-h-[85vh] overflow-y-auto rw-scrollbar rw-fade`} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between mb-4"><h3 className="rw-display font-semibold text-base">{title}</h3><button onClick={onClose} className="text-[var(--muted)] hover:text-[var(--text)]"><X size={18} /></button></div>
        {children}
      </div>
    </div>
  );
}
function Field({ label, children }) { return <label className="block mb-3"><div className="text-xs text-[var(--muted)] mb-1">{label}</div>{children}</label>; }
const inputCls = "w-full bg-[var(--surface2)] border border-[var(--border)] rounded-lg px-3 py-2 text-sm outline-none focus:border-[var(--amber)]";
function ExportBar({ onXlsx, onCsv, onPrint, onWhatsapp }) {
  return (
    <div className="flex gap-1.5 flex-wrap">
      <Btn size="sm" tone="ghost" icon={FileSpreadsheet} onClick={onXlsx}>Excel</Btn>
      <Btn size="sm" tone="ghost" icon={FileDown} onClick={onCsv}>CSV</Btn>
      {onPrint && <Btn size="sm" tone="ghost" icon={Printer} onClick={onPrint}>Print / PDF</Btn>}
      {onWhatsapp && <Btn size="sm" tone="ghost" icon={MessageCircle} onClick={onWhatsapp}>WhatsApp</Btn>}
    </div>
  );
}
const openWhatsapp = (text) => window.open(`https://wa.me/?text=${encodeURIComponent(text)}`, "_blank");

/* =====================================================================
   DASHBOARD
===================================================================== */
function Dashboard({ db, session, isManager, setTab }) {
  const today = todayISO();
  const empById = Object.fromEntries(db.employees.map((e) => [e.id, e]));
  const todaysShifts = db.shifts.filter((s) => s.shift_date === today);

  if (!isManager) {
    const me = session.employee;
    const myShifts = db.shifts.filter((s) => s.employee_id === me.id).sort((a, b) => a.shift_date.localeCompare(b.shift_date));
    const next = myShifts.find((s) => s.shift_date >= today);
    const weekStart = startOfWeek(today);
    const weekShifts = myShifts.filter((s) => s.shift_date >= weekStart && s.shift_date < addDays(weekStart, 7));
    const scheduledHrs = weekShifts.reduce((a, s) => a + hoursBetween(timeShort(s.start_time), timeShort(s.end_time), s.break_mins), 0);
    const holidayRemaining = (me.holiday_allowance || 0) - (me.holiday_used || 0);
    const pendingLeave = db.leaveRequests.filter((l) => l.employee_id === me.id && l.status === "Pending").length;
    return (
      <div className="grid md:grid-cols-3 gap-4">
        <Card className="md:col-span-2">
          <div className="text-xs text-[var(--muted)] mb-2 flex items-center gap-1"><Sunrise size={12} /> My next shift</div>
          {next ? (
            <div className="flex items-center justify-between">
              <div>
                <div className="rw-display text-xl font-semibold">{fmtDate(next.shift_date)}</div>
                <div className="rw-mono text-[var(--amber)] text-lg">{timeShort(next.start_time)} – {timeShort(next.end_time)}</div>
              </div>
              <div className="text-right text-sm text-[var(--muted)]">
                <div>{db.stores.find((s) => s.id === me.store_id)?.name}</div>
                <div>{db.departments.find((d) => d.id === me.department_id)?.name}</div>
              </div>
            </div>
          ) : <div className="text-sm text-[var(--muted)]">No upcoming shifts scheduled.</div>}
        </Card>
        <StatCard label="This week — scheduled vs contracted" value={`${scheduledHrs}h / ${me.contracted_hours}h`} tone={scheduledHrs < me.contracted_hours ? "amber" : "ok"} />
        <StatCard label="Holiday remaining" value={`${holidayRemaining} days`} sub={`${me.holiday_used || 0} used of ${me.holiday_allowance}`} />
        <StatCard label="Pending requests" value={pendingLeave} tone={pendingLeave ? "amber" : "text"} />
        <Card className="md:col-span-3">
          <div className="text-xs text-[var(--muted)] mb-3">Upcoming shifts</div>
          <div className="space-y-2">
            {myShifts.filter((s) => s.shift_date >= today).slice(0, 6).map((s) => <ShiftChip key={s.id} shift={s} />)}
            {myShifts.filter((s) => s.shift_date >= today).length === 0 && <div className="text-sm text-[var(--muted)]">Nothing scheduled yet.</div>}
          </div>
        </Card>
      </div>
    );
  }

  const scheduledToday = todaysShifts.length;
  const lateToday = todaysShifts.filter((s) => s.actual_start && s.actual_start > s.start_time).length;
  const onHolidayToday = db.leaveRequests.filter((l) => l.status === "Approved" && l.start_date <= today && l.end_date >= today).length;
  const pendingApprovals = db.leaveRequests.filter((l) => l.status === "Pending").length + db.lateRequests.filter((l) => l.status === "Pending").length;
  const weekStart = startOfWeek(today);
  const weekShifts = db.shifts.filter((s) => s.shift_date >= weekStart && s.shift_date < addDays(weekStart, 7));
  const byEmp = {};
  db.employees.forEach((e) => { byEmp[e.id] = { scheduled: 0, contracted: e.contracted_hours, name: e.name }; });
  weekShifts.forEach((s) => { if (byEmp[s.employee_id]) byEmp[s.employee_id].scheduled += hoursBetween(timeShort(s.start_time), timeShort(s.end_time), s.break_mins); });

  return (
    <div className="space-y-4">
      <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-4">
        <StatCard label="Scheduled today" value={scheduledToday} />
        <StatCard label="Late arrivals today" value={lateToday} tone={lateToday ? "bad" : "ok"} />
        <StatCard label="On approved leave" value={onHolidayToday} />
        <StatCard label="Pending approvals" value={pendingApprovals} tone={pendingApprovals ? "amber" : "ok"} />
      </div>
      <Card>
        <div className="text-sm font-semibold mb-3">This week — contracted vs scheduled hours</div>
        <div className="space-y-2">
          {Object.values(byEmp).map((e) => {
            const diff = Math.round((e.scheduled - e.contracted) * 100) / 100;
            const pct = Math.min(100, (e.scheduled / (e.contracted || 1)) * 100);
            return (
              <div key={e.name}>
                <div className="flex justify-between text-xs mb-1"><span>{e.name}</span><span className="rw-mono text-[var(--muted)]">{e.scheduled}h / {e.contracted}h {diff !== 0 && <span className={diff < 0 ? "text-[var(--amber)]" : "text-[var(--info)]"}>({diff > 0 ? "+" : ""}{diff}h)</span>}</span></div>
                <div className="h-1.5 rounded-full bg-[var(--surface2)] overflow-hidden"><div className="h-full bg-[var(--amber)]" style={{ width: `${pct}%` }} /></div>
              </div>
            );
          })}
          {Object.values(byEmp).length === 0 && <div className="text-sm text-[var(--muted)]">No employees at your store yet.</div>}
        </div>
      </Card>
      <div className="grid md:grid-cols-2 gap-4">
        <Card>
          <div className="text-sm font-semibold mb-3">Today's rota</div>
          <div className="space-y-2">
            {todaysShifts.length === 0 && <div className="text-sm text-[var(--muted)]">No shifts scheduled today.</div>}
            {todaysShifts.map((s) => (<div key={s.id} className="flex items-center justify-between text-sm"><span>{empById[s.employee_id]?.name}</span><span className="rw-mono text-[var(--muted)]">{timeShort(s.start_time)}–{timeShort(s.end_time)}</span></div>))}
          </div>
        </Card>
        <Card>
          <div className="text-sm font-semibold mb-3">Pending actions</div>
          <div className="space-y-2">
            {db.leaveRequests.filter((l) => l.status === "Pending").map((l) => (<div key={l.id} className="flex items-center justify-between text-sm"><span>{empById[l.employee_id]?.name} — {l.type}</span><Badge tone="amber">Pending</Badge></div>))}
            {db.lateRequests.filter((l) => l.status === "Pending").map((l) => (<div key={l.id} className="flex items-center justify-between text-sm"><span>{empById[l.employee_id]?.name} — Late arrival</span><Badge tone="amber">Pending</Badge></div>))}
            {pendingApprovals === 0 && <div className="text-sm text-[var(--muted)]">All caught up.</div>}
            {pendingApprovals > 0 && <Btn size="sm" tone="ghost" onClick={() => setTab("leave")}>Review all →</Btn>}
          </div>
        </Card>
      </div>
    </div>
  );
}
function ShiftChip({ shift }) {
  const hrs = hoursBetween(timeShort(shift.start_time), timeShort(shift.end_time), shift.break_mins);
  return (
    <div className="rw-punch flex items-center justify-between bg-[var(--surface2)] border border-[var(--border)] rounded-lg pl-4 pr-3 py-2 text-sm">
      <div className="flex items-center gap-3"><div className="rw-hole hidden sm:block" /><div><div className="font-medium">{fmtDate(shift.shift_date)}</div><div className="rw-mono text-[var(--muted)] text-xs">{timeShort(shift.start_time)} – {timeShort(shift.end_time)} · {hrs}h paid</div></div></div>
      <Badge tone={shift.status === "worked" ? "ok" : shift.status === "changed" ? "info" : "muted"}>{shift.status}</Badge>
    </div>
  );
}

/* =====================================================================
   ROTA
===================================================================== */
function RotaView({ db, setDb, session, isManager, logAudit, showToast, reload }) {
  const [weekStart, setWeekStart] = useState(startOfWeek(todayISO()));
  const [modal, setModal] = useState(null);
  const days = Array.from({ length: 7 }, (_, i) => addDays(weekStart, i));
  const employees = isManager ? db.employees : db.employees.filter((e) => e.id === session.employee.id);
  const empById = Object.fromEntries(db.employees.map((e) => [e.id, e]));
  const shiftFor = (empId, date) => db.shifts.find((s) => s.employee_id === empId && s.shift_date === date);

  const saveShift = async (form) => {
    try {
      if (form.id) {
        const before = db.shifts.find((s) => s.id === form.id);
        const patch = { shift_date: form.date, start_time: form.start, end_time: form.end, break_mins: form.breakMins, notes: form.notes || "", status: before.status === "worked" ? "changed" : "scheduled" };
        const [updated] = await restUpdate("shifts", form.id, patch, session.token);
        setDb((d) => ({ ...d, shifts: d.shifts.map((s) => (s.id === form.id ? updated : s)) }));
        logAudit("Edited shift", `${form.employeeName} ${form.date} ${timeShort(before.start_time)}-${timeShort(before.end_time)} → ${form.start}-${form.end}`);
      } else {
        const row = { employee_id: form.employeeId, store_id: empById[form.employeeId]?.store_id, shift_date: form.date, start_time: form.start, end_time: form.end, break_mins: form.breakMins, status: "scheduled", notes: form.notes || "", created_by: session.employee.id, source: "manual" };
        const [created] = await restInsert("shifts", row, session.token);
        setDb((d) => ({ ...d, shifts: [...d.shifts, created] }));
        logAudit("Created shift", `${form.employeeName} ${form.date} ${form.start}-${form.end}`);
      }
      setModal(null); showToast("Shift saved");
    } catch (e) { showToast(e.message, "bad"); }
  };
  const deleteShift = async (shift) => {
    try {
      await restDelete("shifts", shift.id, session.token);
      setDb((d) => ({ ...d, shifts: d.shifts.filter((s) => s.id !== shift.id) }));
      logAudit("Deleted shift", `${shift.shift_date} ${timeShort(shift.start_time)}-${timeShort(shift.end_time)}`);
      showToast("Shift removed");
    } catch (e) { showToast(e.message, "bad"); }
  };

  const buildRotaRows = () => db.shifts.filter((s) => days.includes(s.shift_date) && employees.some((e) => e.id === s.employee_id)).map((s) => ({
    Employee: empById[s.employee_id]?.name || "", Date: s.shift_date, Start: timeShort(s.start_time), End: timeShort(s.end_time), "Break (mins)": s.break_mins, "Paid Hours": hoursBetween(timeShort(s.start_time), timeShort(s.end_time), s.break_mins), Status: s.status,
  }));
  const whatsappText = () => {
    let text = `*Rota — ${fmtDateShort(weekStart)} to ${fmtDateShort(addDays(weekStart, 6))}*\n\n`;
    days.forEach((d) => {
      const dayShifts = db.shifts.filter((s) => s.shift_date === d && employees.some((e) => e.id === s.employee_id));
      if (dayShifts.length === 0) return;
      text += `*${fmtDate(d)}*\n`;
      dayShifts.forEach((s) => { text += `${empById[s.employee_id]?.name}: ${timeShort(s.start_time)}–${timeShort(s.end_time)}\n`; });
      text += `\n`;
    });
    return text;
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
        <div className="flex items-center gap-2">
          <button onClick={() => setWeekStart(addDays(weekStart, -7))} className="p-1.5 rounded-lg bg-[var(--surface2)] border border-[var(--border)]"><ChevronLeft size={16} /></button>
          <div className="rw-display font-semibold">{fmtDateShort(weekStart)} – {fmtDateShort(addDays(weekStart, 6))}</div>
          <button onClick={() => setWeekStart(addDays(weekStart, 7))} className="p-1.5 rounded-lg bg-[var(--surface2)] border border-[var(--border)]"><ChevronRight size={16} /></button>
          <Btn size="sm" tone="ghost" onClick={() => setWeekStart(startOfWeek(todayISO()))}>This week</Btn>
        </div>
        <ExportBar
          onXlsx={() => exportWorkbook([{ name: "Rota", rows: buildRotaRows() }], `rota_${weekStart}.xlsx`)}
          onCsv={() => exportCSV(buildRotaRows(), `rota_${weekStart}.csv`)}
          onPrint={() => window.print()}
          onWhatsapp={() => openWhatsapp(whatsappText())}
        />
      </div>

      <div className="overflow-x-auto rw-scrollbar">
        <div style={{ minWidth: `${140 + days.length * 150}px` }}>
          <div className="grid" style={{ gridTemplateColumns: `140px repeat(${days.length}, 1fr)` }}>
            <div />
            {days.map((d) => (<div key={d} className={`text-center text-xs pb-2 font-medium ${d === todayISO() ? "text-[var(--amber)]" : "text-[var(--muted)]"}`}>{fmtDate(d)}</div>))}
            {employees.map((emp) => (
              <React.Fragment key={emp.id}>
                <div className="py-2 pr-2 text-sm font-medium flex items-center">{emp.name}</div>
                {days.map((d) => {
                  const shift = shiftFor(emp.id, d);
                  return (
                    <div key={d} className="p-1">
                      {shift ? (
                        <div className="rw-punch group relative bg-[var(--surface2)] border border-[var(--border)] rounded-lg pl-3 pr-2 py-2 text-xs">
                          <div className="rw-mono font-medium">{timeShort(shift.start_time)}–{timeShort(shift.end_time)}</div>
                          <div className="text-[var(--muted)]">{hoursBetween(timeShort(shift.start_time), timeShort(shift.end_time), shift.break_mins)}h paid</div>
                          {isManager && (
                            <div className="hidden group-hover:flex gap-1 absolute top-1 right-1">
                              <button onClick={() => setModal({ id: shift.id, date: shift.shift_date, start: timeShort(shift.start_time), end: timeShort(shift.end_time), breakMins: shift.break_mins, notes: shift.notes, employeeName: emp.name })} className="p-1 bg-[var(--bg)] rounded"><Edit3 size={11} /></button>
                              <button onClick={() => deleteShift(shift)} className="p-1 bg-[var(--bg)] rounded text-[var(--bad)]"><Trash2 size={11} /></button>
                            </div>
                          )}
                        </div>
                      ) : isManager ? (
                        <button onClick={() => setModal({ employeeId: emp.id, employeeName: emp.name, date: d, start: "09:00", end: "17:00", breakMins: 30 })} className="w-full h-full min-h-[52px] border border-dashed border-[var(--border)] rounded-lg text-[var(--muted)] hover:border-[var(--amber)] hover:text-[var(--amber)] flex items-center justify-center"><Plus size={14} /></button>
                      ) : (<div className="min-h-[52px] rounded-lg" />)}
                    </div>
                  );
                })}
              </React.Fragment>
            ))}
            {employees.length === 0 && <div className="col-span-full text-sm text-[var(--muted)] py-6">No employees to show.</div>}
          </div>
        </div>
      </div>

      {isManager && employees.length > 0 && <RotaWarnings employees={employees} shifts={db.shifts} days={days} />}

      {modal && (
        <Modal title={modal.id ? "Edit shift" : "Add shift"} onClose={() => setModal(null)}>
          <ShiftForm initial={modal} onSave={saveShift} onCancel={() => setModal(null)} />
        </Modal>
      )}
    </div>
  );
}
function RotaWarnings({ employees, shifts, days }) {
  const weekShifts = shifts.filter((s) => days.includes(s.shift_date));
  const byEmp = {};
  employees.forEach((e) => { byEmp[e.id] = { name: e.name, scheduled: 0, contracted: e.contracted_hours }; });
  weekShifts.forEach((s) => { if (byEmp[s.employee_id]) byEmp[s.employee_id].scheduled += hoursBetween(timeShort(s.start_time), timeShort(s.end_time), s.break_mins); });
  const warnings = [];
  Object.values(byEmp).forEach((e) => {
    if (e.scheduled < e.contracted) warnings.push({ text: `${e.name} is contracted for ${e.contracted}h but scheduled for only ${e.scheduled}h this week (${Math.round((e.contracted - e.scheduled) * 100) / 100}h short).`, tone: "amber" });
    if (e.scheduled > e.contracted * 1.15 && e.contracted > 0) warnings.push({ text: `${e.name} is scheduled well over contracted hours (${e.scheduled}h vs ${e.contracted}h).`, tone: "bad" });
  });
  const sorted = [...weekShifts].sort((a, b) => (a.employee_id + a.shift_date).localeCompare(b.employee_id + b.shift_date));
  for (let i = 0; i < sorted.length - 1; i++) {
    const a = sorted[i], b = sorted[i + 1];
    if (a.employee_id === b.employee_id && addDays(a.shift_date, 1) === b.shift_date) {
      const restHrs = (24 - parseInt(a.end_time)) + parseInt(b.start_time);
      if (restHrs < 11) warnings.push({ text: `${byEmp[a.employee_id]?.name} has less than 11 hours rest between shifts on ${a.shift_date} and ${b.shift_date}.`, tone: "bad" });
    }
  }
  if (warnings.length === 0) return null;
  return (
    <Card className="mt-4">
      <div className="text-sm font-semibold mb-2 flex items-center gap-1.5"><AlertTriangle size={14} className="text-[var(--amber)]" /> Rota warnings</div>
      <div className="space-y-1.5">{warnings.map((w, i) => (<div key={i} className={`text-xs flex items-start gap-1.5 ${w.tone === "bad" ? "text-[#F3B4B0]" : "text-[var(--amber-soft)]"}`}><span>•</span><span>{w.text}</span></div>))}</div>
    </Card>
  );
}
function ShiftForm({ initial, onSave, onCancel }) {
  const [form, setForm] = useState({ breakMins: 30, ...initial });
  const [saving, setSaving] = useState(false);
  return (
    <form onSubmit={async (e) => { e.preventDefault(); setSaving(true); await onSave(form); setSaving(false); }}>
      <Field label="Date"><input type="date" required value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} className={inputCls} /></Field>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Start time"><input type="time" required value={form.start} onChange={(e) => setForm({ ...form, start: e.target.value })} className={inputCls} /></Field>
        <Field label="End time"><input type="time" required value={form.end} onChange={(e) => setForm({ ...form, end: e.target.value })} className={inputCls} /></Field>
      </div>
      <Field label="Break (minutes)"><input type="number" min="0" value={form.breakMins} onChange={(e) => setForm({ ...form, breakMins: Number(e.target.value) })} className={inputCls} /></Field>
      <Field label="Notes (optional)"><input value={form.notes || ""} onChange={(e) => setForm({ ...form, notes: e.target.value })} className={inputCls} placeholder="Cover for..." /></Field>
      <div className="text-xs text-[var(--muted)] mb-4">Paid hours: <span className="rw-mono text-[var(--text)]">{hoursBetween(form.start, form.end, form.breakMins)}h</span></div>
      <div className="flex justify-end gap-2"><Btn tone="ghost" onClick={onCancel}>Cancel</Btn><Btn type="submit" icon={Check} disabled={saving}>{saving ? "Saving…" : "Save shift"}</Btn></div>
    </form>
  );
}

/* =====================================================================
   TIMESHEET
===================================================================== */
function TimesheetView({ db, setDb, session, isManager, logAudit, showToast }) {
  const [range, setRange] = useState({ from: addDays(todayISO(), -13), to: todayISO() });
  const [correcting, setCorrecting] = useState(null);
  const [lateModal, setLateModal] = useState(null);
  const empById = Object.fromEntries(db.employees.map((e) => [e.id, e]));

  let rows = db.shifts.filter((s) => s.shift_date >= range.from && s.shift_date <= range.to);
  if (!isManager) rows = rows.filter((s) => s.employee_id === session.employee.id);
  rows = [...rows].sort((a, b) => b.shift_date.localeCompare(a.shift_date));

  const saveCorrection = async (form) => {
    try {
      const before = db.shifts.find((s) => s.id === form.id);
      const [updated] = await restUpdate("shifts", form.id, { actual_start: form.actualStart, actual_end: form.actualEnd, status: "worked" }, session.token);
      setDb((d) => ({ ...d, shifts: d.shifts.map((s) => (s.id === form.id ? updated : s)) }));
      logAudit("Attendance correction", `${empById[before.employee_id]?.name} ${before.shift_date}: ${timeShort(before.actual_start) || "—"}/${timeShort(before.actual_end) || "—"} → ${form.actualStart}/${form.actualEnd}. Reason: ${form.reason || "n/a"}`);
      setCorrecting(null); showToast("Attendance updated");
    } catch (e) { showToast(e.message, "bad"); }
  };
  const submitLate = async (form) => {
    try {
      const row = { employee_id: session.employee.id, shift_id: form.shiftId, expected_arrival: form.expectedArrival, reason: form.reason, status: "Pending" };
      const [created] = await restInsert("late_requests", row, session.token);
      setDb((d) => ({ ...d, lateRequests: [created, ...d.lateRequests] }));
      logAudit("Late arrival reported", `Expected ${form.expectedArrival}. Reason: ${form.reason}`);
      setLateModal(null); showToast("Manager notified of your late arrival");
    } catch (e) { showToast(e.message, "bad"); }
  };

  const totals = rows.reduce((acc, s) => {
    acc.scheduled += hoursBetween(timeShort(s.start_time), timeShort(s.end_time), s.break_mins);
    if (s.actual_start && s.actual_end) acc.actual += hoursBetween(timeShort(s.actual_start), timeShort(s.actual_end), s.break_mins);
    return acc;
  }, { scheduled: 0, actual: 0 });

  const exportRows = () => rows.map((s) => ({
    Employee: empById[s.employee_id]?.name || "", Date: s.shift_date, "Scheduled Start": timeShort(s.start_time), "Scheduled End": timeShort(s.end_time),
    "Actual Start": timeShort(s.actual_start) || "", "Actual End": timeShort(s.actual_end) || "",
    "Scheduled Hours": hoursBetween(timeShort(s.start_time), timeShort(s.end_time), s.break_mins),
    "Actual Hours": s.actual_start && s.actual_end ? hoursBetween(timeShort(s.actual_start), timeShort(s.actual_end), s.break_mins) : "",
    "Variance (mins)": varianceMins(s), Status: s.status,
  }));

  return (
    <div>
      <div className="flex flex-wrap items-end gap-3 mb-4">
        <Field label="From"><input type="date" value={range.from} onChange={(e) => setRange({ ...range, from: e.target.value })} className={inputCls} /></Field>
        <Field label="To"><input type="date" value={range.to} onChange={(e) => setRange({ ...range, to: e.target.value })} className={inputCls} /></Field>
        <div className="flex-1" />
        <ExportBar onXlsx={() => exportWorkbook([{ name: "Timesheet", rows: exportRows() }], `timesheet_${range.from}_to_${range.to}.xlsx`)} onCsv={() => exportCSV(exportRows(), `timesheet_${range.from}_to_${range.to}.csv`)} onPrint={() => window.print()} />
      </div>
      <div className="grid sm:grid-cols-3 gap-4 mb-4">
        <StatCard label="Scheduled hours" value={`${Math.round(totals.scheduled * 100) / 100}h`} />
        <StatCard label="Actual hours" value={`${Math.round(totals.actual * 100) / 100}h`} tone="ok" />
        <StatCard label="Variance" value={`${Math.round((totals.actual - totals.scheduled) * 100) / 100}h`} tone={totals.actual < totals.scheduled ? "amber" : "text"} />
      </div>
      <Card className="overflow-x-auto rw-scrollbar">
        <table className="w-full text-sm min-w-[760px]">
          <thead><tr className="text-left text-xs text-[var(--muted)] border-b border-[var(--border)]">
            {isManager && <th className="pb-2 pr-3">Employee</th>}<th className="pb-2 pr-3">Date</th><th className="pb-2 pr-3">Scheduled</th><th className="pb-2 pr-3">Actual</th><th className="pb-2 pr-3">Variance</th><th className="pb-2 pr-3">Status</th><th className="pb-2 pr-3"></th>
          </tr></thead>
          <tbody>
            {rows.map((s) => {
              const varMins = varianceMins(s);
              return (
                <tr key={s.id} className="border-b border-[var(--border)]/50">
                  {isManager && <td className="py-2 pr-3">{empById[s.employee_id]?.name}</td>}
                  <td className="py-2 pr-3">{fmtDate(s.shift_date)}</td>
                  <td className="py-2 pr-3 rw-mono">{timeShort(s.start_time)}–{timeShort(s.end_time)}</td>
                  <td className="py-2 pr-3 rw-mono">{s.actual_start ? `${timeShort(s.actual_start)}–${timeShort(s.actual_end)}` : "—"}</td>
                  <td className={`py-2 pr-3 rw-mono ${varMins == null ? "text-[var(--muted)]" : varMins > 5 ? "text-[var(--bad)]" : "text-[var(--ok)]"}`}>{varMins == null ? "—" : `${varMins > 0 ? "+" : ""}${varMins}m`}</td>
                  <td className="py-2 pr-3"><Badge tone={s.status === "worked" ? "ok" : s.status === "changed" ? "info" : "muted"}>{s.status}</Badge></td>
                  <td className="py-2 pr-3 flex gap-1.5">
                    {isManager && <Btn size="sm" tone="ghost" icon={Edit3} onClick={() => setCorrecting(s)}>Correct</Btn>}
                    {!isManager && s.shift_date >= todayISO() && <Btn size="sm" tone="ghost" icon={AlertTriangle} onClick={() => setLateModal(s)}>Report late</Btn>}
                  </td>
                </tr>
              );
            })}
            {rows.length === 0 && <tr><td colSpan={7} className="py-6 text-center text-[var(--muted)] text-sm">No shifts in this range.</td></tr>}
          </tbody>
        </table>
      </Card>
      {correcting && <Modal title="Correct attendance record" onClose={() => setCorrecting(null)}><CorrectionForm shift={correcting} onSave={saveCorrection} onCancel={() => setCorrecting(null)} /></Modal>}
      {lateModal && <Modal title="Report late arrival" onClose={() => setLateModal(null)}><LateForm shift={lateModal} onSave={submitLate} onCancel={() => setLateModal(null)} /></Modal>}
    </div>
  );
}
function varianceMins(s) {
  if (!s.actual_start || !s.actual_end) return null;
  const toMin = (t) => { const [h, m] = timeShort(t).split(":").map(Number); return h * 60 + m; };
  return toMin(s.actual_start) - toMin(s.start_time);
}
function CorrectionForm({ shift, onSave, onCancel }) {
  const [form, setForm] = useState({ id: shift.id, actualStart: timeShort(shift.actual_start) || timeShort(shift.start_time), actualEnd: timeShort(shift.actual_end) || timeShort(shift.end_time), reason: "" });
  const [saving, setSaving] = useState(false);
  return (
    <form onSubmit={async (e) => { e.preventDefault(); setSaving(true); await onSave(form); setSaving(false); }}>
      <div className="text-xs text-[var(--muted)] mb-3">Scheduled: <span className="rw-mono text-[var(--text)]">{timeShort(shift.start_time)}–{timeShort(shift.end_time)}</span></div>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Actual start"><input type="time" required value={form.actualStart} onChange={(e) => setForm({ ...form, actualStart: e.target.value })} className={inputCls} /></Field>
        <Field label="Actual end"><input type="time" required value={form.actualEnd} onChange={(e) => setForm({ ...form, actualEnd: e.target.value })} className={inputCls} /></Field>
      </div>
      <Field label="Reason for correction"><input required value={form.reason} onChange={(e) => setForm({ ...form, reason: e.target.value })} className={inputCls} placeholder="e.g. clock system fault" /></Field>
      <div className="text-[11px] text-[var(--muted)] mb-4">This correction is recorded in the audit log with your name and the original values.</div>
      <div className="flex justify-end gap-2"><Btn tone="ghost" onClick={onCancel}>Cancel</Btn><Btn type="submit" icon={Check} disabled={saving}>{saving ? "Saving…" : "Save correction"}</Btn></div>
    </form>
  );
}
function LateForm({ shift, onSave, onCancel }) {
  const [form, setForm] = useState({ shiftId: shift.id, expectedArrival: timeShort(shift.start_time), reason: "" });
  const [saving, setSaving] = useState(false);
  return (
    <form onSubmit={async (e) => { e.preventDefault(); setSaving(true); await onSave(form); setSaving(false); }}>
      <div className="text-xs text-[var(--muted)] mb-3">Scheduled start: <span className="rw-mono text-[var(--text)]">{timeShort(shift.start_time)}</span> on {fmtDate(shift.shift_date)}</div>
      <Field label="Expected arrival time"><input type="time" required value={form.expectedArrival} onChange={(e) => setForm({ ...form, expectedArrival: e.target.value })} className={inputCls} /></Field>
      <Field label="Reason"><input required value={form.reason} onChange={(e) => setForm({ ...form, reason: e.target.value })} className={inputCls} placeholder="e.g. transport delay" /></Field>
      <div className="flex justify-end gap-2"><Btn tone="ghost" onClick={onCancel}>Cancel</Btn><Btn type="submit" icon={Check} disabled={saving}>{saving ? "Sending…" : "Notify manager"}</Btn></div>
    </form>
  );
}

/* =====================================================================
   LEAVE & REQUESTS
===================================================================== */
function LeaveView({ db, setDb, session, isManager, logAudit, showToast }) {
  const [requesting, setRequesting] = useState(false);
  const empById = Object.fromEntries(db.employees.map((e) => [e.id, e]));
  const me = session.employee;

  const submitLeave = async (form) => {
    try {
      const days = Math.max(1, Math.round((new Date(form.endDate) - new Date(form.startDate)) / 86400000) + 1);
      const row = { employee_id: me.id, type: form.type, start_date: form.startDate, end_date: form.endDate, days, status: "Pending", comment: form.comment };
      const [created] = await restInsert("leave_requests", row, session.token);
      setDb((d) => ({ ...d, leaveRequests: [created, ...d.leaveRequests] }));
      logAudit("Leave requested", `${form.type} ${form.startDate} to ${form.endDate}`);
      setRequesting(false); showToast("Request sent to your manager");
    } catch (e) { showToast(e.message, "bad"); }
  };
  const decide = async (req, decision) => {
    try {
      const [updated] = await restUpdate("leave_requests", req.id, { status: decision, decided_by: me.id, decided_at: new Date().toISOString() }, session.token);
      setDb((d) => ({ ...d, leaveRequests: d.leaveRequests.map((l) => (l.id === req.id ? updated : l)) }));
      if (decision === "Approved" && req.type === "Holiday") {
        const emp = empById[req.employee_id];
        const [updatedEmp] = await restUpdate("employees", req.employee_id, { holiday_used: Math.round(((emp.holiday_used || 0) + req.days) * 100) / 100 }, session.token);
        setDb((d) => ({ ...d, employees: d.employees.map((e) => (e.id === req.employee_id ? updatedEmp : e)) }));
      }
      logAudit(`Leave ${decision.toLowerCase()}`, `${empById[req.employee_id]?.name} — ${req.type} ${req.start_date} to ${req.end_date}`);
      showToast(`Request ${decision.toLowerCase()}`);
    } catch (e) { showToast(e.message, "bad"); }
  };
  const decideLate = async (req, decision) => {
    try {
      const [updated] = await restUpdate("late_requests", req.id, { status: decision, decided_by: me.id }, session.token);
      setDb((d) => ({ ...d, lateRequests: d.lateRequests.map((l) => (l.id === req.id ? updated : l)) }));
      if (decision === "Approved" && req.shift_id) {
        const [updatedShift] = await restUpdate("shifts", req.shift_id, { actual_start: req.expected_arrival, status: "changed" }, session.token);
        setDb((d) => ({ ...d, shifts: d.shifts.map((s) => (s.id === req.shift_id ? updatedShift : s)) }));
      }
      logAudit(`Late arrival ${decision.toLowerCase()}`, `${empById[req.employee_id]?.name}`);
      showToast(`Late arrival ${decision.toLowerCase()}`);
    } catch (e) { showToast(e.message, "bad"); }
  };

  const myLeave = isManager ? db.leaveRequests : db.leaveRequests.filter((l) => l.employee_id === me.id);
  const sortedLeave = [...myLeave].sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
  const pendingLate = db.lateRequests.filter((l) => l.status === "Pending" && (isManager || l.employee_id === me.id));

  return (
    <div className="space-y-4">
      {!isManager && (
        <Card>
          <div className="flex items-center justify-between">
            <div><div className="text-sm font-semibold">Holiday balance</div><div className="text-xs text-[var(--muted)]">{me.holiday_used || 0} used · {(me.holiday_allowance || 0) - (me.holiday_used || 0)} remaining of {me.holiday_allowance}</div></div>
            <Btn icon={Plus} onClick={() => setRequesting(true)}>Request leave</Btn>
          </div>
        </Card>
      )}
      {isManager && pendingLate.length > 0 && (
        <Card>
          <div className="text-sm font-semibold mb-3">Late arrival reports</div>
          <div className="space-y-2">
            {pendingLate.map((l) => (
              <div key={l.id} className="flex items-center justify-between text-sm bg-[var(--surface2)] rounded-lg px-3 py-2">
                <div><div className="font-medium">{empById[l.employee_id]?.name}</div><div className="text-xs text-[var(--muted)]">Expected {timeShort(l.expected_arrival)} · {l.reason}</div></div>
                <div className="flex gap-1.5"><Btn size="sm" tone="ok" icon={Check} onClick={() => decideLate(l, "Approved")}>Approve</Btn><Btn size="sm" tone="danger" icon={X} onClick={() => decideLate(l, "Rejected")}>Reject</Btn></div>
              </div>
            ))}
          </div>
        </Card>
      )}
      <Card className="overflow-x-auto rw-scrollbar">
        <div className="flex items-center justify-between mb-3">
          <div className="text-sm font-semibold">{isManager ? "All leave requests" : "My leave history"}</div>
          {isManager && <ExportBar onXlsx={() => exportWorkbook([{ name: "Leave", rows: sortedLeave.map((l) => ({ Employee: empById[l.employee_id]?.name, Type: l.type, Start: l.start_date, End: l.end_date, Days: l.days, Status: l.status, Comment: l.comment })) }], "leave_requests.xlsx")} onCsv={() => exportCSV(sortedLeave.map((l) => ({ Employee: empById[l.employee_id]?.name, Type: l.type, Start: l.start_date, End: l.end_date, Days: l.days, Status: l.status })), "leave_requests.csv")} />}
        </div>
        <table className="w-full text-sm min-w-[600px]">
          <thead><tr className="text-left text-xs text-[var(--muted)] border-b border-[var(--border)]">{isManager && <th className="pb-2 pr-3">Employee</th>}<th className="pb-2 pr-3">Type</th><th className="pb-2 pr-3">Dates</th><th className="pb-2 pr-3">Days</th><th className="pb-2 pr-3">Status</th>{isManager && <th className="pb-2 pr-3"></th>}</tr></thead>
          <tbody>
            {sortedLeave.map((l) => (
              <tr key={l.id} className="border-b border-[var(--border)]/50">
                {isManager && <td className="py-2 pr-3">{empById[l.employee_id]?.name}</td>}
                <td className="py-2 pr-3">{l.type}</td>
                <td className="py-2 pr-3 rw-mono text-xs">{fmtDateShort(l.start_date)} → {fmtDateShort(l.end_date)}</td>
                <td className="py-2 pr-3">{l.days}</td>
                <td className="py-2 pr-3"><Badge tone={l.status === "Approved" ? "ok" : l.status === "Rejected" ? "bad" : "amber"}>{l.status}</Badge></td>
                {isManager && <td className="py-2 pr-3 flex gap-1.5">{l.status === "Pending" && (<><Btn size="sm" tone="ok" icon={Check} onClick={() => decide(l, "Approved")}>Approve</Btn><Btn size="sm" tone="danger" icon={X} onClick={() => decide(l, "Rejected")}>Reject</Btn></>)}</td>}
              </tr>
            ))}
            {sortedLeave.length === 0 && <tr><td colSpan={6} className="py-6 text-center text-[var(--muted)]">No leave requests yet.</td></tr>}
          </tbody>
        </table>
      </Card>
      {requesting && <Modal title="Request leave" onClose={() => setRequesting(false)}><LeaveForm onSave={submitLeave} onCancel={() => setRequesting(false)} /></Modal>}
    </div>
  );
}
function LeaveForm({ onSave, onCancel }) {
  const [form, setForm] = useState({ type: "Holiday", startDate: todayISO(), endDate: todayISO(), comment: "" });
  const [saving, setSaving] = useState(false);
  return (
    <form onSubmit={async (e) => { e.preventDefault(); setSaving(true); await onSave(form); setSaving(false); }}>
      <Field label="Type"><select value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })} className={inputCls}><option>Holiday</option><option>Sick</option><option>Emergency leave</option><option>Unpaid leave</option><option>Other</option></select></Field>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Start date"><input type="date" required value={form.startDate} onChange={(e) => setForm({ ...form, startDate: e.target.value })} className={inputCls} /></Field>
        <Field label="End date"><input type="date" required value={form.endDate} onChange={(e) => setForm({ ...form, endDate: e.target.value })} className={inputCls} /></Field>
      </div>
      <Field label="Comment (optional)"><textarea value={form.comment} onChange={(e) => setForm({ ...form, comment: e.target.value })} className={inputCls} rows={2} /></Field>
      <div className="flex justify-end gap-2"><Btn tone="ghost" onClick={onCancel}>Cancel</Btn><Btn type="submit" icon={Check} disabled={saving}>{saving ? "Sending…" : "Submit request"}</Btn></div>
    </form>
  );
}

/* =====================================================================
   EMPLOYEES  (manager: add + PIN provisioning)
===================================================================== */
function EmployeesView({ db, session, logAudit, showToast, reload }) {
  const [modal, setModal] = useState(null);
  const [pinModal, setPinModal] = useState(null);
  const storeName = (id) => db.stores.find((s) => s.id === id)?.name || "—";
  const deptName = (id) => db.departments.find((d) => d.id === id)?.name || "—";

  const save = async (form) => {
    try {
      if (form.id) {
        await restUpdate("employees", form.id, { name: form.name, email: form.email, phone: form.phone, job_role: form.jobRole, store_id: form.storeId, department_id: form.deptId, contracted_hours: form.contractedHours, holiday_allowance: form.holidayAllowance, start_date: form.startDate || null, is_manager: !!form.isManager }, session.token);
        logAudit("Updated employee", form.name);
      } else {
        await restInsert("employees", { employee_code: form.employeeCode.trim().toLowerCase(), name: form.name, email: form.email, phone: form.phone, job_role: form.jobRole, store_id: form.storeId, department_id: form.deptId, contracted_hours: form.contractedHours, holiday_allowance: form.holidayAllowance, start_date: form.startDate || null, is_manager: !!form.isManager }, session.token);
        logAudit("Added employee", form.name);
      }
      await reload(); setModal(null); showToast("Employee saved");
    } catch (e) { showToast(e.message, "bad"); }
  };
  const remove = async (emp) => {
    try { await restDelete("employees", emp.id, session.token); await reload(); logAudit("Removed employee", emp.name); showToast("Employee removed"); }
    catch (e) { showToast(e.message, "bad"); }
  };
  const setPin = async (emp, pin) => {
    try {
      const res = await invokeFn("provision-login", { employee_id: emp.id, pin }, session.token);
      logAudit("Set employee login PIN", emp.name);
      showToast(`Login ready for ${emp.name}`);
      return res;
    } catch (e) { showToast(e.message, "bad"); throw e; }
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-4">
        <div className="text-sm text-[var(--muted)]">{db.employees.length} employees</div>
        <Btn icon={Plus} onClick={() => setModal({ storeId: db.stores[0]?.id, deptId: db.departments[0]?.id, contractedHours: 20, holidayAllowance: 28 })}>Add employee</Btn>
      </div>
      <div className="grid md:grid-cols-2 lg:grid-cols-3 gap-3">
        {db.employees.map((e) => (
          <Card key={e.id}>
            <div className="flex items-start justify-between">
              <div><div className="font-semibold flex items-center gap-1.5">{e.name} {e.is_manager && <Badge tone="amber">Manager</Badge>}</div><div className="text-xs text-[var(--muted)]">{e.job_role}</div></div>
              <div className="flex gap-1">
                <button onClick={() => setModal({ id: e.id, name: e.name, email: e.email, phone: e.phone, jobRole: e.job_role, storeId: e.store_id, deptId: e.department_id, contractedHours: e.contracted_hours, holidayAllowance: e.holiday_allowance, startDate: e.start_date, isManager: e.is_manager })} className="p-1.5 rounded-md bg-[var(--surface2)] border border-[var(--border)]"><Edit3 size={12} /></button>
                <button onClick={() => remove(e)} className="p-1.5 rounded-md bg-[var(--surface2)] border border-[var(--border)] text-[var(--bad)]"><Trash2 size={12} /></button>
              </div>
            </div>
            <div className="mt-3 text-xs space-y-1 text-[var(--muted)]">
              <div>{storeName(e.store_id)} · {deptName(e.department_id)}</div>
              <div>Contracted: <span className="rw-mono text-[var(--text)]">{e.contracted_hours}h/wk</span></div>
              <div>Holiday: <span className="rw-mono text-[var(--text)]">{(e.holiday_allowance || 0) - (e.holiday_used || 0)}/{e.holiday_allowance} days left</span></div>
              <div>Login code: <span className="rw-mono text-[var(--text)]">{e.employee_code}</span> {e.auth_user_id ? <Badge tone="ok">Active</Badge> : <Badge tone="amber">No PIN set</Badge>}</div>
            </div>
            <Btn size="sm" tone="ghost" icon={KeyRound} onClick={() => setPinModal(e)}>{e.auth_user_id ? "Reset PIN" : "Set PIN & activate"}</Btn>
          </Card>
        ))}
      </div>
      {modal && <Modal title={modal.id ? "Edit employee" : "Add employee"} onClose={() => setModal(null)}><EmployeeForm initial={modal} stores={db.stores} departments={db.departments} onSave={save} onCancel={() => setModal(null)} /></Modal>}
      {pinModal && <Modal title={`Set PIN for ${pinModal.name}`} onClose={() => setPinModal(null)}><PinForm employee={pinModal} onSet={setPin} onDone={() => setPinModal(null)} /></Modal>}
    </div>
  );
}
function EmployeeForm({ initial, stores, departments, onSave, onCancel }) {
  const [form, setForm] = useState(initial);
  const [saving, setSaving] = useState(false);
  return (
    <form onSubmit={async (e) => { e.preventDefault(); setSaving(true); await onSave(form); setSaving(false); }}>
      <Field label="Full name"><input required value={form.name || ""} onChange={(e) => setForm({ ...form, name: e.target.value })} className={inputCls} /></Field>
      {!form.id && <Field label="Employee code (used for login)"><input required value={form.employeeCode || ""} onChange={(e) => setForm({ ...form, employeeCode: e.target.value })} className={inputCls} placeholder="e.g. emp1042" /></Field>}
      <Field label="Email (optional)"><input type="email" value={form.email || ""} onChange={(e) => setForm({ ...form, email: e.target.value })} className={inputCls} /></Field>
      <Field label="Phone (optional)"><input value={form.phone || ""} onChange={(e) => setForm({ ...form, phone: e.target.value })} className={inputCls} /></Field>
      <Field label="Job role"><input required value={form.jobRole || ""} onChange={(e) => setForm({ ...form, jobRole: e.target.value })} className={inputCls} placeholder="Sales Assistant" /></Field>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Store"><select value={form.storeId} onChange={(e) => setForm({ ...form, storeId: e.target.value })} className={inputCls}>{stores.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}</select></Field>
        <Field label="Department"><select value={form.deptId} onChange={(e) => setForm({ ...form, deptId: e.target.value })} className={inputCls}>{departments.filter((d) => d.store_id === form.storeId).map((d) => <option key={d.id} value={d.id}>{d.name}</option>)}</select></Field>
      </div>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Contracted hours/week"><input type="number" step="0.5" required value={form.contractedHours} onChange={(e) => setForm({ ...form, contractedHours: Number(e.target.value) })} className={inputCls} /></Field>
        <Field label="Holiday allowance (days)"><input type="number" step="0.5" required value={form.holidayAllowance} onChange={(e) => setForm({ ...form, holidayAllowance: Number(e.target.value) })} className={inputCls} /></Field>
      </div>
      <Field label="Start date"><input type="date" value={form.startDate || ""} onChange={(e) => setForm({ ...form, startDate: e.target.value })} className={inputCls} /></Field>
      <label className="flex items-center gap-2 text-sm mb-4"><input type="checkbox" checked={!!form.isManager} onChange={(e) => setForm({ ...form, isManager: e.target.checked })} /> Store manager access</label>
      <div className="flex justify-end gap-2"><Btn tone="ghost" onClick={onCancel}>Cancel</Btn><Btn type="submit" icon={Check} disabled={saving}>{saving ? "Saving…" : "Save"}</Btn></div>
    </form>
  );
}
function PinForm({ employee, onSet, onDone }) {
  const [pin, setPinVal] = useState("");
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState(null);
  const submit = async (e) => {
    e.preventDefault(); setBusy(true);
    try { const res = await onSet(employee, pin); setResult(res); } finally { setBusy(false); }
  };
  if (result) {
    const shareText = `Rotawise login for ${employee.name}\nEmployee code: ${employee.employee_code}\nPIN: ${pin}\nLogin at your Rotawise link.`;
    return (
      <div>
        <div className="text-sm mb-3">Login is ready. Share these details with {employee.name} directly (not in a public group):</div>
        <div className="bg-[var(--surface2)] rounded-lg p-3 rw-mono text-sm mb-4">Code: {employee.employee_code}<br />PIN: {pin}</div>
        <div className="flex justify-end gap-2">
          <Btn tone="ghost" icon={MessageCircle} onClick={() => openWhatsapp(shareText)}>Send via WhatsApp</Btn>
          <Btn onClick={onDone}>Done</Btn>
        </div>
      </div>
    );
  }
  return (
    <form onSubmit={submit}>
      <div className="text-xs text-[var(--muted)] mb-3">Employee code: <span className="rw-mono text-[var(--text)]">{employee.employee_code}</span></div>
      <Field label="New PIN (4–6 digits)"><input required minLength={4} maxLength={6} inputMode="numeric" value={pin} onChange={(e) => setPinVal(e.target.value)} className={inputCls} /></Field>
      <div className="flex justify-end gap-2"><Btn tone="ghost" onClick={onDone}>Cancel</Btn><Btn type="submit" icon={KeyRound} disabled={busy}>{busy ? "Setting…" : "Set PIN"}</Btn></div>
    </form>
  );
}

/* =====================================================================
   STORES
===================================================================== */
function StoresView({ db, session, logAudit, showToast, reload }) {
  const [newStore, setNewStore] = useState("");
  const [newDept, setNewDept] = useState({});
  const addStore = async () => { if (!newStore.trim()) return; try { await restInsert("stores", { name: newStore.trim() }, session.token); await reload(); logAudit("Added store", newStore.trim()); setNewStore(""); showToast("Store added"); } catch (e) { showToast(e.message, "bad"); } };
  const addDept = async (storeId) => { const name = (newDept[storeId] || "").trim(); if (!name) return; try { await restInsert("departments", { name, store_id: storeId }, session.token); await reload(); logAudit("Added department", name); setNewDept({ ...newDept, [storeId]: "" }); showToast("Department added"); } catch (e) { showToast(e.message, "bad"); } };
  const removeStore = async (store) => { try { await restDelete("stores", store.id, session.token); await reload(); logAudit("Removed store", store.name); showToast("Store removed"); } catch (e) { showToast(e.message, "bad"); } };
  const removeDept = async (dept) => { try { await restDelete("departments", dept.id, session.token); await reload(); logAudit("Removed department", dept.name); showToast("Department removed"); } catch (e) { showToast(e.message, "bad"); } };

  return (
    <div className="space-y-4">
      <Card>
        <div className="text-sm font-semibold mb-3">Add a store</div>
        <div className="flex gap-2"><input value={newStore} onChange={(e) => setNewStore(e.target.value)} className={inputCls} placeholder="e.g. Wickwar Forecourt" /><Btn icon={Plus} onClick={addStore}>Add</Btn></div>
      </Card>
      <div className="grid md:grid-cols-2 gap-3">
        {db.stores.map((store) => (
          <Card key={store.id}>
            <div className="flex items-center justify-between mb-3"><div className="font-semibold flex items-center gap-1.5"><Store size={14} /> {store.name}</div><button onClick={() => removeStore(store)} className="text-[var(--bad)]"><Trash2 size={13} /></button></div>
            <div className="space-y-1.5 mb-3">
              {db.departments.filter((d) => d.store_id === store.id).map((d) => (<div key={d.id} className="flex items-center justify-between text-sm bg-[var(--surface2)] rounded-lg px-2.5 py-1.5"><span>{d.name}</span><button onClick={() => removeDept(d)} className="text-[var(--bad)]"><Trash2 size={12} /></button></div>))}
            </div>
            <div className="flex gap-2"><input value={newDept[store.id] || ""} onChange={(e) => setNewDept({ ...newDept, [store.id]: e.target.value })} className={inputCls} placeholder="New department" /><Btn size="sm" icon={Plus} onClick={() => addDept(store.id)}>Add</Btn></div>
          </Card>
        ))}
      </div>
    </div>
  );
}

/* =====================================================================
   EXCEL IMPORT
===================================================================== */
function ExcelImportView({ db, session, logAudit, showToast, reload }) {
  const [rawRows, setRawRows] = useState(null);
  const [headers, setHeaders] = useState([]);
  const [mapping, setMapping] = useState({});
  const [storeId, setStoreId] = useState(db.stores[0]?.id || "");
  const [importing, setImporting] = useState(false);

  const FIELDS = [
    { key: "employee", label: "Employee name", required: true },
    { key: "date", label: "Date", required: true },
    { key: "start", label: "Start time", required: true },
    { key: "end", label: "End time", required: true },
    { key: "breakMins", label: "Break (mins)", required: false },
  ];

  const onFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(evt.target.result, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
      if (!json.length) { showToast("That sheet looks empty", "bad"); return; }
      const cols = Object.keys(json[0]);
      setHeaders(cols);
      setRawRows(json);
      const guess = {};
      cols.forEach((c) => {
        const low = c.toLowerCase();
        if (!guess.employee && /name|employee/.test(low)) guess.employee = c;
        if (!guess.date && /date/.test(low)) guess.date = c;
        if (!guess.start && /start|clock in|from/.test(low)) guess.start = c;
        if (!guess.end && /end|finish|clock out|to/.test(low)) guess.end = c;
        if (!guess.breakMins && /break/.test(low)) guess.breakMins = c;
      });
      setMapping(guess);
    };
    reader.readAsArrayBuffer(file);
  };

  const parseTime = (v) => {
    if (v == null || v === "") return null;
    if (typeof v === "number") { const mins = Math.round(v * 24 * 60); return `${String(Math.floor(mins / 60)).padStart(2, "0")}:${String(mins % 60).padStart(2, "0")}`; }
    const s = String(v).trim();
    const m = s.match(/^(\d{1,2}):(\d{2})/);
    if (m) return `${m[1].padStart(2, "0")}:${m[2]}`;
    return null;
  };
  const parseDate = (v) => {
    if (v == null || v === "") return null;
    if (typeof v === "number") { const d = XLSX.SSF.parse_date_code(v); return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`; }
    const s = String(v).trim();
    const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
    const uk = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (uk) { let yr = uk[3].length === 2 ? `20${uk[3]}` : uk[3]; return `${yr}-${uk[2].padStart(2, "0")}-${uk[1].padStart(2, "0")}`; }
    return null;
  };

  const preview = (rawRows || []).map((row, i) => {
    const name = String(row[mapping.employee] || "").trim();
    const emp = db.employees.find((e) => e.name.toLowerCase() === name.toLowerCase());
    const date = parseDate(row[mapping.date]);
    const start = parseTime(row[mapping.start]);
    const end = parseTime(row[mapping.end]);
    const breakMins = mapping.breakMins ? Number(row[mapping.breakMins]) || 0 : 0;
    const errors = [];
    if (!name) errors.push("missing name");
    else if (!emp) errors.push("employee not found");
    if (!date) errors.push("bad/missing date");
    if (!start) errors.push("bad/missing start time");
    if (!end) errors.push("bad/missing end time");
    return { i, name, employeeId: emp?.id, date, start, end, breakMins, errors };
  });
  const validRows = preview.filter((r) => r.errors.length === 0);
  const invalidRows = preview.filter((r) => r.errors.length > 0);

  const doImport = async () => {
    setImporting(true);
    try {
      const payload = validRows.map((r) => ({ employee_id: r.employeeId, store_id: storeId, shift_date: r.date, start_time: r.start, end_time: r.end, break_mins: r.breakMins, status: "scheduled", source: "excel_import", created_by: session.employee.id }));
      if (payload.length) await restInsert("shifts", payload, session.token);
      await restInsert("excel_imports", { store_id: storeId, uploaded_by: session.employee.id, filename: "uploaded_sheet.xlsx", row_count: payload.length, status: "completed" }, session.token);
      logAudit("Excel rota import", `${payload.length} shifts imported, ${invalidRows.length} rows skipped`);
      await reload();
      showToast(`Imported ${payload.length} shifts${invalidRows.length ? ` (${invalidRows.length} skipped)` : ""}`);
      setRawRows(null); setHeaders([]); setMapping({});
    } catch (e) { showToast(e.message, "bad"); }
    finally { setImporting(false); }
  };

  return (
    <div className="space-y-4">
      <Card>
        <div className="text-sm font-semibold mb-2">Upload an existing rota spreadsheet</div>
        <div className="text-xs text-[var(--muted)] mb-3">Any column layout works — you'll map the columns below. Names must match an existing employee record exactly.</div>
        <input type="file" accept=".xlsx,.xls,.csv" onChange={onFile} className="text-sm" />
      </Card>

      {rawRows && (
        <>
          <Card>
            <div className="text-sm font-semibold mb-3">Map columns</div>
            <div className="grid sm:grid-cols-2 gap-3">
              {FIELDS.map((f) => (
                <Field key={f.key} label={f.label + (f.required ? " *" : "")}>
                  <select value={mapping[f.key] || ""} onChange={(e) => setMapping({ ...mapping, [f.key]: e.target.value })} className={inputCls}>
                    <option value="">— not mapped —</option>
                    {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                  </select>
                </Field>
              ))}
              <Field label="Import into store">
                <select value={storeId} onChange={(e) => setStoreId(e.target.value)} className={inputCls}>{db.stores.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}</select>
              </Field>
            </div>
          </Card>

          <Card>
            <div className="flex items-center justify-between mb-3">
              <div className="text-sm font-semibold">Preview — {validRows.length} ready, {invalidRows.length} need attention</div>
              <Btn icon={Upload} onClick={doImport} disabled={importing || !validRows.length}>{importing ? "Importing…" : `Import ${validRows.length} shifts`}</Btn>
            </div>
            <div className="overflow-x-auto rw-scrollbar max-h-72 overflow-y-auto">
              <table className="w-full text-sm min-w-[560px]">
                <thead><tr className="text-left text-xs text-[var(--muted)] border-b border-[var(--border)]"><th className="pb-2 pr-3">Name</th><th className="pb-2 pr-3">Date</th><th className="pb-2 pr-3">Start</th><th className="pb-2 pr-3">End</th><th className="pb-2 pr-3">Break</th><th className="pb-2 pr-3">Status</th></tr></thead>
                <tbody>
                  {preview.map((r) => (
                    <tr key={r.i} className="border-b border-[var(--border)]/50">
                      <td className="py-1.5 pr-3">{r.name || "—"}</td>
                      <td className="py-1.5 pr-3 rw-mono">{r.date || "—"}</td>
                      <td className="py-1.5 pr-3 rw-mono">{r.start || "—"}</td>
                      <td className="py-1.5 pr-3 rw-mono">{r.end || "—"}</td>
                      <td className="py-1.5 pr-3 rw-mono">{r.breakMins}</td>
                      <td className="py-1.5 pr-3">{r.errors.length ? <Badge tone="bad">{r.errors.join(", ")}</Badge> : <Badge tone="ok">Ready</Badge>}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
        </>
      )}
    </div>
  );
}

/* =====================================================================
   AUDIT
===================================================================== */
function AuditView({ db }) {
  return (
    <Card className="overflow-x-auto rw-scrollbar">
      <div className="flex items-center justify-between mb-3">
        <div className="text-sm font-semibold">Audit log</div>
        <ExportBar onXlsx={() => exportWorkbook([{ name: "Audit", rows: db.auditLog.map((a) => ({ Time: new Date(a.created_at).toLocaleString("en-GB"), Actor: a.actor_label, Action: a.action, Details: a.details })) }], "audit_log.xlsx")} onCsv={() => exportCSV(db.auditLog.map((a) => ({ Time: new Date(a.created_at).toLocaleString("en-GB"), Actor: a.actor_label, Action: a.action, Details: a.details })), "audit_log.csv")} />
      </div>
      <table className="w-full text-sm min-w-[600px]">
        <thead><tr className="text-left text-xs text-[var(--muted)] border-b border-[var(--border)]"><th className="pb-2 pr-3">Time</th><th className="pb-2 pr-3">Actor</th><th className="pb-2 pr-3">Action</th><th className="pb-2 pr-3">Details</th></tr></thead>
        <tbody>
          {db.auditLog.map((a) => (<tr key={a.id} className="border-b border-[var(--border)]/50"><td className="py-2 pr-3 rw-mono text-xs text-[var(--muted)] whitespace-nowrap">{new Date(a.created_at).toLocaleString("en-GB")}</td><td className="py-2 pr-3">{a.actor_label}</td><td className="py-2 pr-3">{a.action}</td><td className="py-2 pr-3 text-[var(--muted)]">{a.details}</td></tr>))}
          {db.auditLog.length === 0 && <tr><td colSpan={4} className="py-6 text-center text-[var(--muted)]">No activity yet.</td></tr>}
        </tbody>
      </table>
    </Card>
  );
}
