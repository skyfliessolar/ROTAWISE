/* Rotawise shared Firebase layer (hub, Goods Orders).
   Everything runs on the free Firebase plan, straight from the browser; firestore.rules does the guarding.

   Shops: anyone can create one. Each shop has its own Firebase login (shop-<id>@shops.rotawise.app)
   whose password is kept in hub_keys/<id>, so:
     open shop     -> the key is readable by anyone, the hub signs in by itself
     password shop -> only the shop itself and the admin can read the key; staff type it once per device
   Shop data lives under hub_shops/<id>/... and only that shop's login (or the admin) can touch it.
   The business itself (the owner's stores and Gas staff logins) uses the space 'main'. */
export const CFG = {
  apiKey: "AIzaSyBGR-i8LHuPHPL4sSCFJCwzYRqgnTSeVH8",
  authDomain: "rotawise-a2b16.firebaseapp.com",
  projectId: "rotawise-a2b16",
  storageBucket: "rotawise-a2b16.firebasestorage.app",
  messagingSenderId: "506340621414",
  appId: "1:506340621414:web:2f5ff84b2ddfef0905c1a7"
};
export const OWNER_UID = '5nI7ZEVbX8Uya6Kc2HGeEpu5tFx1';
export const SHOP_DOMAIN = 'shops.rotawise.app';
const STAFF_DOMAIN = 'staff.rotawise-gas.app';
const V = '10.14.1', B = `https://www.gstatic.com/firebasejs/${V}/`;

export const shopEmail = id => `shop-${id}@${SHOP_DOMAIN}`;
export const shopIdOf = email => { const m = new RegExp('^shop-([a-z0-9]+)@' + SHOP_DOMAIN.replace(/\./g, '\\.') + '$').exec(String(email || '').toLowerCase()); return m ? m[1] : null; };
const authPw = raw => 'rws-' + raw;                         // Firebase wants 6+ characters; shop passwords can be 4
const pinPw = pin => 'rwgas-' + pin;
const staffEmail = (code, v) => `${code}.v${v || 1}@${STAFF_DOMAIN}`;
const normCode = s => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
const rand = (n, abc = 'abcdefghijkmnpqrstuvwxyz23456789') => { const a = new Uint32Array(n); crypto.getRandomValues(a); return Array.from(a, x => abc[x % abc.length]).join(''); };
const cleanName = s => String(s || '').trim().replace(/\s+/g, ' ');

let boot;
export function fb() {
  return boot ||= (async () => {
    const [A, F, Au] = await Promise.all([import(B + 'firebase-app.js'), import(B + 'firebase-firestore.js'), import(B + 'firebase-auth.js')]);
    const app = A.getApps().find(a => a.name === '[DEFAULT]') || A.initializeApp(CFG);
    const db = F.getFirestore(app), auth = Au.getAuth(app);
    await auth.authStateReady();
    return { A, F, Au, app, db, auth };
  })();
}
/* a second, throw-away sign-in so the admin can act as a shop without losing their own session */
async function side(fn) {
  const { A, F, Au } = await fb();
  const app = A.getApps().find(a => a.name === 'shop-op') || A.initializeApp(CFG, 'shop-op');
  const auth = Au.getAuth(app), db = F.getFirestore(app);
  try { return await fn(auth, db); } finally { try { await Au.signOut(auth); } catch (e) {} }
}
const friendly = e => {
  const c = (e && (e.code || e.message)) || '';
  if (/wrong-password|invalid-credential|user-not-found|invalid-login/.test(c)) return 'wrong';
  if (/network|unavailable|offline/i.test(c)) return 'No internet connection. Try again when you are online.';
  if (/too-many-requests/.test(c)) return 'Too many tries. Wait a few minutes and try again.';
  if (/permission/.test(c)) return 'Not allowed. Ask your admin.';
  return (e && e.message) || String(e);
};
export { friendly };

/* ---------- who is signed in ---------- */
export async function whoami() {
  const { F, db, auth } = await fb();
  const u = auth.currentUser; if (!u) return null;
  const sid = shopIdOf(u.email);
  if (sid) {
    const s = await F.getDoc(F.doc(db, 'hub_shops', sid));
    if (!s.exists()) return null;
    return { kind: 'shop', uid: u.uid, shopId: sid, space: sid, name: s.data().name, access: s.data().access, hidden: !!s.data().hidden, admin: false, manager: true, role: 'shop' };
  }
  if (u.uid === OWNER_UID) return { kind: 'admin', uid: u.uid, space: 'main', name: 'Admin', email: u.email, admin: true, manager: true, role: 'admin', stores: [] };
  const d = await F.getDoc(F.doc(db, 'gas_staff', u.uid));
  if (!d.exists() || d.data().active !== true) return { kind: 'off', uid: u.uid };
  const x = d.data(), role = x.role || 'staff';
  return { kind: 'staff', uid: u.uid, space: 'main', name: x.name || x.code, code: x.code, email: x.email || '', role,
    admin: role === 'admin', manager: role === 'admin' || role === 'manager', stores: Array.isArray(x.stores) ? x.stores : [], apps: x.apps || null };
}
export async function signOut() { const { Au, auth } = await fb(); try { await Au.signOut(auth); } catch (e) {} }

/* personal login: Gas login code + PIN, or the owner's email + password */
export async function signInPerson(id, pin) {
  const { F, Au, db, auth } = await fb();
  id = String(id || '').trim(); pin = String(pin || '').trim();
  try {
    if (id.includes('@')) { await Au.signInWithEmailAndPassword(auth, id.toLowerCase(), pin); return { ok: true }; }
    const code = normCode(id); if (!code || !pin) return { error: 'wrong' };
    const L = await F.getDoc(F.doc(db, 'gas_logins', code)); if (!L.exists()) return { error: 'wrong' };
    await Au.signInWithEmailAndPassword(auth, staffEmail(code, L.data().v), pinPw(pin));
    return { ok: true, code };
  } catch (e) { return { error: friendly(e) }; }
}

/* ---------- shops ---------- */
export async function listShops() {
  const { F, db } = await fb();
  const s = await F.getDocs(F.collection(db, 'hub_shops'));
  return s.docs.map(d => ({ id: d.id, ...d.data() })).filter(x => !x.hidden).sort((a, b) => (a.nameLower || '').localeCompare(b.nameLower || ''));
}
export async function enterShop(id, typed) {
  const { F, Au, db, auth } = await fb();
  let raw = typed;
  if (raw == null) {
    let k; try { k = await F.getDoc(F.doc(db, 'hub_keys', id)); } catch (e) { return { error: 'locked' }; }
    if (!k.exists()) return { error: 'locked' };
    raw = k.data().pw;
  }
  try { if (auth.currentUser) await Au.signOut(auth); await Au.signInWithEmailAndPassword(auth, shopEmail(id), authPw(raw)); }
  catch (e) { return { error: friendly(e) }; }
  return { ok: true };
}
export async function createShop(name, password, { keepSession, hidden } = {}) {
  const { F, Au, db, auth } = await fb();
  name = cleanName(name);
  if (name.length < 2 || name.length > 60) throw new Error('Give the shop a name (2 to 60 letters).');
  if (password && String(password).length < 4) throw new Error('The shop password needs at least 4 characters.');
  const same = await F.getDocs(F.query(F.collection(db, 'hub_shops'), F.where('nameLower', '==', name.toLowerCase())));
  if (!same.empty) throw new Error('There is already a shop with that name. Pick it from the list instead.');
  const id = rand(12), raw = password ? String(password) : rand(16), now = Date.now();
  const make = async (a, d) => {
    await Au.createUserWithEmailAndPassword(a, shopEmail(id), authPw(raw));
    await F.setDoc(F.doc(d, 'hub_keys', id), { pw: raw, updatedAt: now });
    await F.setDoc(F.doc(d, 'hub_shops', id), { name, nameLower: name.toLowerCase(), access: password ? 'password' : 'open', createdAt: now, updatedAt: now, ...(hidden ? { hidden: true } : {}) });
  };
  if (keepSession) await side(make);
  else { if (auth.currentUser) await Au.signOut(auth); await make(auth, db); }
  return { id, name, access: password ? 'password' : 'open', pw: raw };
}
/* open <-> password, or a new password. Called by the shop itself or by the admin. */
export async function setShopAccess(id, mode, password) {
  const { F, Au, db, auth } = await fb();
  if (mode !== 'open' && mode !== 'password') throw new Error('Unknown setting');
  if (mode === 'password' && (!password || String(password).length < 4)) throw new Error('The shop password needs at least 4 characters.');
  const k = await F.getDoc(F.doc(db, 'hub_keys', id)); if (!k.exists()) throw new Error('Shop not found');
  const old = k.data().pw, raw = mode === 'open' ? rand(16) : String(password);
  const amShop = auth.currentUser && shopIdOf(auth.currentUser.email) === id;
  if (amShop) await Au.updatePassword(auth.currentUser, authPw(raw));
  else await side(async a => { const c = await Au.signInWithEmailAndPassword(a, shopEmail(id), authPw(old)); await Au.updatePassword(c.user, authPw(raw)); });
  try { await F.setDoc(F.doc(db, 'hub_keys', id), { pw: raw, updatedAt: Date.now() }); }
  catch (e) { // put the old password back so the key and the login never disagree
    if (amShop) await Au.updatePassword(auth.currentUser, authPw(old)); throw e; }
  await F.updateDoc(F.doc(db, 'hub_shops', id), { access: mode, updatedAt: Date.now() });
  return { pw: raw };
}
export async function renameShop(id, name) {
  const { F, db } = await fb(); name = cleanName(name);
  if (name.length < 2 || name.length > 60) throw new Error('Give the shop a name (2 to 60 letters).');
  const same = await F.getDocs(F.query(F.collection(db, 'hub_shops'), F.where('nameLower', '==', name.toLowerCase())));
  if (same.docs.some(d => d.id !== id)) throw new Error('There is already a shop with that name. Pick another name.');
  await F.updateDoc(F.doc(db, 'hub_shops', id), { name, nameLower: name.toLowerCase(), hidden: false, updatedAt: Date.now() });
}
export async function hideShop(id, hidden) { const { F, db } = await fb(); await F.updateDoc(F.doc(db, 'hub_shops', id), { hidden: !!hidden, updatedAt: Date.now() }); }
/* admin: every shop with its password */
export async function adminShops() {
  const { F, db } = await fb();
  const [s, k] = await Promise.all([F.getDocs(F.collection(db, 'hub_shops')), F.getDocs(F.collection(db, 'hub_keys'))]);
  const keys = new Map(k.docs.map(d => [d.id, d.data().pw]));
  return s.docs.map(d => ({ id: d.id, ...d.data(), pw: keys.get(d.id) || null })).sort((a, b) => (a.nameLower || '').localeCompare(b.nameLower || ''));
}
export async function myShopKey(id) { const { F, db } = await fb(); const k = await F.getDoc(F.doc(db, 'hub_keys', id)); return k.exists() ? k.data().pw : null; }

/* ---------- app access (business staff) ---------- */
export async function appModes() { const { F, db } = await fb(); try { const s = await F.getDoc(F.doc(db, 'hub_settings', 'apps')); return s.exists() ? s.data() : {}; } catch (e) { return {}; } }
export async function setAppModes(v) { const { F, db } = await fb(); await F.setDoc(F.doc(db, 'hub_settings', 'apps'), v); }
export async function staffList() {
  const { F, db } = await fb();
  const s = await F.getDocs(F.collection(db, 'gas_staff'));
  return s.docs.map(d => ({ id: d.id, ...d.data() })).filter(x => !x.replacedBy && !x.owner).sort((a, b) => String(a.name || '').localeCompare(String(b.name || '')));
}
export async function setStaffApps(uid, apps) { const { F, db } = await fb(); await F.updateDoc(F.doc(db, 'gas_staff', uid), { apps }); }

/* ---------- generic data helpers for a space (hub_shops/<space>/<coll>) ---------- */
export async function space(sp) {
  const { F, db } = await fb();
  const col = c => F.collection(db, 'hub_shops', sp, c), ref = (c, id) => F.doc(db, 'hub_shops', sp, c, id);
  const rows = s => s.docs.map(d => ({ id: d.id, ...d.data() }));
  return {
    F, db, col, ref,
    async where(c, f, v) { return rows(await F.getDocs(F.query(col(c), F.where(f, '==', v)))); },
    async all(c) { return rows(await F.getDocs(col(c))); },
    async get(c, id) { const s = await F.getDoc(ref(c, id)); return s.exists() ? { id: s.id, ...s.data() } : null; },
    async add(c, o) { const r = await F.addDoc(col(c), o); return { id: r.id, ...o }; },
    async set(c, id, o, merge) { await F.setDoc(ref(c, id), o, merge ? { merge: true } : undefined); },
    async patch(c, id, o) { await F.updateDoc(ref(c, id), o); },
    async del(c, id) { await F.deleteDoc(ref(c, id)); },
    batch() { return F.writeBatch(db); },
    newRef(c) { return F.doc(col(c)); }
  };
}
/* the business's store list (from Gas settings) */
export async function businessStores() {
  const { F, db } = await fb();
  try { const s = await F.getDoc(F.doc(db, 'gas_settings', 'main')); const a = s.exists() && Array.isArray(s.data().stores) ? s.data().stores.filter(x => x.active !== false) : [];
    return a.length ? a.map((x, i) => ({ id: x.id, name: x.name || ('Store ' + (i + 1)) })) : [{ id: 'main', name: 'Store 1' }]; }
  catch (e) { return [{ id: 'main', name: 'Store 1' }]; }
}
