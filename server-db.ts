import fs from 'fs/promises';
import fsSync from 'fs';
import path from 'path';
import { initializeApp, cert, getApps, type ServiceAccount } from 'firebase-admin/app';
import { getDatabase, type Database } from 'firebase-admin/database';
import { DatabaseState, Vehicle, RateMaster, Branch } from './src/types.js';

// สาขาเริ่มต้น (ข้อมูลเดิมทั้งหมดจะถูกผูกกับสาขานี้)
export const DEFAULT_BRANCH_ID = 'br-nakhonsawan';

function defaultBranches(): Branch[] {
  return [
    { id: 'br-hq', name: 'สำนักงานใหญ่ (HQ)', password: '9999', isHQ: true, status: 'active' },
    { id: DEFAULT_BRANCH_ID, name: 'นครสวรรค์', password: '1234', status: 'active' },
    { id: 'br-kamphaengphet', name: 'กำแพงเพชร', password: '1234', status: 'active' },
    { id: 'br-phitsanulok', name: 'พิษณุโลก', password: '1234', status: 'active' },
    { id: 'br-maesot', name: 'แม่สอด', password: '1234', status: 'active' },
    { id: 'br-sai3', name: 'สาย3', password: '1234', status: 'active' },
    { id: 'br-chiangmai', name: 'เชียงใหม่', password: '1234', status: 'active' },
  ];
}

const DB_FILE = path.join(process.cwd(), 'db.json');
const SEED_MASTERS_FILE = path.join(process.cwd(), 'seed-masters.json');
// snapshot cache ลง disk เร่ง boot (อ่าน local <5วิ แทนโหลด Firebase 35MB 2-16 นาที)
// Firebase = source of truth เสมอ; snapshot แค่เร่งความเร็ว boot + sync Firebase พื้นหลัง
// เขียนลง uploads/ (mount :rw บน NAS) — process.cwd()/dist mount :ro เขียนไม่ได้ (CFC เจอ snapshot ไม่เคยถูกเขียน)
const SNAPSHOT_FILE = process.env.CACHE_SNAPSHOT_FILE || path.join(process.cwd(), 'uploads', '.cache-snapshot.json');
const ENABLE_SNAPSHOT = process.env.DISABLE_CACHE_SNAPSHOT !== '1'; // ปิดได้ถ้าจำเป็น

// ---------------------------------------------------------------------------
// Firebase Realtime Database (เก็บข้อมูลจริง) — เปิดใช้เมื่อมี serviceAccountKey
// ออกแบบให้ "ดาวน์โหลดเกือบศูนย์": อ่าน Firebase ครั้งเดียวตอนบูต -> cache ใน memory
// ทุก getDb() ตอบจาก cache, saveDb() เขียนกลับ (write ไม่ถูกคิดเป็น download)
// ---------------------------------------------------------------------------
const FIREBASE_KEY_PATH = process.env.FIREBASE_KEY_PATH || path.join(process.cwd(), 'serviceAccountKey.json');
const FIREBASE_DB_URL =
  process.env.FIREBASE_DB_URL ||
  'https://excel-billing-cycle-summarizer-default-rtdb.asia-southeast1.firebasedatabase.app';

let firebaseDb: Database | null = null;
let cache: DatabaseState | null = null; // in-memory cache (ลดการอ่าน Firebase)
let loadingPromise: Promise<DatabaseState> | null = null; // กันโหลด Firebase พร้อมกันหลาย request (race ตอน cache null)
const FIREBASE_READ_TIMEOUT_MS = Number(process.env.FIREBASE_READ_TIMEOUT_MS) || 60000; // อ่าน Firebase เกินนี้ = throw (กัน getDb ค้างตลอดกาล -> server แฮงก์)

function loadServiceAccount(): ServiceAccount | null {
  // 1) จาก env var (สำหรับ deploy บน Render/Railway — ไม่ต้อง commit ไฟล์ key)
  const raw = process.env.FIREBASE_SERVICE_ACCOUNT;
  if (raw && raw.trim().startsWith('{')) {
    try { return JSON.parse(raw) as ServiceAccount; } catch { /* ตกไปอ่านไฟล์ */ }
  }
  // 2) จากไฟล์ (สำหรับ local dev)
  if (fsSync.existsSync(FIREBASE_KEY_PATH)) {
    return JSON.parse(fsSync.readFileSync(FIREBASE_KEY_PATH, 'utf-8')) as ServiceAccount;
  }
  return null;
}

function initFirebase(): Database | null {
  if (firebaseDb) return firebaseDb;
  const serviceAccount = loadServiceAccount();
  if (!serviceAccount) return null; // ยังไม่ใส่ key -> ใช้ db.json
  try {
    console.log('[firebase] เริ่ม initializeApp...'); // จับเวลา: ถ้าค้างนานหลังบรรทัดนี้ = init ค้าง
    const app = getApps().length
      ? getApps()[0]
      : initializeApp({ credential: cert(serviceAccount), databaseURL: FIREBASE_DB_URL });
    console.log('[firebase] initializeApp เสร็จ -> getDatabase...');
    firebaseDb = getDatabase(app);
    console.log('[firebase] getDatabase เสร็จ (ยังไม่ได้อ่าน root)'); // ✅ จริงย้ายไปหลัง read สำเร็จ (กัน false positive)
    return firebaseDb;
  } catch (e) {
    console.error('Firebase init ล้มเหลว ใช้ db.json แทน:', (e as Error).message);
    return null;
  }
}

// โหลด Master จริงที่ดึงมาจากไฟล์ Excel (ถ้ามี) ไม่งั้น fallback เป็นตัวอย่าง
function loadSeedMasters(): { vehicles: Vehicle[]; rateMasters: RateMaster[] } {
  try {
    const raw = fsSync.readFileSync(SEED_MASTERS_FILE, 'utf-8');
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed.vehicles) && Array.isArray(parsed.rateMasters)) {
      // ผูกข้อมูล seed กับสาขาเริ่มต้น (นครสวรรค์)
      const vehicles = parsed.vehicles.map((v: Vehicle) => ({ branchId: DEFAULT_BRANCH_ID, ...v }));
      const rateMasters = parsed.rateMasters.map((r: RateMaster) => ({ branchId: DEFAULT_BRANCH_ID, ...r }));
      return { vehicles, rateMasters };
    }
  } catch {
    /* ไม่มีไฟล์ -> ใช้ตัวอย่าง */
  }
  const now = new Date().toISOString();
  return {
    vehicles: [
      { id: 'veh-001', branchId: DEFAULT_BRANCH_ID, plateNo: '70-1234', driverName: 'สมชาย ขับดี', vehicleType: '6 ล้อ', status: 'active' },
      { id: 'veh-002', branchId: DEFAULT_BRANCH_ID, plateNo: '82-5678', driverName: 'สมหญิง ส่งไว', vehicleType: '4 ล้อ', status: 'active' },
    ],
    rateMasters: [
      { id: 'rate-001', branchId: DEFAULT_BRANCH_ID, destinationName: 'อ.เมือง จ.นว', provinceName: 'นครสวรรค์', provinceShort: 'นว', districtName: 'เมือง', priceType: 'flat', price: 700, effectiveFrom: '2026-05-01', effectiveTo: null, status: 'active', remark: '', createdBy: 'system', createdAt: now },
    ],
  };
}

// ---------------------------------------------------------------------------
// Seed: Master data (รถ/ราคา จากไฟล์ Excel จริง + กลุ่มผู้รับ/กฎตัวหาร)
// ---------------------------------------------------------------------------
function seedState(): DatabaseState {
  const { vehicles, rateMasters } = loadSeedMasters();
  return {
    settings: { geminiModel: 'gemini-3.5-flash' },
    branches: defaultBranches(),
    cycles: [],
    vehicles,
    rateMasters,
    rateOverrides: [],
    rateMasterHistory: [],
    receiverGroups: [
      { id: 'grp-macro', branchId: DEFAULT_BRANCH_ID, groupName: 'แม็คโคร/เซลส์', status: 'active' },
    ],
    receiverGroupAliases: [
      { id: 'al-1', branchId: DEFAULT_BRANCH_ID, receiverGroupId: 'grp-macro', aliasName: 'แม็คโคร', status: 'active' },
      { id: 'al-2', branchId: DEFAULT_BRANCH_ID, receiverGroupId: 'grp-macro', aliasName: 'MK', status: 'active' },
      { id: 'al-3', branchId: DEFAULT_BRANCH_ID, receiverGroupId: 'grp-macro', aliasName: 'CP AXTRA', status: 'active' },
      { id: 'al-4', branchId: DEFAULT_BRANCH_ID, receiverGroupId: 'grp-macro', aliasName: 'ซีพี แอ็กซ์ตร้า', status: 'active' },
      { id: 'al-5', branchId: DEFAULT_BRANCH_ID, receiverGroupId: 'grp-macro', aliasName: 'เซลส์', status: 'active' },
    ],
    conversionRules: [
      {
        id: 'rule-yupi', branchId: DEFAULT_BRANCH_ID, ruleName: 'ยูปี้ หาร 3', senderKeyword: 'ซีโน', receiverGroupId: '',
        productKeyword: 'ยูปี้', productSizeKeyword: '', divisor: 3, roundingMethod: 'half_up',
        applyLevel: 'receipt', status: 'active', effectiveFrom: '2020-01-01', effectiveTo: null,
        remark: 'ยูปี้ทุกชนิด ส่งแม็คโคร/ซีพีแอ็กซ์ตร้า',
      },
      {
        id: 'rule-pringles', branchId: DEFAULT_BRANCH_ID, ruleName: 'พริงเกิล หาร 3', senderKeyword: 'ซีโน', receiverGroupId: '',
        productKeyword: 'พริงเกิล', productSizeKeyword: '', divisor: 3, roundingMethod: 'half_up',
        applyLevel: 'receipt', status: 'active', effectiveFrom: '2020-01-01', effectiveTo: null,
        remark: 'พริงเกิลส ทุกรส',
      },
    ],
    manualBoxSenders: [
      { id: 'mbs-cpconsumer', branchId: DEFAULT_BRANCH_ID, senderKeyword: 'คอนซูเมอร์', note: 'ซ.พี.คอนซูเมอร์ โพรดักส์ — ส่งเป็นชิ้น ต้องกรอกจำนวนกล่องเอง', status: 'active' },
    ],
    destinationOverrides: [],
    moneyCategories: defaultMoneyCategories(),
    tripDocuments: [],
    fuelEntries: [],
    deductions: [],
    oilPrices: [], // [ทดลอง] ราคาน้ำมัน OR
    tripDistances: [], // [ทดลอง] cache ระยะลูป
    mountainRoutes: [], // [ทดลอง] master น้ำมันขึ้นเขา
  };
}

// ประเภทรายการเงินตั้งต้น (income = รายได้เพิ่ม, deduction = หักออก)
function defaultMoneyCategories() {
  const b = DEFAULT_BRANCH_ID;
  return [
    { id: 'cat-bill_update', branchId: b, name: 'ค่าอัพเดทบิล', kind: 'income' as const, status: 'active' as const, builtin: true },
    { id: 'cat-phone', branchId: b, name: 'ค่าโทรศัพท์', kind: 'deduction' as const, status: 'active' as const, builtin: true },
    { id: 'cat-gps', branchId: b, name: 'ค่า GPS', kind: 'deduction' as const, status: 'active' as const, builtin: true },
    { id: 'cat-loan', branchId: b, name: 'ยืมเงิน', kind: 'deduction' as const, status: 'active' as const, builtin: true },
    { id: 'cat-gps_yearly', branchId: b, name: 'GPS รายปี', kind: 'deduction' as const, status: 'active' as const, builtin: true },
    { id: 'cat-insurance', branchId: b, name: 'ประกัน', kind: 'deduction' as const, status: 'active' as const, builtin: true },
    { id: 'cat-other', branchId: b, name: 'อื่นๆ', kind: 'deduction' as const, status: 'active' as const, builtin: true },
  ];
}

// migrate รายการหักเก่า (มีแค่ type) -> ใส่ categoryId/kind
function migrateDeductions(list: any[]): any[] {
  return (list ?? []).map((d) => {
    if (d.categoryId && d.kind) return d;
    const kind = d.kind ?? (d.type === 'bill_update' ? 'income' : 'deduction');
    const categoryId = d.categoryId ?? `cat-${d.type ?? 'other'}`;
    return { ...d, kind, categoryId };
  });
}

// เติม branchId ให้ record เก่าที่ยังไม่มี (ผูกกับสาขาเริ่มต้น)
function withBranch<T extends object>(list: T[] | undefined, fallback: T[]): T[] {
  return (list ?? fallback).map((x: any) => (x.branchId ? x : { ...x, branchId: DEFAULT_BRANCH_ID }));
}

// Firebase ตัด array ว่างทิ้ง -> อ่านกลับมาเป็น undefined -> .map พัง
// คืนค่า array ที่ขาดให้ trip/receipt ทุกตัว
function normalizeTrips(list: any[]): any[] {
  return (list ?? []).map((t) => ({
    ...t,
    warnings: t.warnings ?? [],
    breakdown: t.breakdown ?? { normal: t.tripAmount ?? 0, collect: 0, peat: 0 },
    receipts: (t.receipts ?? []).map((r: any) => ({
      ...r,
      items: r.items ?? [],
      adjustments: r.adjustments ?? [],
      normalQty: r.normalQty ?? r.totalQty ?? 0,
      collectQty: r.collectQty ?? 0,
      collectPrice: r.collectPrice ?? null,
      collectFlatPrice: r.collectFlatPrice ?? null,
      peatQty: r.peatQty ?? 0,
      peatPrice: r.peatPrice ?? null,
    })),
  }));
}

// ---------------------------------------------------------------------------
// บันทึกแบบ "เฉพาะ record" (granular) — คอลเลกชันใหญ่เก็บใน Firebase แบบ id-keyed
//   { "rcp-abc": {...}, "rcp-def": {...} }  แทน array [...]
// -> บันทึก 1 ใบ = เขียนแค่ /tripDocuments/<id> (ไม่เขียนทั้ง DB) เร็วคงที่
// อ่านได้ทั้ง 2 รูปแบบ (array เดิม + object ใหม่) ด้วย toArray() ของเก่าจึงไม่หาย
// ---------------------------------------------------------------------------
const ID_KEYED: (keyof DatabaseState)[] = [
  'tripDocuments', 'fuelEntries', 'deductions', 'rateOverrides', 'rateMasters', 'rateMasterHistory',
  // master ที่แก้ผ่าน masterRoutes (CRUD ทีละรายการ) -> id-keyed = เขียนแค่ record เดียว ไม่เขียนทั้ง tree
  'vehicles', 'receiverGroups', 'receiverGroupAliases', 'conversionRules', 'manualBoxSenders', 'destinationOverrides', 'moneyCategories',
  'oilPrices', 'tripDistances', 'mountainRoutes', // [ทดลอง] ราคาน้ำมัน OR + cache ระยะลูป + master ขึ้นเขา
];
export function isIdKeyed(collKey: keyof DatabaseState): boolean {
  return ID_KEYED.includes(collKey);
}
const deepClean = <T>(x: T): T => JSON.parse(JSON.stringify(x)); // ตัด undefined (RTDB ไม่รับ)
const arrToMap = (arr: any[]): Record<string, any> => {
  const m: Record<string, any> = {};
  for (const r of arr || []) if (r && r.id) m[r.id] = r;
  return m;
};
// รับได้ทั้ง array (รูปแบบเก่า) และ object id-keyed (รูปแบบใหม่) -> คืน array เสมอ
const toArray = (x: any): any[] => (Array.isArray(x) ? x : x && typeof x === 'object' ? Object.values(x) : []);

// migrate: เติม key ที่ขาดให้ db เก่า
export function ensureShape(state: Partial<DatabaseState>): DatabaseState {
  const seed = seedState();
  // คอลเลกชัน id-keyed เก็บได้ทั้ง array (เก่า) และ map (ใหม่) -> แปลงเป็น array ก่อนเสมอ
  // คงพฤติกรรมเดิมเป๊ะ: node หายจริง (null/undefined) -> seed | array ว่าง (ลบหมดเอง) -> คงว่าง
  const kArr = <T,>(v: any): T[] | undefined => (v == null ? undefined : (toArray(v) as T[]));
  return {
    settings: { ...seed.settings, ...(state.settings || {}) },
    branches: state.branches && state.branches.length ? state.branches : seed.branches,
    cycles: state.cycles ?? [],
    vehicles: withBranch(kArr(state.vehicles), seed.vehicles),
    rateMasters: withBranch(kArr(state.rateMasters), seed.rateMasters),
    rateOverrides: toArray(state.rateOverrides), // อ่านได้ทั้ง array (เก่า) และ id-keyed map (ใหม่)
    rateMasterHistory: toArray(state.rateMasterHistory),
    receiverGroups: withBranch(kArr(state.receiverGroups), seed.receiverGroups),
    receiverGroupAliases: withBranch(kArr(state.receiverGroupAliases), seed.receiverGroupAliases),
    conversionRules: withBranch(kArr(state.conversionRules), seed.conversionRules),
    manualBoxSenders: withBranch(kArr(state.manualBoxSenders), seed.manualBoxSenders),
    destinationOverrides: withBranch(toArray(state.destinationOverrides), []),
    moneyCategories: withBranch(kArr(state.moneyCategories), seed.moneyCategories),
    tripDocuments: normalizeTrips(withBranch(toArray(state.tripDocuments), [])),
    fuelEntries: withBranch(toArray(state.fuelEntries), []),
    deductions: withBranch(migrateDeductions(toArray(state.deductions)), []),
    oilPrices: toArray(state.oilPrices), // [ทดลอง] อ่านได้ทั้ง array/map (ว่าง = [])
    tripDistances: toArray(state.tripDistances), // [ทดลอง] cache ระยะลูป
    mountainRoutes: toArray(state.mountainRoutes), // [ทดลอง] master น้ำมันขึ้นเขา
    fuelPolicy: state.fuelPolicy || undefined, // [ทดลอง] ค่าตั้งสูตร (singleton object; ว่าง = ใช้ default ในโค้ด)
  };
}

function withTimeout<T>(p: Promise<T>, ms: number, label: string): Promise<T> {
  let timer: NodeJS.Timeout;
  const timeout = new Promise<never>((_, reject) => {
    timer = setTimeout(() => {
      console.error(`[firebase] ⏱ guard เตะ: ${label} เกิน ${ms}ms -> throw`);
      reject(new Error(`อ่าน Firebase (${label}) เกิน ${ms}ms (network/connection ค้าง)`));
    }, ms);
    if (typeof timer.unref === 'function') timer.unref();
  });
  // clear timer เมื่อ p settle (สำเร็จ/ล้ม) -> ไม่ emit log timeout เก้อ + ไม่ค้าง timer (Codex P3)
  return Promise.race([p, timeout]).finally(() => clearTimeout(timer));
}

// รายชื่อ top-level node ทั้งหมดใน RTDB (อ่านทีละ node แทน root ทั้งก้อน)
const ROOT_NODES = [
  'settings', 'branches', 'cycles', 'vehicles', 'rateMasters', 'rateOverrides', 'rateMasterHistory',
  'receiverGroups', 'receiverGroupAliases', 'conversionRules', 'manualBoxSenders', 'destinationOverrides',
  'moneyCategories', 'tripDocuments', 'fuelEntries', 'deductions', 'tripDistances', 'oilPrices',
  'mountainRoutes', 'fuelPolicy',
] as const;

// อ่าน Firebase ทีละ top-level node (ไม่อ่าน root ทั้งก้อน) — กัน .val() build object 35MB
// แบบ sync ครั้งเดียว = block event loop นาน (พบว่าค้าง 30+ นาที). อ่านทีละ node + await ระหว่าง
// node = yield event loop ให้ /api/config, healthcheck ตอบได้ระหว่างโหลด
// ⚠️ Trade-off (Codex P2): ไม่ atomic เท่า once('/') — ถ้ามีคนเขียนหลาย node ระหว่างอ่านตอน boot
//    cache อาจผสม 2 จุดเวลา. ยอมรับได้เพราะ NEOSIAM มี instance เดียวเขียน (ดู memory
//    render-cache-overwrites-firebase) + boot มักไม่มีคนใช้ + ครั้งถัดไป reload/write-through แก้เอง.
//    แลกกับการไม่ block event loop 30 นาที (prod ล่ม) ซึ่งร้ายแรงกว่ามาก
async function readRootWithTimeout(fb: Database): Promise<Partial<DatabaseState> | null> {
  const out: Record<string, any> = {};
  let hasAny = false;
  const deadline = Date.now() + FIREBASE_READ_TIMEOUT_MS; // total timeout ครอบทั้ง loop (Codex P2)
  for (const node of ROOT_NODES) {
    const remain = deadline - Date.now();
    if (remain <= 0) throw new Error(`อ่าน Firebase เกิน ${FIREBASE_READ_TIMEOUT_MS}ms รวมทุก node (network ช้า)`);
    const t0 = Date.now();
    // timeout ต่อ node = เวลาที่เหลือของ deadline รวม (ไม่ให้รวมกันเกินลิมิตเดียว)
    const snap = await withTimeout(fb.ref(`/${node}`).once('value'), remain, `อ่าน ${node}`);
    const val = snap.val(); // .val() ของ node ใหญ่ (tripDocuments) อาจ block ตรงนี้ -> log เวลาจะเห็น
    if (val != null) { out[node] = val; hasAny = true; }
    console.log(`[firebase] node ${node}: ${Date.now() - t0}ms`); // จับเวลาต่อ node -> เห็นตัวที่ block/ช้า
    await new Promise((r) => setImmediate(r)); // yield event loop ระหว่าง node (กัน block ยาว)
  }
  return hasAny ? (out as Partial<DatabaseState>) : null;
}

// โหลด DB จากแหล่งจริง (Firebase หรือ db.json) — เรียกครั้งเดียว (ผ่าน lock ใน getDb)
async function loadDbFromSource(): Promise<DatabaseState> {
  const fb = initFirebase();
  if (fb) {
    console.log('[firebase] เริ่มอ่าน root (once)...'); // ถ้าค้างนานหลังบรรทัดนี้ = read/network ค้าง (ไม่ใช่ init)
    const val = await readRootWithTimeout(fb); // มี timeout กันค้าง
    console.log('✅ ใช้ Firebase Realtime Database — อ่าน root เสร็จ (read ได้จริง)'); // ✅ จริงตรงนี้ (ไม่ใช่ตอน getDatabase)
    const loaded = ensureShape(val || {});
    // ⚠️ ต้อง set cache ก่อน saveDb/flushCollection (มันอ่านจาก cache) ไม่งั้นเขียน empty ทับ = ข้อมูลหาย (Codex P1)
    cache = loaded;
    if (!val) await saveDb(loaded); // ว่างเปล่า -> seed ขึ้น Firebase
    else {
      // migrate ครั้งเดียว: id-keyed ที่ยังเป็น array -> map (เขียนแค่ node นั้น)
      for (const k of ID_KEYED) {
        const cur = (val as any)[k as string];
        if (Array.isArray(cur) && cur.length) await flushCollection(k);
      }
    }
    return loaded;
  }
  // Fallback: db.json (เมื่อยังไม่ใส่ serviceAccountKey)
  try {
    const content = await fs.readFile(DB_FILE, 'utf-8');
    const loaded = ensureShape(JSON.parse(content) as Partial<DatabaseState>);
    cache = loaded;
    return loaded;
  } catch {
    const seeded = seedState();
    cache = seeded;      // set ก่อน saveDb (อ่าน cache)
    await saveDb(seeded);
    return seeded;
  }
}

let bootedFromSnapshot = false; // boot จาก snapshot (ยังไม่ยืนยันกับ Firebase) -> snapshot ต้อง "ตรวจ" ก่อนอนุญาต write

// ตรวจ snapshot กับ Firebase 1 ครั้งตอน boot (Firebase = source of truth)
// สำเร็จ -> ทับ cache ด้วย Firebase + set bootedFromSnapshot=false (verified)
// ล้มเหลว (Firebase timeout) -> throw ต่อ; ผู้เรียกต้องไม่ set false (ให้ retry) - Codex P1a
async function verifySnapshotAgainstFirebase(): Promise<DatabaseState> {
  const fresh = await loadDbFromSource(); // โหลด Firebase จริง (มี timeout) -> set cache ในนั้น; throw ถ้าค้าง
  cache = fresh;
  bootedFromSnapshot = false; // ยืนยันสำเร็จแล้วเท่านั้นถึงปลดล็อก write
  await writeSnapshotNow();
  return cache!;
}

// readOnlySnapshot(): คืน snapshot ที่ยังไม่ verify — ใช้เฉพาะ read path ที่ยอมข้อมูลเก่าเสี้ยววินาที
// (เช่น /api/state ครั้งแรกตอน boot) เพื่อไม่ให้ผู้ใช้รอ Firebase; ไม่ใช้กับ write เด็ดขาด
// จำกัดอายุ: ถ้า verify Firebase ไม่สำเร็จนานเกิน STALE_LIMIT -> หยุดเสิร์ฟ snapshot (กัน stale ไม่มีที่สิ้นสุด - Codex)
const SNAPSHOT_STALE_LIMIT_MS = Number(process.env.SNAPSHOT_STALE_LIMIT_MS) || 15 * 60 * 1000; // 15 นาที
let snapshotBootAt = 0; // เวลาที่เริ่ม boot จาก snapshot (0 = ไม่ได้ boot จาก snapshot)
export function peekCache(): DatabaseState | null {
  if (!(cache && bootedFromSnapshot)) return null;
  // snapshot เก่าเกินลิมิต (verify Firebase ยังไม่สำเร็จ) -> ไม่เสิร์ฟ (ให้ getDb รอ/throw แทน)
  if (snapshotBootAt && Date.now() - snapshotBootAt > SNAPSHOT_STALE_LIMIT_MS) return null;
  return cache;
}

// โหลด snapshot เข้า cache ให้เสร็จก่อน (เร็ว จาก disk) แยกจากการ verify Firebase
// -> peekCache() ใช้ได้ทันทีตั้งแต่ boot (Codex P2: /api/state ไม่ต้องรอ Firebase)
let snapshotLoadPromise: Promise<void> | null = null;
async function ensureSnapshotLoaded(): Promise<void> {
  if (cache) return;                       // มี cache แล้ว (snapshot หรือ verified) -> ไม่ต้อง
  if (snapshotLoadPromise) return snapshotLoadPromise;
  snapshotLoadPromise = (async () => {
    const snap = await readSnapshot();
    if (snap && !cache) { cache = snap; bootedFromSnapshot = true; snapshotBootAt = Date.now(); } // ไม่ทับถ้ามีคนโหลดของจริงไปแล้ว
  })().finally(() => { snapshotLoadPromise = null; });
  return snapshotLoadPromise;
}

// เรียกตอน boot: โหลด snapshot เข้า cache ก่อน (peekCache พร้อมทันที) แล้วค่อย verify Firebase พื้นหลัง
// ล้มเหลว -> log แล้ว rethrow (Codex P3: ไม่ให้ caller log "cache พร้อม" หลอกทั้งที่โหลดไม่สำเร็จ)
export async function warmCacheOnBoot(): Promise<void> {
  await ensureSnapshotLoaded();   // snapshot พร้อม -> /api/state ตอบได้แม้ Firebase ยังช้า
  try { await getDb(); }
  catch (e) {
    console.error('[boot] ❌ โหลด cache ล้มเหลว (จะ retry ตอน request):', (e as Error)?.message || e);
    throw e; // ให้ caller รู้ว่ายังไม่พร้อมจริง (log สำเร็จเฉพาะเมื่อโหลดได้)
  }
}

// read เร็วสำหรับ /api/state: รอแค่ snapshot โหลด (จาก disk เร็ว) ไม่รอ Firebase verify
// คืน snapshot ถ้ามี (แม้ Firebase ช้า); ถ้าไม่มี snapshot -> null (ให้ route ไป await getDb ปกติ)
export async function peekCacheAsync(): Promise<DatabaseState | null> {
  if (cache && !bootedFromSnapshot) return cache; // verified แล้ว
  await ensureSnapshotLoaded(); // รอ snapshot โหลดเข้า cache (เร็ว) — ไม่ trigger verify Firebase
  return peekCache();
}

// getDb(): source of truth เสมอ (สำหรับทั้ง read+write ที่ต้องข้อมูลถูก)
// boot จาก snapshot -> verify กับ Firebase ให้เสร็จก่อนคืน (กัน write ทับด้วย snapshot เก่า - Codex P1)
// ยอมช้าครั้งแรก (โหลด Firebase 1 ครั้ง) แลกกับข้อมูลถูก 100%; หลังจากนั้น cache -> เร็วตลอด
export async function getDb(): Promise<DatabaseState> {
  if (cache && !bootedFromSnapshot) return cache; // verified แล้ว -> เร็ว
  if (loadingPromise) return loadingPromise;       // กำลังโหลด/verify -> รอ promise เดียว
  loadingPromise = (async () => {
    try {
      await ensureSnapshotLoaded(); // อ่าน snapshot เข้า cache ก่อน (peekCache ใช้ได้ระหว่าง verify)
      if (cache && bootedFromSnapshot) {
        // มี snapshot แต่ยังไม่ verify -> verify กับ Firebase ให้เสร็จก่อนคืน (throw ถ้า Firebase ค้าง)
        return await verifySnapshotAgainstFirebase();
      }
      // ไม่มี snapshot -> โหลด Firebase ตรง
      cache = await loadDbFromSource();
      await writeSnapshotNow();
      return cache;
    }
    finally { loadingPromise = null; }
  })();
  return loadingPromise;
}

// เขียนทั้ง DB (ใช้ตอน seed / bulk / migrate) — คอลเลกชันใหญ่เขียนแบบ id-keyed
export async function saveDb(state: DatabaseState): Promise<void> {
  cache = state; // อัปเดต cache ในหน่วยความจำ
  const fb = initFirebase();
  if (fb) {
    const out: any = deepClean(state); // ตัด undefined (RTDB ไม่รับ) ทั้งก้อนก่อน
    for (const k of ID_KEYED) out[k as string] = arrToMap(out[k as string] || []); // re-key จาก array ที่ clean แล้ว
    await fb.ref('/').set(out); // write-through (การเขียนไม่ถูกคิดเป็น download)
  } else {
    await fs.writeFile(DB_FILE, JSON.stringify(deepClean(state), null, 2), 'utf-8');
  }
  scheduleSnapshot();
}

// fallback db.json: เขียนทั้งไฟล์ (sandbox local เล็ก เขียนเร็ว)
async function persistLocal(): Promise<void> {
  if (cache) await fs.writeFile(DB_FILE, JSON.stringify(deepClean(cache), null, 2), 'utf-8');
}

// ---- snapshot cache ลง disk (เร่ง boot) ----
// เขียนแบบ atomic (temp -> rename) กันไฟล์เสียถ้า crash กลางเขียน
async function writeSnapshotNow(): Promise<void> {
  // เขียนเฉพาะเมื่อ cache = ของจริงจาก Firebase แล้ว (ยืนยันแล้ว ไม่ใช่ snapshot ที่ยังไม่ verify)
  // กัน snapshot ที่ยังไม่ verify เขียนทับตัวเอง / กันเขียนตอน cache ยังเป็น snapshot เก่า
  if (!ENABLE_SNAPSHOT || !cache || bootedFromSnapshot) return;
  try {
    const tmp = `${SNAPSHOT_FILE}.tmp`;
    await fs.writeFile(tmp, JSON.stringify(deepClean(cache)), 'utf-8');
    await fs.rename(tmp, SNAPSHOT_FILE); // atomic replace
  } catch { /* snapshot ล้มเหลวไม่ critical (Firebase ยังเป็น source of truth) */ }
}
// debounce: หลัง write เงียบ 5 วิ ค่อยเขียน snapshot (ไม่เขียน 35MB ทุก write = ไม่ block)
let snapshotTimer: NodeJS.Timeout | null = null;
function scheduleSnapshot(): void {
  if (!ENABLE_SNAPSHOT) return;
  if (snapshotTimer) clearTimeout(snapshotTimer);
  snapshotTimer = setTimeout(() => { snapshotTimer = null; void writeSnapshotNow(); }, 5000);
  snapshotTimer.unref?.(); // ไม่กัน process exit
}
// อ่าน snapshot ตอน boot -> คืน state ถ้าอ่านได้ (null ถ้าไม่มี/เสีย)
async function readSnapshot(): Promise<DatabaseState | null> {
  if (!ENABLE_SNAPSHOT) return null;
  try {
    const content = await fs.readFile(SNAPSHOT_FILE, 'utf-8');
    return ensureShape(JSON.parse(content) as Partial<DatabaseState>);
  } catch { return null; }
}

// บันทึก/แก้ 1 record — เขียนแค่ /<coll>/<id> (เร็วคงที่ ไม่ขึ้นกับขนาด DB)
// ผู้เรียกต้องอัปเดต cache array แล้ว (db = getDb() เป็น reference เดียวกับ cache)
export async function saveRecord(collKey: keyof DatabaseState, record: { id: string }): Promise<void> {
  const fb = initFirebase();
  if (fb) await fb.ref(`/${String(collKey)}/${record.id}`).set(deepClean(record));
  else await persistLocal();
  scheduleSnapshot();
}
// บันทึกหลาย record ครั้งเดียว (multi-path update — 1 round-trip)
export async function saveRecords(collKey: keyof DatabaseState, records: { id: string }[]): Promise<void> {
  if (!records.length) return;
  const fb = initFirebase();
  if (fb) {
    const upd: Record<string, any> = {};
    for (const r of records) upd[r.id] = deepClean(r);
    await fb.ref(`/${String(collKey)}`).update(upd);
  } else await persistLocal();
  scheduleSnapshot();
}
export async function removeRecord(collKey: keyof DatabaseState, id: string): Promise<void> {
  const fb = initFirebase();
  if (fb) await fb.ref(`/${String(collKey)}/${id}`).remove();
  else await persistLocal();
  scheduleSnapshot();
}
export async function removeRecords(collKey: keyof DatabaseState, ids: string[]): Promise<void> {
  if (!ids.length) return;
  const fb = initFirebase();
  if (fb) {
    const upd: Record<string, any> = {};
    for (const id of ids) upd[id] = null; // null ใน update = ลบ node นั้น
    await fb.ref(`/${String(collKey)}`).update(upd);
  } else await persistLocal();
  scheduleSnapshot();
}
// เขียนทั้งคอลเลกชัน (คอลเลกชันเล็กที่เปลี่ยนพร้อม trip เช่น cycles)
export async function flushCollection(collKey: keyof DatabaseState): Promise<void> {
  const fb = initFirebase();
  const arr = deepClean((cache as any)?.[collKey] || []); // clean ก่อน (ตัด undefined)
  if (fb) await fb.ref(`/${String(collKey)}`).set(ID_KEYED.includes(collKey) ? arrToMap(arr) : arr);
  else await persistLocal();
  scheduleSnapshot();
}
