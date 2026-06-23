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
    const app = getApps().length
      ? getApps()[0]
      : initializeApp({ credential: cert(serviceAccount), databaseURL: FIREBASE_DB_URL });
    firebaseDb = getDatabase(app);
    console.log('✅ ใช้ Firebase Realtime Database (cache + write-through)');
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
    rateMasterHistory: [],
    receiverGroups: [
      { id: 'grp-macro', groupName: 'แม็คโคร/เซลส์', status: 'active' },
    ],
    receiverGroupAliases: [
      { id: 'al-1', receiverGroupId: 'grp-macro', aliasName: 'แม็คโคร', status: 'active' },
      { id: 'al-2', receiverGroupId: 'grp-macro', aliasName: 'MK', status: 'active' },
      { id: 'al-3', receiverGroupId: 'grp-macro', aliasName: 'CP AXTRA', status: 'active' },
      { id: 'al-4', receiverGroupId: 'grp-macro', aliasName: 'ซีพี แอ็กซ์ตร้า', status: 'active' },
      { id: 'al-5', receiverGroupId: 'grp-macro', aliasName: 'เซลส์', status: 'active' },
    ],
    conversionRules: [
      {
        id: 'rule-yupi', ruleName: 'ยูปี้ 14 กรัม หาร 3', senderKeyword: 'ซีโน', receiverGroupId: 'grp-macro',
        productKeyword: 'ยูปี้', productSizeKeyword: '14 กรัม', divisor: 3, roundingMethod: 'half_up',
        applyLevel: 'receipt', status: 'active', effectiveFrom: '2026-05-01', effectiveTo: null,
        remark: 'เฉพาะงานซีโนที่ส่งแม็คโคร/เซลส์',
      },
      {
        id: 'rule-pringles', ruleName: 'พริงเกิล 42 กรัม หาร 3', senderKeyword: 'ซีโน', receiverGroupId: 'grp-macro',
        productKeyword: 'พริงเกิล', productSizeKeyword: '42 กรัม', divisor: 3, roundingMethod: 'half_up',
        applyLevel: 'receipt', status: 'active', effectiveFrom: '2026-05-01', effectiveTo: null,
        remark: 'ร้านค้าปกติไม่หาร',
      },
    ],
    manualBoxSenders: [
      { id: 'mbs-cpconsumer', senderKeyword: 'คอนซูเมอร์', note: 'ซ.พี.คอนซูเมอร์ โพรดักส์ — ส่งเป็นชิ้น ต้องกรอกจำนวนกล่องเอง', status: 'active' },
    ],
    moneyCategories: defaultMoneyCategories(),
    tripDocuments: [],
    fuelEntries: [],
    deductions: [],
  };
}

// ประเภทรายการเงินตั้งต้น (income = รายได้เพิ่ม, deduction = หักออก)
function defaultMoneyCategories() {
  return [
    { id: 'cat-bill_update', name: 'ค่าอัพเดทบิล', kind: 'income' as const, status: 'active' as const, builtin: true },
    { id: 'cat-phone', name: 'ค่าโทรศัพท์', kind: 'deduction' as const, status: 'active' as const, builtin: true },
    { id: 'cat-gps', name: 'ค่า GPS', kind: 'deduction' as const, status: 'active' as const, builtin: true },
    { id: 'cat-loan', name: 'ยืมเงิน', kind: 'deduction' as const, status: 'active' as const, builtin: true },
    { id: 'cat-gps_yearly', name: 'GPS รายปี', kind: 'deduction' as const, status: 'active' as const, builtin: true },
    { id: 'cat-insurance', name: 'ประกัน', kind: 'deduction' as const, status: 'active' as const, builtin: true },
    { id: 'cat-other', name: 'อื่นๆ', kind: 'deduction' as const, status: 'active' as const, builtin: true },
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

// migrate: เติม key ที่ขาดให้ db เก่า
export function ensureShape(state: Partial<DatabaseState>): DatabaseState {
  const seed = seedState();
  return {
    settings: { ...seed.settings, ...(state.settings || {}) },
    branches: state.branches && state.branches.length ? state.branches : seed.branches,
    cycles: state.cycles ?? [],
    vehicles: withBranch(state.vehicles, seed.vehicles),
    rateMasters: withBranch(state.rateMasters, seed.rateMasters),
    rateMasterHistory: state.rateMasterHistory ?? [],
    receiverGroups: state.receiverGroups ?? seed.receiverGroups,
    receiverGroupAliases: state.receiverGroupAliases ?? seed.receiverGroupAliases,
    conversionRules: state.conversionRules ?? seed.conversionRules,
    manualBoxSenders: state.manualBoxSenders ?? seed.manualBoxSenders,
    moneyCategories: state.moneyCategories ?? seed.moneyCategories,
    tripDocuments: withBranch(state.tripDocuments, []),
    fuelEntries: withBranch(state.fuelEntries, []),
    deductions: withBranch(migrateDeductions(state.deductions as any[]), []),
  };
}

export async function getDb(): Promise<DatabaseState> {
  // 1) ถ้ามี cache อยู่แล้ว -> ตอบทันที (ไม่อ่าน Firebase = ไม่มี download)
  if (cache) return cache;

  const fb = initFirebase();
  if (fb) {
    // 2) Firebase: อ่านครั้งเดียวตอนบูต -> เก็บ cache
    const snap = await fb.ref('/').once('value');
    const val = snap.val() as Partial<DatabaseState> | null;
    cache = ensureShape(val || {});
    if (!val) await saveDb(cache); // ว่างเปล่า -> seed ขึ้น Firebase
    return cache;
  }

  // 3) Fallback: db.json (เมื่อยังไม่ใส่ serviceAccountKey)
  try {
    const content = await fs.readFile(DB_FILE, 'utf-8');
    cache = ensureShape(JSON.parse(content) as Partial<DatabaseState>);
    return cache;
  } catch {
    cache = seedState();
    await saveDb(cache);
    return cache;
  }
}

export async function saveDb(state: DatabaseState): Promise<void> {
  cache = state; // อัปเดต cache ในหน่วยความจำ
  // ตัด undefined ออก (RTDB ไม่รับ undefined) ด้วยการ round-trip JSON
  const clean = JSON.parse(JSON.stringify(state));
  const fb = initFirebase();
  if (fb) {
    await fb.ref('/').set(clean); // write-through (การเขียนไม่ถูกคิดเป็น download)
  } else {
    await fs.writeFile(DB_FILE, JSON.stringify(clean, null, 2), 'utf-8');
  }
}
