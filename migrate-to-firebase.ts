// ============================================================================
// Migration ครั้งเดียว: ดันข้อมูลใน db.json ขึ้น Firebase Realtime Database
//
// วิธีใช้ (หลังวาง serviceAccountKey.json แล้ว):
//   npm run migrate:firebase           # ดันขึ้น (ถ้า Firebase มีข้อมูลอยู่จะหยุด)
//   npm run migrate:firebase -- --force  # เขียนทับ Firebase ด้วยข้อมูลจาก db.json
// ============================================================================
import fs from 'fs';
import path from 'path';
import { initializeApp, cert, getApps, type ServiceAccount } from 'firebase-admin/app';
import { getDatabase } from 'firebase-admin/database';
import { ensureShape } from './server-db.js';
import { DatabaseState } from './src/types.js';

const FORCE = process.argv.includes('--force');
const KEY_PATH = process.env.FIREBASE_KEY_PATH || path.join(process.cwd(), 'serviceAccountKey.json');
const DB_URL =
  process.env.FIREBASE_DB_URL ||
  'https://excel-billing-cycle-summarizer-default-rtdb.asia-southeast1.firebasedatabase.app';
const DB_FILE = path.join(process.cwd(), 'db.json');

async function main() {
  // 1) ตรวจ key
  if (!fs.existsSync(KEY_PATH)) {
    console.error(`❌ ไม่พบ ${KEY_PATH} — ดาวน์โหลด service account key แล้ววางก่อน`);
    process.exit(1);
  }
  // 2) อ่าน db.json + normalize
  if (!fs.existsSync(DB_FILE)) {
    console.error('❌ ไม่พบ db.json (ไม่มีข้อมูลให้ migrate)');
    process.exit(1);
  }
  const local: DatabaseState = ensureShape(JSON.parse(fs.readFileSync(DB_FILE, 'utf-8')));
  console.log(
    `📦 db.json: รอบ ${local.cycles.length}, รถ ${local.vehicles.length}, ราคา ${local.rateMasters.length}, ` +
      `ใบกระจาย ${local.tripDocuments.length}, ค่าน้ำมัน ${local.fuelEntries.length}, รายการหัก ${local.deductions.length}`
  );

  // 3) init Firebase
  const serviceAccount = JSON.parse(fs.readFileSync(KEY_PATH, 'utf-8')) as ServiceAccount;
  const app = getApps().length ? getApps()[0] : initializeApp({ credential: cert(serviceAccount), databaseURL: DB_URL });
  const fb = getDatabase(app);

  // 4) guard: Firebase มีข้อมูลอยู่แล้วไหม
  const snap = await fb.ref('/').once('value');
  const existing = snap.val() as Partial<DatabaseState> | null;
  const hasData = existing && ((existing.cycles?.length ?? 0) > 0 || (existing.tripDocuments?.length ?? 0) > 0);
  if (hasData && !FORCE) {
    console.error(
      `⚠️ Firebase มีข้อมูลอยู่แล้ว (รอบ ${existing!.cycles?.length ?? 0}, ใบกระจาย ${existing!.tripDocuments?.length ?? 0}) — ` +
        `หยุดเพื่อกันเขียนทับ ถ้าต้องการเขียนทับให้ใช้:  npm run migrate:firebase -- --force`
    );
    process.exit(1);
  }

  // 5) เขียนขึ้น Firebase (ตัด undefined)
  const clean = JSON.parse(JSON.stringify(local));
  await fb.ref('/').set(clean);
  console.log('✅ ดันข้อมูลขึ้น Firebase สำเร็จ — เปิดเซิร์ฟเวอร์ใหม่จะใช้ Firebase แล้ว');
  process.exit(0);
}

main().catch((e) => {
  console.error('❌ migrate ล้มเหลว:', e);
  process.exit(1);
});
