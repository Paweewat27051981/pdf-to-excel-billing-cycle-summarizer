// ============================================================================
// เลขบัญชีคนขับ + ใบปะสรุปยอดโอน (pure — ใช้ร่วม client/server)
// เจ้าของสั่ง 25 พ.ย.69: แก้ได้เฉพาะ HQ/admin · สาขาดู/พิมพ์ของตัวเองได้ · ห้ามหลุดทาง /api/state
// ============================================================================
import type { BankAccount, Vehicle, VehicleSummary } from './types';
import { normPlate, round2 } from './calc';

export const bankDigits = (s: any): string => String(s ?? '').replace(/\D/g, '');

// เลขบัญชีธนาคารไทย: 10 หลัก (ส่วนใหญ่) หรือ 12 หลัก (ออมสิน/ธ.ก.ส./ธอส.)
//   9/11 หลัก = เกือบแน่ว่าเลข 0 นำหน้าหายหรือพิมพ์ขาด -> บล็อก ห้ามเติมเอง
export function checkAccountNo(raw: any): { ok: true; digits: string } | { ok: false; error: string } {
  const s = String(raw ?? '').trim();
  //   จุดไม่ใช่ตัวคั่นเลขบัญชี: "123456789.0" (9 หลัก 0 หาย) ห้ามกลายเป็น 10 หลัก
  if (/[^\d\s\-–]/.test(s)) return { ok: false, error: 'เลขบัญชีมีตัวอักษรอื่นปน' };
  const digits = bankDigits(s);
  if (!digits) return { ok: false, error: 'ไม่มีเลขบัญชี' };
  if (digits.length !== 10 && digits.length !== 12) {
    return { ok: false, error: `เลขบัญชีมี ${digits.length} หลัก (ต้อง 10 หรือ 12 หลัก) — เลข 0 นำหน้าอาจหาย` };
  }
  if (/^0+$/.test(digits)) return { ok: false, error: 'เลขบัญชีเป็นศูนย์ทั้งหมด' };
  return { ok: true, digits };
}

// แสดงแบบใบปะ: 10 หลัก -> xxx-x-xxxxx-x · อื่น ๆ แสดงตัวเลขล้วน
export function fmtAccountNo(digits: string): string {
  const d = bankDigits(digits);
  return d.length === 10 ? `${d.slice(0, 3)}-${d.slice(3, 4)}-${d.slice(4, 9)}-${d.slice(9)}` : d;
}

// id ตายตัวต่อสาขา+ทะเบียน -> 1 คันมี 1 ระเบียนเสมอ (ห้ามมีตัวอักษรต้องห้ามของ Firebase key)
export const bankAccountId = (branchId: string, plateNo: string): string =>
  `bank-${branchId}-${normPlate(plateNo)}`.replace(/[.#$\[\]/\s]/g, '_');

export const normDriver = (s: string): string => (s || '').replace(/\s+/g, '');
// เทียบ "ชื่อในไฟล์ใบปะ" กับระบบเท่านั้น: ไม่สนคำนำหน้า ("สำเริง" = "นายสำเริง") — ไฟล์ทีมใส่บ้างไม่ใส่บ้าง
//   ห้ามใช้กับด่าน "คนขับเปลี่ยน" ของใบปะ (ชื่อในระบบทั้งคู่ ต้องเข้มเหมือนเดิม — เรื่องเงิน)
const TITLE = /^(นางสาว|นาง|นาย|น\.ส\.?|นส\.|(mrs|mr|ms|miss)(\.|\s)+)\s*/i;
const fileName = (s: string): string => normDriver((s || '').trim().replace(TITLE, ''));

// รถของสาขานี้ตามทะเบียน (active ก่อน)
export function findVehicle(vehicles: Vehicle[], branchId: string, plateNo: string): Vehicle | undefined {
  const k = normPlate(plateNo);
  if (!k) return undefined;
  const list = vehicles.filter((v) => v.branchId === branchId && normPlate(v.plateNo) === k);
  return list.find((v) => v.status === 'active') || list[0];
}

export interface PayslipRow { plateNo: string; driverName: string; accountNo: string; net: number; note: string }
export interface PayslipHold { plateNo: string; driverName: string; reason: string; net: number }

// แยก "โอนได้" กับ "ยังโอนไม่ได้" — สองกลุ่มรวมกันต้องเท่ารับสุทธิทุกคันเสมอ (ห้ามมีคันหายเงียบ)
export function buildPayslip(
  sums: VehicleSummary[], accounts: BankAccount[], branchId: string,
): { pay: PayslipRow[]; hold: PayslipHold[]; payTotal: number; holdTotal: number } {
  const byPlate = new Map<string, BankAccount>();
  for (const a of accounts) if (a.branchId === branchId && a.accountNo) byPlate.set(normPlate(a.plateNo), a);
  const pay: PayslipRow[] = [];
  const hold: PayslipHold[] = [];
  for (const s of sums) {
    const a = byPlate.get(normPlate(s.plateNo));
    const base = { plateNo: s.plateNo, driverName: s.driverName, net: s.netReceive };
    if (!a) hold.push({ ...base, reason: 'ยังไม่มีเลขบัญชี' });
    else if (normDriver(a.driverName) !== normDriver(s.driverName)) {
      // คนขับในทะเบียนรถเปลี่ยนหลังผูกเลขบัญชี -> เลขอาจเป็นของคนเก่า ให้ HQ ยืนยันก่อน
      hold.push({ ...base, reason: `คนขับเปลี่ยนจาก "${a.driverName || '-'}" — ให้ HQ ยืนยันเลขบัญชีก่อนโอน` });
    } else if (s.netReceive <= 0) hold.push({ ...base, reason: 'รับสุทธิไม่เกิน 0 — ไม่ต้องโอน' });
    else pay.push({ ...base, accountNo: fmtAccountNo(a.accountNo), note: a.note || '' });
  }
  const sum = (xs: { net: number }[]) => round2(xs.reduce((t, x) => t + x.net, 0));
  return { pay, hold, payTotal: sum(pay), holdTotal: sum(hold) };
}

// ---------------------------------------------------------------------------
// นำเข้าจากไฟล์ใบปะ — แถวดิบ (server อ่าน Excel แล้ว) -> ผลเทียบก่อนบันทึก
// ---------------------------------------------------------------------------
export interface BankSheetRow { sheet: string; row: number; plate: string; driver: string; account: string; note: string; numeric: boolean }
export type BankImportStatus = 'new' | 'changed' | 'same' | 'error';
export interface BankImportRow extends BankSheetRow {
  status: BankImportStatus; error?: string; warnings: string[]; digits: string; oldAccount: string; systemDriver: string; newNote: string;
}

const FILE_NAME_NOTE = 'ชื่อในใบปะ: ';

export function planBankImport(
  rows: BankSheetRow[], vehicles: Vehicle[], accounts: BankAccount[], branchId: string,
): { rows: BankImportRow[]; blocked: string; skippedSheets: string[] } {
  // ไฟล์รวมหลายสาขา (1 ชีตต่อสาขา เช่น "สรุปยอดจ่าย - รถร่วมภาคเหนือ") -> ใช้เฉพาะชีตที่ทะเบียนส่วนใหญ่อยู่สาขานี้
  //   ชีตเดียว = กติกาเดิม (ทะเบียนนอกสาขาเกินครึ่ง = บล็อกทั้งไฟล์ กันนำเข้าผิดสาขา)
  const sheets = [...new Set(rows.map((r) => r.sheet))];
  const localOf = (sh: string) => {
    const ps = [...new Set(rows.filter((r) => r.sheet === sh).map((r) => normPlate(r.plate)).filter(Boolean))];
    return { all: ps.length, local: ps.filter((p) => findVehicle(vehicles, branchId, p)).length };
  };
  const skippedSheets = sheets.length < 2 ? [] : sheets.filter((sh) => { const c = localOf(sh); return c.local * 2 < c.all; });
  if (skippedSheets.length && skippedSheets.length === sheets.length) {
    return { rows: [], skippedSheets, blocked: `ไม่มีชีตของสาขานี้ในไฟล์ (ทุกชีตเป็นทะเบียนสาขาอื่น: ${skippedSheets.join(', ')}) — เลือกสาขาให้ถูกก่อนนำเข้า` };
  }
  const skippedAll = skippedSheets.reduce((t, sh) => t + localOf(sh).all, 0); // นับก่อนกรองแถวทิ้ง
  rows = rows.filter((r) => !skippedSheets.includes(r.sheet));
  // ไฟล์ส่วนใหญ่เป็นของสาขาอื่น แต่มีชีตเล็กที่ทะเบียนตรงสาขานี้ (เช่นรถที่ลงทะเบียน 2 สาขา) -> เตือนทุกแถวให้ตรวจว่าเลือกสาขาถูก
  const keptLocal = new Set(rows.map((r) => normPlate(r.plate)).filter((p) => p && findVehicle(vehicles, branchId, p))).size;
  const wrongBranch = skippedSheets.length > 0 && keptLocal < 3 && skippedAll > keptLocal
    ? `ไฟล์นี้ส่วนใหญ่เป็นทะเบียนสาขาอื่น (${skippedAll} คัน) มีของสาขานี้แค่ ${keptLocal} คัน — ตรวจว่าเลือกสาขาถูก` : '';
  const digitsByPlate = new Map<string, Set<string>>();
  const namesByPlate = new Map<string, Set<string>>();
  for (const r of rows) {
    const k = normPlate(r.plate); if (!k) continue;
    const set = digitsByPlate.get(k) || new Set<string>(); set.add(bankDigits(r.account)); digitsByPlate.set(k, set);
    const names = namesByPlate.get(k) || new Set<string>(); names.add(fileName(r.driver)); namesByPlate.set(k, names);
  }
  const out: BankImportRow[] = [];
  const seen = new Set<string>();
  for (const r of rows) {
    const k = normPlate(r.plate);
    if (!k) continue;
    const v = findVehicle(vehicles, branchId, r.plate);
    const cur = accounts.find((a) => a.branchId === branchId && normPlate(a.plateNo) === k);
    const base: BankImportRow = {
      ...r, status: 'error', warnings: [], digits: '', oldAccount: cur?.accountNo || '', systemDriver: v?.driverName || '', newNote: '',
    };
    if (!v) { out.push({ ...base, error: 'ไม่มีทะเบียนนี้ในสาขา' }); continue; }
    if ((digitsByPlate.get(k)?.size || 0) > 1) { out.push({ ...base, error: 'ทะเบียนนี้มีเลขบัญชีมากกว่า 1 เลขในไฟล์' }); continue; }
    if (seen.has(k)) continue; // ซ้ำแต่เลขเดียวกัน -> เอาแถวแรก
    seen.add(k);
    const c = checkAccountNo(r.account);
    if (c.ok === false) { out.push({ ...base, error: c.error }); continue; }
    if (r.numeric) base.warnings.push('เซลล์เก็บเป็นตัวเลข เลข 0 นำหน้าอาจหาย — ตรวจอีกครั้ง');
    if (wrongBranch) base.warnings.push(wrongBranch);
    // เจ้าของเคาะ 25 พ.ย.69: บัญชีไม่ต้องเป็นชื่อเดียวกับคนขับ -> ชื่อไม่ตรง = เตือนแต่บันทึกได้ (จดชื่อในไฟล์ไว้เป็นหมายเหตุ)
    //   "ยศพล สวนม่วง (พิไลพร)" = คนขับ ยศพล โอนเข้าชื่อ พิไลพร
    const fileDriver = (r.driver || '').replace(/\s+/g, ' ').trim();
    const paren = /[(（]([^)）]*)[)）]/.exec(fileDriver)?.[1]?.trim() || '';
    const bare = fileDriver.replace(/[(（][^)）]*[)）]/g, '').trim();
    const nameDiff = !!bare && fileName(bare) !== fileName(v.driverName);
    // ข้อยกเว้นเดียวที่ยังกัน: เคยผูกเลขแล้ว คนขับในระบบเปลี่ยนทีหลัง และไฟล์ไม่ได้เป็นชื่อคนใหม่
    //   ไฟล์อาจเก่า (รวมเลขที่ HQ ลบ/แก้ไปแล้ว) ถ้าปล่อยผ่าน = ยืนยันเลขคนเก่าให้คนใหม่โดยไม่มีใครตรวจ
    if (cur && normDriver(cur.driverName) !== normDriver(v.driverName) && (nameDiff || !bare)) {
      out.push({ ...base, error: `คนขับในระบบเปลี่ยนจาก "${cur.driverName || '-'}" เป็น "${v.driverName || '-'}" หลังผูกเลขบัญชี เงินอาจเข้าคนเก่า — กด ✎ "แก้เลขบัญชี" ในคอลัมน์เลขบัญชี ตรวจแล้วบันทึกยืนยัน` }); continue;
    }
    if (nameDiff) base.warnings.push(`ชื่อในไฟล์ "${fileDriver}" ไม่ตรงกับคนขับในระบบ "${v.driverName || '-'}"`);
    if (!bare) base.warnings.push('ไม่มีชื่อคนขับในไฟล์ให้เทียบ');
    if ((namesByPlate.get(k)?.size || 0) > 1) base.warnings.push('ทะเบียนนี้อยู่หลายแถวในไฟล์และชื่อไม่เหมือนกัน — ใช้แถวแรก');
    if (cur && !cur.accountNo) base.warnings.push('เคยมีคนลบเลขบัญชีของคันนี้ออก — ตรวจก่อนบันทึก');
    // หมายเหตุ = ข้อความคนพิมพ์ + ชื่อที่ระบบจดเอง
    //   · ข้อความจากไฟล์มาก่อน · คงของเดิมเฉพาะเมื่อเลขและคนขับไม่เปลี่ยน (เลขใหม่ห้ามติดหมายเหตุเลขเก่า)
    //   · ส่วนที่ระบบจดเอง ("ชื่อในใบปะ: ...") คิดใหม่จากไฟล์ทุกครั้ง ไม่ค้างชื่อเก่า
    const unchanged = !!cur && cur.accountNo === c.digits && normDriver(cur.driverName) === normDriver(v.driverName);
    const human = (t: string) => t.split(' · ').filter((x) => x && !x.startsWith(FILE_NAME_NOTE)).join(' · ');
    const carried = !unchanged && !!cur?.note && r.note === cur.note; // ใบปะที่ export ไปแล้วแก้แต่เลข = หมายเหตุของเลขเก่า
    const auto = nameDiff ? FILE_NAME_NOTE + fileDriver : paren ? `โอนเข้าชื่อ ${paren}` : '';
    const text = r.note && !carried ? (nameDiff ? human(r.note) : r.note) : unchanged ? human(cur!.note || '') : '';
    const note = [text, auto && !text.includes(auto) ? auto : ''].filter(Boolean).join(' · ').slice(0, 200);
    const same = unchanged && (cur!.note || '') === note;
    out.push({ ...base, digits: c.digits, newNote: note, status: same ? 'same' : cur?.accountNo ? 'changed' : 'new' });
  }
  // แก้เลขในช่องที่ผสานหลายคัน = ทุกคันในกลุ่มเปลี่ยนตาม -> เตือนให้ตรวจว่าตั้งใจเปลี่ยนทุกคันจริง
  for (const r of out) {
    if (r.status !== 'changed' || bankDigits(r.oldAccount) === r.digits) continue;
    const n = out.filter((x) => x.status === 'changed' && x.digits === r.digits && bankDigits(x.oldAccount) !== x.digits).length;
    if (n > 1) r.warnings.push(`เลขบัญชีนี้ถูกเปลี่ยนพร้อมกัน ${n} คัน — ตรวจว่าตั้งใจเปลี่ยนทุกคัน`);
  }
  const miss = out.filter((r) => r.error === 'ไม่มีทะเบียนนี้ในสาขา').length;
  // ชื่อไม่ตรงรายคันอนุโลม แต่ถ้าไม่ตรงเกินครึ่งไฟล์ (4 แถวขึ้นไป) = น่าจะเรียง/เลื่อนแถวผิด ทุกคันจะโอนสลับคน
  const diff = out.filter((r) => r.systemDriver && r.warnings.some((w) => w.startsWith('ชื่อในไฟล์'))).length;
  const blocked = out.length && miss * 2 > out.length
    ? `ทะเบียนในไฟล์ไม่อยู่ในสาขานี้ ${miss}/${out.length} แถว — น่าจะเป็นไฟล์ของสาขาอื่น ไม่นำเข้าทั้งไฟล์`
    : diff >= 4 && diff * 2 > out.length
      ? `ชื่อคนขับไม่ตรงระบบ ${diff}/${out.length} แถว — ไฟล์น่าจะเรียงแถวสลับกัน ตรวจว่าทะเบียน/ชื่อ/เลขบัญชีอยู่แถวเดียวกัน ไม่นำเข้าทั้งไฟล์` : '';
  return { rows: out, blocked, skippedSheets };
}
