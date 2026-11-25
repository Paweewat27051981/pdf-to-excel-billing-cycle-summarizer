// ทดสอบเลขบัญชีคนขับ/ใบปะ/นำเข้า + ไฟล์ Excel ที่ export จริง — รัน: npx tsx src/bankAccounts.test.mts
// เลขบัญชีในไฟล์นี้เป็นเลขสมมติทั้งหมด (ห้ามใส่เลขจริงลงโค้ด)
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { checkAccountNo, fmtAccountNo, bankAccountId, buildPayslip, planBankImport } from './bankAccounts.js';
import { parseBankSheets } from '../bank-import.js';
import { exportPerVehicleReport } from './excel-export.js';

let pass = 0, fail = 0;
const check = (name: string, got: any, want: any) => {
  const ok = JSON.stringify(got) === JSON.stringify(want); ok ? pass++ : fail++;
  console.log(`${ok ? '✓' : '✗ FAIL'} ${name}: ได้ ${JSON.stringify(got)}${ok ? '' : ` (คาด ${JSON.stringify(want)})`}`);
};

// --- ตรวจเลขบัญชี
check('10 หลักมีขีด', checkAccountNo('012-3-45678-9'), { ok: true, digits: '0123456789' });
check('12 หลัก', checkAccountNo('020123456789').ok, true);
for (const bad of ['123456789', '12345678901', '', '-', 'กสิกร 0123456789', '0000000000', '1.12E+09', '123456789.0', '1234567890.00']) check(`บล็อก "${bad}"`, checkAccountNo(bad).ok, false);
check('แสดงผล', [fmtAccountNo('0123456789'), fmtAccountNo('020123456789')], ['012-3-45678-9', '020123456789']);
check('id ปลอดภัยกับ Firebase', /[.#$\[\]/\s]/.test(bankAccountId('br-x', '3ฒฎ-745 ตู้.#')), false);

// --- ใบปะ: ห้ามมีคันหาย
const S = (plateNo: string, driverName: string, netReceive: number) => ({ plateNo, driverName, netReceive, totalTripAmount: netReceive, lines: [] }) as any;
const A = (plateNo: string, accountNo: string, driverName: string, note = '', branchId = 'b') => ({ id: plateNo, branchId, plateNo, accountNo, driverName, note, history: [] }) as any;
const sums = [S('A-1', 'นาย ก', 100.1), S('B-2', 'นาย ข', 200.2), S('C-3', 'นาย ใหม่', 300.3), S('D-4', 'นาย ง', -50), S('E-5', 'นาย จ', 10)];
const ps = buildPayslip(sums, [A('A-1', '0123456789', 'นาย  ก', 'โอนเป็นคุณ x'), A('C-3', '0123456780', 'นาย เก่า'), A('D-4', '0123456781', 'นาย ง'), A('E-5', '', 'นาย จ'), A('B-2', '9999999999', 'นาย ข', '', 'สาขาอื่น')], 'b');
check('โอนได้เฉพาะคันที่ครบ', ps.pay.map((p) => [p.plateNo, p.accountNo, p.note]), [['A-1', '012-3-45678-9', 'โอนเป็นคุณ x']]);
check('กลุ่มยังไม่โอน', ps.hold.map((h) => h.plateNo), ['B-2', 'C-3', 'D-4', 'E-5']);
check('คนขับเปลี่ยน = กันไว้', /คนขับเปลี่ยน/.test(ps.hold[1].reason), true);
check('โอน + ยังไม่โอน = รับสุทธิทุกคัน', Math.round((ps.payTotal + ps.holdTotal) * 100) / 100, 560.6);

// --- อ่านไฟล์ใบปะ (โครงตามไฟล์ทีม: หัวแถว 4, เซลล์ผสาน, หมายเหตุท้ายแถวไม่มีหัว, แถวรวม)
const ws = XLSX.utils.aoa_to_sheet([
  ['ใบปะ — สรุปยอดโอนเข้าบัญชีคนขับ สาขาทดสอบ'], ['รอบ ต.ค. 69'], [],
  ['ทะเบียน', 'คนขับ', 'เลขบัญชี', 'รับสุทธิ'],
  ['A-1', 'นาย ก', '012-3-45678-9', 1], ['B-2', 'นาย ข', '', 2], ['C-3', 'นาย ค (คุณ รับแทน)', '0508000001', 3, 'โทรถามแล้ว ให้โอนเป็นคุณรับแทน'],
  ['D-4', 'นาย ง', 391690763, 4], ['E-5', 'นาย จ', '', 5], ['F-6', 'คนละคน', '0123456784', 6], ['G-7', 'นาย ช', 1128000001, 7], ['H-8', 'นาย ซ', '', 8],
  ['รวมทุกคัน', '', '', 36], [], ['ทะเบียน', 'คนขับ', 'ยังไม่โอน เพราะ', 'รับสุทธิ'], ['Z-9', 'x', '0123456785', 1],
]);
ws['!merges'] = [{ s: { r: 4, c: 2 }, e: { r: 5, c: 2 } }];
const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['ทะเบียน', 'รายได้'], ['A-1', 5]]), 'อื่น'); XLSX.utils.book_append_sheet(wb, ws, 'ใบปะ');
const rows = parseBankSheets(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
check('อ่านแถว (ข้ามว่าง/แถวรวม/ตารางล่าง)', rows.map((r) => r.plate), ['A-1', 'B-2', 'C-3', 'D-4', 'F-6', 'G-7']);
check('เซลล์ผสานได้เลขเดียวกัน', rows[1].account, '012-3-45678-9');
check('หมายเหตุท้ายแถวไม่มีหัวคอลัมน์', rows[2].note, 'โทรถามแล้ว ให้โอนเป็นคุณรับแทน');
check('เลข 0 นำหน้าในข้อความไม่หาย', rows[2].account, '0508000001');
const xls = parseBankSheets(XLSX.write(wb, { type: 'buffer', bookType: 'biff8' }));
check('.xls ได้ผลเท่ากัน', xls.map((r) => r.plate + r.account), rows.map((r) => r.plate + r.account));
const V = (plateNo: string, driverName: string) => ({ id: plateNo, branchId: 'b', plateNo, driverName, status: 'active' }) as any;
const vs = [V('A-1', 'นาย ก'), V('B-2', 'นาย ข'), V('C-3', 'นาย ค'), V('D-4', 'นาย ง'), V('F-6', 'นาย ฉ'), V('G-7', 'นาย ช')];
const plan = planBankImport(rows, vs, [A('A-1', '0123456789', 'นาย ก'), A('B-2', '1111111111', 'นาย ข')], 'b');
check('สถานะ', plan.rows.map((r) => r.plate + ':' + r.status), ['A-1:same', 'B-2:changed', 'C-3:new', 'D-4:error', 'F-6:error', 'G-7:new']);
check('เลขตัวเลข 9 หลัก บล็อก (0 หาย)', /9 หลัก/.test(plan.rows[3].error || ''), true);
check('ชื่อคนขับไม่ตรง บล็อก', /ไม่ตรงกับคนขับ/.test(plan.rows[4].error || ''), true);
check('เซลล์ตัวเลข 10 หลัก เตือน', plan.rows[5].warnings.length, 1);
check('ไฟล์สาขาอื่น บล็อกทั้งไฟล์', !!planBankImport(rows, [vs[0]], [], 'b').blocked, true);
check('ไม่บล็อกไฟล์ปกติ', plan.blocked, '');
const dup = planBankImport([{ ...rows[0] }, { ...rows[0], account: '0123456788' }], vs, [], 'b');
check('ทะเบียนเดียวสองเลข บล็อกทั้งคู่', dup.rows.map((r) => r.status), ['error', 'error']);
check('รถสาขาอื่นไม่ถูกจับ', planBankImport([rows[0]], [{ ...vs[0], branchId: 'x' }], [], 'b').rows[0].status, 'error');
check('คนขับเปลี่ยน + ไฟล์เก่า = บล็อก', planBankImport([rows[0]], [V('A-1', 'คนใหม่')], [A('A-1', '0123456789', 'นาย ก')], 'b').rows[0].status, 'error');

check('ไม่มีชื่อคนขับในไฟล์ = บล็อก (ห้ามยืนยันคนใหม่ให้เลขเก่า)', planBankImport([{ ...rows[0], driver: '' }], [V('A-1', 'คนใหม่')], [A('A-1', '0123456789', 'นาย ก')], 'b').rows[0].status, 'error');
check('เลขใหม่ไม่ติดหมายเหตุผู้รับโอนเดิม', planBankImport([rows[0]], vs, [A('A-1', '0999999999', 'นาย ก', 'โอนเข้าชื่อ เก่า')], 'b').rows[0].newNote, '');
check('เลขเดิมคนเดิม คงหมายเหตุ', planBankImport([rows[0]], vs, [A('A-1', '0123456789', 'นาย ก', 'โอนเข้าชื่อ เก่า')], 'b').rows[0].status, 'same');
check('ชื่อในวงเล็บเป็นหมายเหตุ', planBankImport([{ ...rows[0], driver: 'นาย ก (ข รับแทน)' }], vs, [], 'b').rows[0].newNote, 'โอนเข้าชื่อ ข รับแทน');
const ws2 = XLSX.utils.aoa_to_sheet([['ทะเบียน', 'คนขับ', 'เลขบัญชี', 'รับสุทธิ'], ['A-1', 'ก', '0123456789', 1, 'โอนรวม'], ['B-2', 'ข', '', 1]]);
ws2['!merges'] = [{ s: { r: 1, c: 2 }, e: { r: 2, c: 2 } }, { s: { r: 1, c: 4 }, e: { r: 2, c: 4 } }];
const wb2 = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb2, ws2, 's');
check('หมายเหตุผสานหลายแถว', parseBankSheets(XLSX.write(wb2, { type: 'buffer', bookType: 'xlsx' })).map((r) => r.note), ['โอนรวม', 'โอนรวม']);

// --- ไฟล์ Excel ที่ export จริง (ดักไฟล์ดาวน์โหลด)
let blob: Blob | null = null;
(globalThis as any).document = { createElement: () => ({ click() {} }) };
(URL as any).createObjectURL = (b: Blob) => { blob = b; return 'x'; }; (URL as any).revokeObjectURL = () => {};
(globalThis as any).fetch = async () => ({ ok: false, headers: { get: () => '' } });
const T = (no: string, plate: string, amt: number, qty: number, d: string, p: string) => ({ id: no, cycleId: 'c', branchId: 'b', documentNo: no, documentDate: '2026-10-01', plateNo: plate, tripAmount: amt, billingQty: qty, totalQty: qty + 9, rateType: 'flat', rateValue: amt, breakdown: {}, receipts: [{ districtRaw: d, provinceRaw: p }] }) as any;
const vx = [V('A-1', 'นาย ก'), V('B-2', 'นาย ข')];
await exportPerVehicleReport({ id: 'c', name: 'ต.ค. 69' } as any, 'ทดสอบ', [T('J1', 'A-1', 1000, 100, 'เมือง', 'ตาก'), T('J2', 'A-1', 500, 20, 'แม่สอด', 'ตาก'), T('J3', 'B-2', 300, 30, 'เมือง', 'ตาก')], [], [], vx, undefined, [], {
  splitZoneText: 'ตาก: แม่สอด', logoUrl: 'x', bankAccounts: [A('A-1', '0123456789', 'นาย ก', 'หมายเหตุ1')], branchId: 'b',
});
const out = new ExcelJS.Workbook(); await out.xlsx.load(await (blob as any).arrayBuffer());
const dump = (n: string) => { const o: any[] = []; out.getWorksheet(n)!.eachRow((r) => o.push((r.values as any[]).slice(1))); return o; };
const sum = dump('สรุปรวม');
check('ตารางหลัก', sum.filter((r) => r[0] === 'A-1')[0].slice(0, 5), ['A-1', 'นาย ก', 100, 10, 1000]);
check('ตารางพื้นที่ + หมายเหตุ', sum.filter((r) => r[0] === 'A-1')[1], ['A-1', 'นาย ก', 20, 25, 500, 5, 0, 0, 0, 495, 'แม่สอด']);
check('แถวรวม = รายได้รวม ÷ ชิ้นรวม', sum.filter((r) => r[0] === 'รวมทุกคัน').map((r) => [r[2], r[3]]), [[130, 10], [20, 25]]);
const pay = dump('ใบปะ');
check('ใบปะ รวมยอดทั้งคัน เป็นข้อความ', pay.find((r) => r[0] === 'A-1'), ['A-1', 'นาย ก', '012-3-45678-9', 1485, 'หมายเหตุ1']);
check('ใบปะ ยอดรวม', pay.filter((r) => /^รวม/.test(String(r[0]))).map((r) => [r[0], r[3]]), [['รวมยอดโอน', 1485], ['รวมที่ยังไม่โอน', 297], ['รวมทั้งหมด (เท่ากับรับสุทธิในชีตสรุปรวม)', 1782]]);
const round = parseBankSheets(Buffer.from(await (blob as any).arrayBuffer()));
check('ไฟล์ที่ export นำเข้ากลับได้', round.map((r) => [r.plate, r.account, r.note]), [['A-1', '012-3-45678-9', 'หมายเหตุ1']]);
await exportPerVehicleReport({ id: 'c', name: 'x' } as any, 'ทดสอบ', [T('J1', 'A-1', 1000, 100, 'แม่สอด', 'ตาก')], [], [], vx, undefined, []);
const out2 = new ExcelJS.Workbook(); await out2.xlsx.load(await (blob as any).arrayBuffer());
check('ไม่ตั้งค่า = ไม่มีใบปะ/ไม่แยกตาราง', [!!out2.getWorksheet('ใบปะ'), out2.getWorksheet('สรุปรวม')!.rowCount], [false, 6]);

console.log(`\nผ่าน ${pass}/${pass + fail}`);
if (fail) process.exit(1);
