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
check('สถานะ', plan.rows.map((r) => r.plate + ':' + r.status), ['A-1:same', 'B-2:changed', 'C-3:new', 'D-4:error', 'F-6:new', 'G-7:new']);
check('เลขตัวเลข 9 หลัก บล็อก (0 หาย)', /9 หลัก/.test(plan.rows[3].error || ''), true);
check('ชื่อคนขับไม่ตรง = เตือน + จดชื่อในใบปะ ไม่บล็อก (เจ้าของเคาะ 25 พ.ย.69)', [plan.rows[4].error, /ไม่ตรงกับคนขับ/.test(plan.rows[4].warnings.join()), plan.rows[4].newNote], [undefined, true, 'ชื่อในใบปะ: คนละคน']);
check('ใบปะชื่อไม่ตรงอยู่แล้ว + คนขับเปลี่ยนทีหลัง + ไฟล์เดิม = กันไว้', planBankImport([rows[4]], [V('F-6', 'คนใหม่')], [A('F-6', '0123456784', 'นาย ฉ', 'ชื่อในใบปะ: คนละคน')], 'b').rows[0].status, 'error');
check('คนขับเปลี่ยน + ไฟล์ชื่อคนใหม่ เลขเดิม = ยืนยันได้', planBankImport([{ ...rows[4], driver: 'คนใหม่' }], [V('F-6', 'คนใหม่')], [A('F-6', '0123456784', 'นาย ฉ')], 'b').rows[0].status, 'changed');
check('เลขผิดหลักยังบล็อก แม้ชื่อไม่ตรง', planBankImport([{ ...rows[4], account: '12345' }], vs, [], 'b').rows[0].status, 'error');
check('ผูกแล้ว คนขับเปลี่ยน + ไฟล์ชื่อคนที่สาม = กันไว้', planBankImport([{ ...rows[0], driver: 'คนที่สาม', account: '0555555555' }], [V('A-1', 'คนใหม่')], [A('A-1', '0123456789', 'นาย ก')], 'b').rows[0].status, 'error');
check('HQ ลบเลขแล้ว + คนขับเปลี่ยน + ไฟล์เก่า = กันไว้', planBankImport([rows[0]], [V('A-1', 'คนใหม่')], [A('A-1', '', 'นาย ก')], 'b').rows[0].status, 'error');
check('HQ ลบเลขแล้ว คนขับเดิม = เตือน', planBankImport([rows[0]], vs, [A('A-1', '', 'นาย ก')], 'b').rows[0].warnings.length, 1);
const P = (r: any, note: string, acc = '0123456789') => planBankImport([{ ...rows[0], ...r }], vs, [A('A-1', acc, 'นาย ก', note)], 'b').rows[0];
check('แก้ชื่อในไฟล์ถูกแล้ว หมายเหตุชื่อเดิมหายไป', [P({}, 'ชื่อในใบปะ: คนละคน').status, P({}, 'ชื่อในใบปะ: คนละคน').newNote], ['changed', '']);
check('หมายเหตุคนพิมพ์อยู่ ส่วนที่ระบบจดหาย', P({}, 'โทรแล้ว · ชื่อในใบปะ: คนละคน').newNote, 'โทรแล้ว');
check('หมายเหตุในไฟล์ + ชื่อไม่ตรง เก็บทั้งคู่', P({ driver: 'คนที่สาม', note: 'โทรแล้ว' }, '').newNote, 'โทรแล้ว · ชื่อในใบปะ: คนที่สาม');
check('ใบปะ export แล้วแก้แต่เลข ไม่ติดหมายเหตุเลขเก่า', P({ account: '0555555555', note: 'ชื่อในใบปะ: คนละคน' }, 'ชื่อในใบปะ: คนละคน').newNote, '');
check('ใบปะ export แล้วนำเข้ากลับ = same', P({ note: 'ชื่อในใบปะ: คนละคน' }, 'ชื่อในใบปะ: คนละคน').status, 'same');
check('นำเข้าซ้ำ ชื่อไม่ตรง+หมายเหตุ = same พร้อมเตือน', [P({ driver: 'คนที่สาม', note: 'โทรแล้ว' }, 'โทรแล้ว · ชื่อในใบปะ: คนที่สาม').status, P({ driver: 'คนที่สาม', note: 'โทรแล้ว' }, 'โทรแล้ว · ชื่อในใบปะ: คนที่สาม').warnings.length], ['same', 1]);
check('ทะเบียนซ้ำชื่อต่างกัน = เตือน', planBankImport([rows[0], { ...rows[0], driver: 'อีกคน' }], vs, [], 'b').rows[0].warnings.length, 1);
const shift = (n: number) => planBankImport(Array.from({ length: 6 }, (_, i) => ({ ...rows[0], plate: 'S-' + i, driver: 'd' + (i < n ? i + 1 : i), account: '012345678' + i })), Array.from({ length: 6 }, (_, i) => V('S-' + i, 'd' + i)), [], 'b').blocked;
check('แก้เลขในช่องผสาน = เตือนทุกคันที่เปลี่ยน', planBankImport([rows[0], { ...rows[0], plate: 'C-3', driver: 'นาย ค' }].map((r) => ({ ...r, account: '0555555555' })), vs, [A('A-1', '0123456789', 'นาย ก'), A('C-3', '0123456789', 'นาย ค')], 'b').rows.map((r) => r.warnings.length), [1, 1]);
check('แถวเลื่อนเกินครึ่งไฟล์ = บล็อก · ไม่กี่คัน = ผ่าน', [!!shift(5), !!shift(3)], [true, false]);
check('นำเข้าไฟล์เดิมซ้ำ (ชื่อไม่ตรง) = same', planBankImport([rows[4]], vs, [A('F-6', '0123456784', 'นาย ฉ', 'ชื่อในใบปะ: คนละคน')], 'b').rows[0].status, 'same');
check('ไม่มีชื่อในไฟล์ (ยังไม่เคยผูก) = เตือนแต่บันทึกได้', planBankImport([{ ...rows[0], driver: '' }], vs, [], 'b').rows[0].status, 'new');
check('เซลล์ตัวเลข 10 หลัก เตือน', plan.rows[5].warnings.length, 1);
check('ไฟล์สาขาอื่น บล็อกทั้งไฟล์', !!planBankImport(rows, [vs[0]], [], 'b').blocked, true);
check('ไม่บล็อกไฟล์ปกติ', plan.blocked, '');
const dup = planBankImport([{ ...rows[0] }, { ...rows[0], account: '0123456788' }], vs, [], 'b');
check('ทะเบียนเดียวสองเลข บล็อกทั้งคู่', dup.rows.map((r) => r.status), ['error', 'error']);
check('รถสาขาอื่นไม่ถูกจับ', planBankImport([rows[0]], [{ ...vs[0], branchId: 'x' }], [], 'b').rows[0].status, 'error');
check('คนขับเปลี่ยน + ไฟล์ชื่อคนเก่า = ยังกันไว้', planBankImport([rows[0]], [V('A-1', 'คนใหม่')], [A('A-1', '0123456789', 'นาย ก')], 'b').rows[0].status, 'error');

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
check('ใบปะ รวมยอดทั้งคัน เป็นข้อความ', pay.find((r) => r[0] === 'A-1'), ['A-1', 'นาย ก', '012-3-45678-9', 1485, 1485, 'หมายเหตุ1']);
check('ใบปะ ยอดรวม', pay.filter((r) => /^รวม/.test(String(r[0]))).map((r) => [r[0], r[3]]), [['รวมยอดโอน', 1485], ['รวมที่ยังไม่โอน', 297], ['รวมทั้งหมด (เท่ากับรับสุทธิในชีตสรุปรวม)', 1782]]);
const round = parseBankSheets(Buffer.from(await (blob as any).arrayBuffer()));
// หลายคันบัญชีเดียวกัน -> เรียงติดกัน + ผสานช่องเลขบัญชี + นำเข้ากลับได้ครบทุกคัน
await exportPerVehicleReport({ id: 'c', name: 'x' } as any, 'ทดสอบ', [T('J1', 'A-1', 1000, 100, 'เมือง', 'ตาก'), T('J3', 'B-2', 300, 30, 'เมือง', 'ตาก'), T('J4', 'C-3', 200, 30, 'เมือง', 'ตาก')], [], [], [...vx, V('C-3', 'นาย ค')], undefined, [], {
  bankAccounts: [A('A-1', '0123456789', 'นาย ก'), A('B-2', '0222222222', 'นาย ข', 'หมายเหตุ2'), A('C-3', '0123456789', 'นาย ค', 'หมายเหตุ3')], branchId: 'b',
});
const wb3 = new ExcelJS.Workbook(); await wb3.xlsx.load(await (blob as any).arrayBuffer()); const p3 = wb3.getWorksheet('ใบปะ')!;
const rowsOf = (ws: any) => { const o: any[] = []; ws.eachRow((r: any) => o.push((r.values as any[]).slice(1))); return o; };
check('บัญชีเดียวกันเรียงติดกัน', rowsOf(p3).filter((r) => /^[A-C]-/.test(r[0])).map((r) => r[0]), ['A-1', 'C-3', 'B-2']);
const rowNo = (pl: string) => { let n = 0; p3.eachRow((r: any, i: number) => { if (r.getCell(1).value === pl) n = i; }); return n; };
check('ผสานช่องเลขบัญชีเฉพาะ A-1..C-3', (p3.model as any).merges.filter((m: string) => /^C/.test(m)), [`C${rowNo('A-1')}:C${rowNo('C-3')}`]);
check('รวมเข้าบัญชีนี้ = ยอดรวมของกลุ่ม ผสานช่องเดียว', [p3.getCell(rowNo('A-1'), 5).value, (p3.model as any).merges.filter((m: string) => /^E/.test(m))], [1188, [`E${rowNo('A-1')}:E${rowNo('C-3')}`]]);
check('บัญชีเดี่ยว ยอดโอน = รับสุทธิ', p3.getCell(rowNo('B-2'), 5).value, 297);
check('หัวตาราง', (p3.getRow(rowNo('A-1') - 1).values as any[]).slice(1), ['ทะเบียน', 'คนขับ', 'เลขบัญชี', 'รับสุทธิ', 'ยอดโอนเข้าบัญชี', 'หมายเหตุ']);
const tot = rowsOf(p3).find((r) => r[0] === 'รวมทุกคัน'); check('รวมยอดโอน = ผลรวมคอลัมน์ยอดโอน', [tot[3], tot[4]], [1485, 1485]);
check('นำเข้ากลับได้ทุกคัน + หมายเหตุ', parseBankSheets(Buffer.from(await (blob as any).arrayBuffer())).map((r) => [r.plate, r.account, r.note]), [['A-1', '012-3-45678-9', ''], ['C-3', '012-3-45678-9', 'หมายเหตุ3'], ['B-2', '022-2-22222-2', 'หมายเหตุ2']]);
check('ไฟล์ที่ export นำเข้ากลับได้', round.map((r) => [r.plate, r.account, r.note]), [['A-1', '012-3-45678-9', 'หมายเหตุ1']]);
await exportPerVehicleReport({ id: 'c', name: 'x' } as any, 'ทดสอบ', [T('J1', 'A-1', 1000, 100, 'แม่สอด', 'ตาก')], [], [], vx, undefined, []);
const out2 = new ExcelJS.Workbook(); await out2.xlsx.load(await (blob as any).arrayBuffer());
check('ไม่ตั้งค่า = ไม่มีใบปะ/ไม่แยกตาราง', [!!out2.getWorksheet('ใบปะ'), out2.getWorksheet('สรุปรวม')!.rowCount], [false, 6]);

// --- r44: ไฟล์รวมหลายสาขา (โครงตามไฟล์ทีม "สรุปยอดจ่าย - รถร่วมภาคเหนือ" 25 พ.ย.69 · เลขสมมติ)
const mwb = XLSX.utils.book_new();
const own = XLSX.utils.aoa_to_sheet([
  ['ใบปะ — สรุปยอดโอนเข้าบัญชีคนขับ สาขาบี'], ['รอบ ต.ค. 69'], [],
  ['ทะเบียน', 'คนขับ', 'เลขบัญชี', 'รับสุทธิ'],
  ['A-1', 'ก', '0123456789', 10], ['C-3', 'นาย ค', '0222222222', 20, 1234.56], ['D-4', 'นาย ง', '0222222222', 30, 1234.56], ['F-6', 'นาย ฉ', '0333333333', 5, '08.10.2569 โทรแล้ว'],
  ['รวมทุกคัน', '', '', 65],
]);
XLSX.utils.book_append_sheet(mwb, own, 'สาขาบี');
XLSX.utils.book_append_sheet(mwb, XLSX.utils.aoa_to_sheet([[], [], [], ['ทะเบียน', 'คนขับ', 'เลขบัญชี', 'รับสุทธิ'], ...['X-1', 'X-2', 'X-3', 'X-4', 'X-5'].map((p, i) => [p, 'คนอื่น', '044444444' + i, 1])]), 'สาขาอื่น');
const outside = XLSX.utils.aoa_to_sheet([
  ['รายละเอียดกระจายสินค้า รอบวันที่ 16-30 ต.ค.'], ['ลำดับ', 'ทะเบียนรถ', 'เจ้าของรถ', 'จำนวนชิ้น', 'ค่าบรรทุก'], ['1', 'G-7', 'นาย ช', 10, 100], [],
  ['สรุปค่าบรรทุกรถนอก'], ['กำหนดจ่าย 20 พ.ย.'],
  ['ลำดับ', 'ทะเบียนรถ', ' ชื่อ - นามสกุล', '', 'ธนาคาร / เลขบัญชี / สาขา', '', '', '', 'สุทธิจ่าย'],
  ['1', 'G-7', 'นาย ช', 'นาย ช', 'กสิกรไทย  055-5-55555-5', '', '', '', 99],
  ['', '', '', '', '', '', '', '', 99],
]);
outside['!merges'] = [{ s: { r: 6, c: 2 }, e: { r: 6, c: 3 } }, { s: { r: 7, c: 4 }, e: { r: 7, c: 7 } }];
XLSX.utils.book_append_sheet(mwb, outside, 'รถนอก');
const mrows = parseBankSheets(XLSX.write(mwb, { type: 'buffer', bookType: 'xlsx' }));
check('คอลัมน์ธนาคาร: ชื่อธนาคาร+เลข 2 กลุ่ม = ไม่เดา (ตัวตรวจปฏิเสธ)', (() => { const w = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(w, XLSX.utils.aoa_to_sheet([['ทะเบียน', 'ธนาคาร / เลขบัญชี / สาขา'], ['G-7', 'กสิกรไทย 0555555555 สาขา 0123']]), 's'); return checkAccountNo(parseBankSheets(XLSX.write(w, { type: 'buffer', bookType: 'xlsx' }))[0].account).ok; })(), false);
check('คอลัมน์เลขบัญชีปกติ ยังไม่ตัดชื่อธนาคาร (เข้มเหมือนเดิม)', (() => { const w = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(w, XLSX.utils.aoa_to_sheet([['ทะเบียน', 'เลขบัญชี'], ['G-7', 'กสิกร 0555555555']]), 's'); return checkAccountNo(parseBankSheets(XLSX.write(w, { type: 'buffer', bookType: 'xlsx' }))[0].account).ok; })(), false);
check('ชีตรถนอก หัว "ธนาคาร / เลขบัญชี / สาขา" อ่านได้', mrows.filter((r) => r.sheet === 'รถนอก').map((r) => [r.plate, r.driver, checkAccountNo(r.account).ok]), [['G-7', 'นาย ช', true]]);
check('ยอดรวมกลุ่ม (ตัวเลข) ไม่กลายเป็นหมายเหตุ · ข้อความยังเก็บ', mrows.filter((r) => r.sheet === 'สาขาบี').map((r) => r.note), ['', '', '', '08.10.2569 โทรแล้ว']);
const mvs = [V('A-1', 'นาย ก'), V('C-3', 'นาย ค'), V('D-4', 'นาย ง'), V('F-6', 'นาย ฉ'), V('G-7', 'นาย ช')];
const mp = planBankImport(mrows, mvs, [], 'b');
check('ไฟล์รวมหลายสาขา = ข้ามชีตสาขาอื่น ไม่บล็อก', [mp.blocked, (mp as any).skippedSheets, mp.rows.map((r) => r.plate + ':' + r.status)], ['', ['สาขาอื่น'], ['A-1:new', 'C-3:new', 'D-4:new', 'F-6:new', 'G-7:new']]);
check('ไฟล์รวม ไม่มีชีตของสาขานี้ = บล็อก', /ไม่มีชีตของสาขานี้/.test(planBankImport(mrows, mvs, [], 'zz').blocked), true);
check('ไฟล์ชีตเดียวของสาขาอื่น ยังบล็อกเหมือนเดิม', /ไม่อยู่ในสาขานี้/.test(planBankImport(mrows.filter((r) => r.sheet === 'สาขาอื่น'), mvs, [], 'b').blocked), true);
check('คำนำหน้า นาย/นางสาว/น.ส. ไม่นับว่าชื่อไม่ตรง', ['ก', 'นางสาว ก', 'น.ส.ก', 'นาง ก'].map((d) => planBankImport([{ ...mrows[0], driver: d }], mvs, [], 'b').rows[0].warnings.length), [0, 0, 0, 0]);
check('ชื่อต่างจริงยังเตือน', planBankImport([{ ...mrows[0], driver: 'นาย ข' }], mvs, [], 'b').rows[0].warnings.length, 1);
check('ใบปะ: ด่านคนขับเปลี่ยนยังเข้ม (ชื่อในระบบต่างกัน = กันไว้)', buildPayslip([S('A-1', 'ก', 10)], [A('A-1', '0123456789', 'นาย ก')], 'b').hold.length, 1);
check('ชื่อขึ้นต้นคล้ายคำนำหน้า ไม่ถูกตัด (Mrinal / Missy)', ['Mrinal', 'Missy'].map((d) => planBankImport([{ ...mrows[0], driver: d }], [V('A-1', 'inal')], [], 'b').rows[0].warnings.length), [1, 1]);
check('Mr. ตามด้วยจุด/ช่องว่าง ตัดได้', planBankImport([{ ...mrows[0], driver: 'Mr. Somchai' }], [V('A-1', 'Somchai')], [], 'b').rows[0].warnings.length, 0);
const B1 = (cellText: string, head = 'ธนาคาร / เลขบัญชี / สาขา') => { const w = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(w, XLSX.utils.aoa_to_sheet([['ทะเบียน', head], ['G-7', cellText]]), 's'); return parseBankSheets(XLSX.write(w, { type: 'buffer', bookType: 'xlsx' }))[0]; };
check('เลข 9 หลัก + รหัสสาขาเว้นวรรค ห้ามต่อกันเป็น 10 หลัก', checkAccountNo(B1('กสิกร 123-4-56789 1').account).ok, false);
check('ขีดยาว – ในคอลัมน์ธนาคาร อ่านได้', checkAccountNo(B1('กสิกรไทย 055–5–55555–5').account).ok, true);
check('หัว "หมายเหตุ ตรวจเลขบัญชี" ไม่ถูกจับเป็นคอลัมน์เลขบัญชี', (() => { const w = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(w, XLSX.utils.aoa_to_sheet([['ทะเบียน', 'หมายเหตุ ตรวจเลขบัญชี', 'เลขบัญชี'], ['G-7', 'โทรแล้ว', '0555555555']]), 's'); const r = parseBankSheets(XLSX.write(w, { type: 'buffer', bookType: 'xlsx' }))[0]; return [r.account, r.note]; })(), ['0555555555', 'โทรแล้ว']);
check('คอลัมน์ "หมายเหตุ" จริง ตัวเลขล้วนยังเก็บ (เบอร์โทร)', (() => { const w = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(w, XLSX.utils.aoa_to_sheet([['ทะเบียน', 'เลขบัญชี', 'หมายเหตุ'], ['G-7', '0555555555', '0812345678']]), 's'); return parseBankSheets(XLSX.write(w, { type: 'buffer', bookType: 'xlsx' }))[0].note; })(), '0812345678');
const half = (local: number) => planBankImport([...Array.from({ length: local }, (_, i) => ({ ...mrows[0], sheet: 'k', plate: 'A-1', driver: 'ก' })), { ...mrows[0], sheet: 'k', plate: 'Z-1' }, { ...mrows[0], sheet: 'k', plate: 'Z-2' }, { ...mrows[0], sheet: 'k', plate: 'C-3', driver: 'นาย ค', account: '0222222222' }, { ...mrows[0], sheet: 'o', plate: 'X-9' }], mvs, [], 'b');
check('ชีตที่ทะเบียนสาขานี้ครึ่งพอดี (นับไม่ซ้ำ) ยังนำเข้า', half(3).skippedSheets, ['o']);
const tiny = planBankImport([{ ...mrows[0], sheet: 'เล็ก' }, ...['X-1', 'X-2', 'X-3', 'X-4'].map((p, i) => ({ ...mrows[0], sheet: 'ใหญ่', plate: p, account: '044444444' + i }))], mvs, [], 'b');
check('ไฟล์สาขาอื่นที่มีชีตเล็กตรงสาขานี้ = เตือนเลือกสาขาผิด', [tiny.skippedSheets, /ตรวจว่าเลือกสาขาถูก/.test(tiny.rows[0].warnings.join())], [['ใหญ่'], true]);
check('ไฟล์รวมปกติ (สาขานี้หลายคัน) ไม่เตือนเลือกสาขาผิด', mp.rows.some((r) => /เลือกสาขาถูก/.test(r.warnings.join())), false);

console.log(`\nผ่าน ${pass}/${pass + fail}`);
if (fail) process.exit(1);
