// ทดสอบแยกตารางสรุปรวม + ประสิทธิภาพ — เรียกโค้ดจริง  รัน: npx tsx src/summarySplit.test.mts
import { splitSummaries, tripZones, loadEfficiency } from './summarySplit.js';
import { summarizeByVehicle } from './calc.js';

let pass = 0, fail = 0;
const check = (name: string, got: any, want: any) => {
  const ok = JSON.stringify(got) === JSON.stringify(want); ok ? pass++ : fail++;
  console.log(`${ok ? '✓' : '✗ FAIL'} ${name}: ได้ ${JSON.stringify(got)}${ok ? '' : ` (คาด ${JSON.stringify(want)})`}`);
};
const T = (no: string, plate: string, amount: number, qty: number, stops: [string, string][]) => ({
  id: no, cycleId: 'c', branchId: 'b', documentNo: no, plateNo: plate, tripAmount: amount, billingQty: qty, totalQty: qty + 50,
  receipts: stops.map(([d, p]) => ({ districtRaw: d, provinceRaw: p })),
}) as any;
const MH: [string, string] = ['ขุนยวม', 'แม่ฮ่องสอน'], CM: [string, string] = ['สันทราย', 'เชียงใหม่'];
const veh = [{ plateNo: 'A-1', driverName: 'ก', branchId: 'b' }, { plateNo: 'B-2', driverName: 'ข', branchId: 'b' }, { plateNo: 'C-3', driverName: 'ค', branchId: 'b' }, { plateNo: 'D-4', driverName: 'ง', branchId: 'b' }] as any[];
const D = (plate: string, kind: string, amount: number, docNo = '', label = 'x') => ({ id: plate + amount, cycleId: 'c', branchId: 'b', plateNo: plate, kind, amount, docNo, label, categoryId: 'cat-' + label }) as any;

// --- การจับพื้นที่
check('ไม่ตั้งค่า = ไม่แยก', tripZones(T('1', 'A-1', 1, 1, [MH]), ''), []);
check('undefined = ไม่แยก', tripZones(T('1', 'A-1', 1, 1, [MH]), undefined), []);
check('ทั้งจังหวัด', tripZones(T('1', 'A-1', 1, 1, [CM, MH]), 'แม่ฮ่องสอน'), ['แม่ฮ่องสอน']);
check('จังหวัด: อำเภอ ใช้ชื่ออำเภอเป็นหมายเหตุ', tripZones(T('1', 'A-1', 1, 1, [['พบพระ', 'ตาก'], ['แม่สอด', 'ตาก']]), 'ตาก: แม่สอด'), ['แม่สอด']);
check('จังหวัดตรงแต่อำเภอไม่ตรง', tripZones(T('1', 'A-1', 1, 1, [['เมืองตาก', 'ตาก']]), 'ตาก: แม่สอด'), []);
check('อำเภอว่างไม่ถือว่าตรง', tripZones(T('1', 'A-1', 1, 1, [['', 'ตาก']]), 'ตาก: แม่สอด'), []);
check('อำเภอชื่อสั้นกว่าค่าที่ตั้ง ไม่ถือว่าตรง', tripZones(T('1', 'A-1', 1, 1, [['แม่', 'ตาก']]), 'ตาก: แม่สอด'), []);
check('มีคำนำหน้า อ./จ.', tripZones(T('1', 'A-1', 1, 1, [['อ.แม่สอด', 'จ.ตาก']]), 'ตาก: แม่สอด'), ['แม่สอด']);

check('ตั้งค่าใส่คำนำหน้า จ./อ.', [tripZones(T('1', 'A-1', 1, 1, [MH]), 'จ.แม่ฮ่องสอน'), tripZones(T('1', 'A-1', 1, 1, [['แม่สอด', 'ตาก']]), 'จังหวัดตาก: อำเภอแม่สอด')], [['แม่ฮ่องสอน'], ['แม่สอด']]);
check('ทะเบียนซ้ำใช้คนขับคันที่ใช้งาน', summarizeByVehicle('c', [T('1', 'A-1', 1, 1, [CM])], [], [], [{ plateNo: 'A-1', driverName: 'เก่า', status: 'inactive' }, { plateNo: 'A-1', driverName: 'ใหม่', status: 'active' }] as any)[0].driverName, 'ใหม่');

// --- ประสิทธิภาพ = รายได้ ÷ ชิ้น (ตัวเลขจากภาพทีม)
check('บฉ-8198', loadEfficiency(10411, 1148.78), 9.06);
check('บร-4584 แม่ฮ่องสอน', loadEfficiency(11439, 570), 20.07);
check('ชิ้น 0 ไม่เป็น Infinity', loadEfficiency(500, 0), 0);

// --- แยกยอด
const trips = [
  T('JB1', 'A-1', 24448, 3123.83, [CM]), T('JB2', 'A-1', 7626, 380, [MH]), T('JB3', 'A-1', 3813, 190, [CM, MH]), // บ้าน = หลัก
  T('JB4', 'B-2', 1409, 206, [CM]), T('JB5', 'B-2', 37291.06, 3137, [MH]),   // บ้าน = พื้นที่
  T('JB6', 'C-3', 10411, 1148.78, [CM]),                                      // หลักล้วน
  T('JB7', 'D-4', 39494.54, 3743, [MH]),                                      // พื้นที่ล้วน
];
const fuel = [{ cycleId: 'c', plateNo: 'A-1', amount: 6000 }, { cycleId: 'c', plateNo: 'B-2', amount: 16500 }, { cycleId: 'c', plateNo: 'D-4', amount: 15200 }, { cycleId: 'c', plateNo: 'E-5', amount: 100 }] as any[];
const ded = [D('D-4', 'deduction', 4400), D('A-1', 'deduction', 333.33, '', 'v'), D('A-1', 'income', 777.77, 'JB2'), D('A-1', 'income', 55.55, 'JB1', 'y'), D('B-2', 'income', 100.1, '', 'z'), D('B-2', 'deduction', 20, 'JB4', 'w')];
const r = splitSummaries('c', trips, fuel, ded, veh, 'แม่ฮ่องสอน');
const m = (p: string) => r.main.find((x) => x.plateNo === p), z = (p: string) => r.zone.find((x) => x.plateNo === p);
check('แถวหลัก', r.main.map((x) => x.plateNo), ['A-1', 'B-2', 'C-3', 'E-5']);
check('แถวพื้นที่', r.zone.map((x) => x.plateNo), ['A-1', 'B-2', 'D-4']);
check('ใบผสมไปตารางพื้นที่ทั้งใบ', [z('A-1')!.trip, z('A-1')!.qty], [11439, 570]);
check('หมายเหตุ', z('A-1')!.note, 'แม่ฮ่องสอน');
check('น้ำมันอยู่ฝั่งรายได้มากกว่า (หลัก)', [m('A-1')!.fuel, z('A-1')!.fuel], [6000, 0]);
check('น้ำมันอยู่ฝั่งรายได้มากกว่า (พื้นที่)', [m('B-2')!.fuel, z('B-2')!.fuel], [0, 16500]);
check('รายการผูกใบตามใบไป', [m('A-1')!.inc, z('A-1')!.inc, m('B-2')!.ded, z('B-2')!.ded], [55.55, 777.77, 20, 0]);
check('รายการไม่ผูกใบอยู่ฝั่งบ้าน', [m('A-1')!.ded, z('A-1')!.ded, z('B-2')!.inc], [333.33, 0, 100.1]);
check('รถที่มีแต่น้ำมัน อยู่ตารางหลัก', m('E-5')!.fuel, 100);
check('จำนวนชิ้นใช้ยอดหลังหาร', m('C-3')!.qty, 1148.78);
check('พื้นที่ล้วน = ยอดเดิมทั้งก้อน', [z('D-4')!.net, z('D-4')!.ded], [19499.59, 4400]);
const all = summarizeByVehicle('c', trips, fuel, ded, veh);
for (const s of all) {
  const parts = [m(s.plateNo), z(s.plateNo)].filter(Boolean) as any[];
  const sum = (k: string) => Math.round(parts.reduce((a, x) => a + x[k], 0) * 100) / 100;
  check(`${s.plateNo} หลัก+พื้นที่ = ยอดเดิม`, [sum('trip'), sum('d1'), sum('fuel'), sum('inc'), sum('ded'), sum('net')],
    [s.totalTripAmount, s.deduction1Percent, s.fuelTotal, s.incomeAdd, s.deductionTotal, s.netReceive]);
  for (const x of parts) check(`${s.plateNo} แถวบวกลบลงตัว`, Math.round((x.trip - x.d1 - x.fuel + x.inc - x.ded) * 100) / 100, x.net);
}
const none = splitSummaries('c', trips, fuel, ded, veh, '');
check('ไม่ตั้งพื้นที่ = ตารางเดียวเท่าเดิม', [none.zone.length, none.main.map((x) => x.net)], [0, all.map((s) => s.netReceive)]);

console.log(`\nผ่าน ${pass}/${pass + fail}`);
if (fail) process.exit(1);
