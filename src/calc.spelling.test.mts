// ทดสอบชื่อพื้นที่สะกดด้วยตัวหน้าตาเหมือน (ฎ/ฏ) — เรียก computeTripDocument ตัวจริง
// รัน: npx tsx src/calc.spelling.test.mts   (exit 1 ถ้ามีเคสตก)
// ที่มา: 21 พ.ย.69 JB0226015659 สาย3 จัสทรานสะกด "เขาคิชฌกูฎ" ราคาในระบบ "เขาคิชฌกูฏ" -> ไม่เจอราคา
import { computeTripDocument, normPlate, textContains, placeContains } from './calc.js';

const R = (id: string, dist: string, type: 'flat' | 'piece', price: number) => ({
  id, branchId: 'b', provinceName: 'จันทบุรี', provinceShort: '', districtName: dist, destinationName: dist, price,
  priceType: type, productCategory: 'normal', effectiveFrom: '2020-01-01', status: 'active',
}) as any;
const rates = [
  R('k-f', 'เขาคิชฌกูฏ(จากเมือง)', 'flat', 3120), R('k-p', 'เขาคิชฌกูฏ(จากเมือง)', 'piece', 15),
  R('m-f', 'เมืองจันทบุรี', 'flat', 3020), R('m-p', 'เมืองจันทบุรี', 'piece', 15),
];
const ctx = {
  cycleId: 'c', cycle: { year: 2026, month: 11, half: 'first' }, vehicles: [], rates, rateOverrides: new Map(),
  groups: [], aliases: [], rules: [], manualBoxSenders: [], destOverrides: [], minBoxes: null, fileName: 't',
} as any;
const trip = (stops: [string, number][]) => ({
  documentNo: 'T', documentDate: '2026-11-03', plateNo: 'X', provinceRaw: 'จันทบุรี', districtRaw: stops[0][0],
  receipts: stops.map(([d, q], i) => ({
    receiptNo: 'R' + i, receiverName: 'r', senderName: 's', provinceRaw: 'จันทบุรี', districtRaw: d,
    items: [{ productName: 'สินค้า', quantity: q, unit: 'กล่อง' }],
  })),
}) as any;

let pass = 0, fail = 0;
const check = (name: string, got: any, want: any) => {
  const ok = got === want; ok ? pass++ : fail++;
  console.log(`${ok ? '✓' : '✗ FAIL'} ${name}: ได้ ${got} (คาด ${want})`);
};
// สะกด ฎ (แบบจัสทราน) ต้องเจอราคาเหมือนสะกด ฏ
for (const d of ['เขาคิชฌกูฎ', 'เขาคิชฌกูฏ', 'คิชฌกูฎ']) {
  const t = computeTripDocument(trip([[d, 10]]), ctx, () => 'x');
  check(`"${d}" เจอราคาเหมา`, t.receipts[0].flatPrice, 3120);
  check(`"${d}" เจอราคาชิ้น`, t.receipts[0].piecePrice, 15);
}
// เคสจริง: เขาคิชฌกูฎ 48+8 + เมือง 227 กล่อง -> ทุกจุดต้องมีราคา (ไม่มีจุด "ไม่เจอราคาชิ้น")
const real = computeTripDocument(trip([['เขาคิชฌกูฎ', 48], ['เขาคิชฌกูฎ', 8], ['เมืองจันทบุรี', 227]]), ctx, () => 'x');
check('เคสจริง ไม่มีคำเตือนไม่เจอราคา', real.warnings.filter((w: string) => /ไม่เจอราคา/.test(w)).length, 0);
check('เคสจริง ราคาชิ้นรวมทุกจุด 283 × 15', real.tripAmount, 4245);
// ทะเบียนรถห้ามรวม ฎ/ฏ (หมวดอักษรคนละหมวด เช่น 3ฒฎ-745 มีจริงในระบบ) — ไม่งั้นรถ 2 คันกลายเป็นคันเดียว
check('ทะเบียน ฎ กับ ฏ ต้องเป็นคนละคัน', normPlate('3ฒฎ-745') === normPlate('3ฒฏ-745'), false);
check('ทะเบียนเดิมยังเทียบได้ (ตัด "ตู้"/ช่องว่าง)', normPlate('3ฒฎ-745 ตู้') === normPlate('3ฒฎ-745'), true);
// ชื่อผู้รับ/สินค้า/คีย์เวิร์ดกฎ ต้องเทียบเข้มแบบเดิม — รวม ฎ/ฏ เฉพาะชื่อพื้นที่ (Codex P2)
check('textContains ไม่รวม ฎ/ฏ (ผู้รับ/สินค้า)', textContains('ร้านกฎหมาย', 'กฏหมาย'), false);
check('placeContains รวม ฎ/ฏ (ชื่อพื้นที่)', placeContains('เขาคิชฌกูฎ', 'คิชฌกูฏ'), true);
console.log(`\nผ่าน ${pass}/${pass + fail}`);
if (fail) process.exit(1);
