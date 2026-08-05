// [ทดลอง] test เทียบ Excel — รัน: npx tsx src/experimental/fuelTripCalc.test.mts
import { computeFuelTrip, DEFAULT_FUEL_POLICY } from './fuelTripCalc.js';

let pass = 0, fail = 0;
const near = (a: number, b: number, eps = 0.01) => Math.abs(a - b) <= eps;
function check(name: string, got: number, want: number, eps = 0.01) {
  const ok = near(got, want, eps);
  console.log(`${ok ? '✓' : '✗ FAIL'} ${name}: ได้ ${got}  (คาด ${want})`);
  ok ? pass++ : fail++;
}

// helper: คิด 1 เที่ยว
const calc = (min: number, max: number, price: number, stores = 3, boxes = 250) =>
  computeFuelTrip({ distanceMinKm: min, distanceMaxKm: max, storeCount: stores, appliedFuelPrice: price, actualBoxes: boxes });

console.log('=== ชุดทดสอบบังคับ (เทียบ Excel) ===\n');

// 0-55 กม. ราคา 38 -> เวลา 2.2, เบี้ยขับ 555, น้ำมัน 209, ค่าเที่ยว 764, เรท 3.056
let r = calc(0, 55, 38);
check('0-55 travelHours', r.travelHours, 2.2);
check('0-55 driverAllowance', r.driverAllowance, 555);
check('0-55 fuelCost', r.fuelCost, 209);
check('0-55 tripRate', r.tripRateBeforeRounding, 764);
check('0-55 boxRate', r.boxRateUnrounded, 3.056, 0.001);

// 56-84 ราคา 38 -> เบี้ยขับ 729, น้ำมัน 532, ค่าเที่ยว 1261
r = calc(56, 84, 38);
check('56-84 driverAllowance', r.driverAllowance, 729);
check('56-84 fuelCost', r.fuelCost, 532);
check('56-84 tripRate', r.tripRateBeforeRounding, 1261);

// ⭐ REGRESSION 85-113 ราคา 38 -> เบี้ยขับต้อง 677 (ไม่ใช่ 729!), น้ำมัน 752.40, ค่าเที่ยว 1429.40, เรท 5.7176
r = calc(85, 113, 38);
check('★85-113 driverAllowance = 677 ไม่ใช่ 729', r.driverAllowance, 677);
check('85-113 fuelCost', r.fuelCost, 752.4);
check('85-113 tripRate', r.tripRateBeforeRounding, 1429.4);
check('85-113 boxRate', r.boxRateUnrounded, 5.7176, 0.0001);

// 114-142 ราคา 38 -> เบี้ยขับ 793, น้ำมัน 972.80, ค่าเที่ยว 1765.80, เรท 7.0632
r = calc(114, 142, 38);
check('114-142 driverAllowance', r.driverAllowance, 793);
check('114-142 fuelCost', r.fuelCost, 972.8);
check('114-142 tripRate', r.tripRateBeforeRounding, 1765.8);
check('114-142 boxRate', r.boxRateUnrounded, 7.0632, 0.0001);

// ราคาน้ำมันเปลี่ยน: 114-142 ราคา 40 -> น้ำมัน 1024, ค่าเที่ยว 1817
r = calc(114, 142, 40);
check('114-142@40 fuelCost', r.fuelCost, 1024);
check('114-142@40 tripRate', r.tripRateBeforeRounding, 1817);

// ขอบช่วง: 84 ใช้ near(50), 85 ใช้ far(75)
check('ขอบ 84 -> speed 50', calc(56, 84, 38).speedKmh, 50);
check('ขอบ 85 -> speed 75', calc(85, 113, 38).speedKmh, 75);

// จำนวนร้าน +1 -> unloading +0.5 ชม, เบี้ยขับ +75
const base = calc(0, 55, 38, 3);
const plus1 = calc(0, 55, 38, 4);
check('เพิ่ม 1 ร้าน -> unloadingHours +0.5', plus1.unloadingHours - base.unloadingHours, 0.5);
check('เพิ่ม 1 ร้าน -> driverAllowance +75', plus1.driverAllowance - base.driverAllowance, 75);

// payment mode: 250 กล่อง (<=260) = เหมาทั้งเที่ยว
check('250 กล่อง -> LUMP_SUM finalPayment = tripRate', calc(0, 55, 38, 3, 250).finalPayment, roundTest(764));
// 300 กล่อง (>260) = เรทต่อกล่อง × 300
const perBox = calc(0, 55, 38, 3, 300);
check('300 กล่อง -> PER_BOX', perBox.finalPayment, roundTest(perBox.boxRateUnrounded * 300));

function roundTest(v: number) { const s = Number((v * 100).toFixed(6)); return Math.floor(s + 0.5) / 100; }

// ===== computeLoopTripCost (ระยะลูปจริง) =====
import { computeLoopTripCost } from './fuelTripCalc.js';
console.log('\n=== computeLoopTripCost (ลูป 177 กม, 4 จุด, ดีเซล 36.69) ===');
const loop = computeLoopTripCost(177, 4, 36.69);
// oneWay 88.5 > 84 -> speed 75 | travel 177/75=2.36 | unload 4*30/60=2 | เบี้ยขับ (2.36+2)*150=654
// น้ำมัน 177/10*36.69=649.413 | รวม 1303.413 -> 1303.41
check('loop speed (oneWay 88.5>84 -> 75)', loop.speedKmh, 75);
check('loop travelHours (177/75)', loop.travelHours, 2.36, 0.001);
check('loop unloadingHours (4 จุด)', loop.unloadingHours, 2);
check('loop driverAllowance', loop.driverAllowance, 654);
check('loop fuelCost (177/10*36.69)', loop.fuelCost, 649.413, 0.001);
check('loop totalCost (ปัด)', loop.totalCost, 1303.41, 0.001);
// ลูปสั้น: 60 กม (oneWay 30<=84 -> speed 50)
check('loop สั้น speed (oneWay 30<=84 -> 50)', computeLoopTripCost(60, 2, 38).speedKmh, 50);
// loopKm=0 -> ไม่พัง
check('loopKm 0 -> totalCost 0', computeLoopTripCost(0, 3, 38).totalCost, 0);
// น้ำมันขึ้นเขา: +5 ลิตร × 36.69 = 183.45 บวกเข้า fuelCost (177 กม เดิม)
import { DEFAULT_FUEL_POLICY as POL } from './fuelTripCalc.js';
const mtn = computeLoopTripCost(177, 4, 36.69, POL, 5);
check('mountainLiters echo', mtn.mountainLiters, 5);
check('mountainFuelCost (5*36.69)', mtn.mountainFuelCost, 183.45, 0.001);
check('fuelCost รวมขึ้นเขา (649.413+183.45)', mtn.fuelCost, 832.863, 0.001);
check('totalCost รวมขึ้นเขา (654+832.863)', mtn.totalCost, 1486.86, 0.01);
// ลิตรขึ้นเขาติดลบ/NaN -> ปัดเป็น 0 (กัน master กรอกพลาด)
check('mountainLiters ลบ -> 0', computeLoopTripCost(177, 4, 36.69, POL, -3).mountainFuelCost, 0);
// km<=0 แต่มีลิตรขึ้นเขา -> ยัง 0 (ไม่มีระยะ = ไม่มีข้อมูล ไม่คิดอะไรเลย)
check('km 0 + ขึ้นเขา -> 0', computeLoopTripCost(0, 4, 36.69, POL, 5).totalCost, 0);

console.log(`\n=== สรุป: ผ่าน ${pass} | ตก ${fail} ===`);
console.log('policy ที่ใช้:', JSON.stringify(DEFAULT_FUEL_POLICY));
if (fail > 0) process.exit(1);
