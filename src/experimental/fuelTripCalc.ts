// ============================================================================
// [ทดลอง] แกนคำนวณค่าจ้างรถร่วมจากราคาน้ำมัน (BAND mode) — แยกจากระบบค่าเที่ยวเดิม 100%
// สูตรจาก "มาตรฐานการจ่าย เบี้ยเลี้ยง รถร่วม.xlsx"
// หลักการ: ค่าทั้งหมดมาจาก policy (ห้าม hardcode ในฟังก์ชันคำนวณ) เพื่อทำ version ได้ภายหลัง
// ============================================================================

/** ปัดครึ่งขึ้น (ROUND_HALF_UP) n ตำแหน่ง — ตัด noise ของ floating point ก่อน */
export function roundHalfUp(value: number, decimals = 2): number {
  if (!Number.isFinite(value)) return 0;
  const f = Math.pow(10, decimals);
  const scaled = Number((value * f).toFixed(6));
  return (Math.sign(scaled) * Math.floor(Math.abs(scaled) + 0.5)) / f;
}

/** policy = ค่าตั้งต้นที่ทำ version ได้ (เก็บเป็น master ไม่ hardcode) */
export interface FuelTripPolicy {
  nearSpeedKmh: number;          // ความเร็วช่วงใกล้ (default 50) สำหรับ distanceMax <= speedThreshold
  farSpeedKmh: number;           // ความเร็วช่วงไกล (default 75)
  speedThresholdKm: number;      // จุดตัดความเร็ว (default 84 -> <=84 ใช้ near, >84 ใช้ far)
  fuelEfficiencyKmPerL: number;  // อัตราสิ้นเปลือง (default 10 km/L) — ต่อประเภทรถได้
  driverHourlyRate: number;      // ค่าแรงคนขับต่อ ชม. (default 1200/8 = 150)
  unloadingMinutesPerStore: number; // นาทีลงสินค้าต่อร้าน (default 30)
  baseBoxes: number;             // กล่องฐานสำหรับหาเรทต่อกล่อง (default 250)
  lumpSumBoxThreshold: number;   // เส้นแบ่งเหมา: actual <= นี้ = เหมาทั้งเที่ยว (default 260 — รอยืนยันนโยบาย)
  finalRoundingDecimals: number; // ปัดยอดสุดท้ายกี่ตำแหน่ง (default 2)
}

export const DEFAULT_FUEL_POLICY: FuelTripPolicy = {
  nearSpeedKmh: 50,
  farSpeedKmh: 75,
  speedThresholdKm: 84,
  fuelEfficiencyKmPerL: 10,
  driverHourlyRate: 1200 / 8, // = 150
  unloadingMinutesPerStore: 30,
  baseBoxes: 250,
  lumpSumBoxThreshold: 260, // ⚠️ รอ Admin ยืนยัน (เอกสารเดิม 250-260 ไม่ชัด)
  finalRoundingDecimals: 2,
};

export interface FuelTripInput {
  distanceMinKm: number;         // ต่ำสุดของช่วงระยะ (BAND)
  distanceMaxKm: number;         // สูงสุดของช่วงระยะ (BAND) — ใช้คิดเวลา + เลือกความเร็ว
  storeCount: number;            // จำนวนร้านที่ลงสินค้า
  appliedFuelPrice: number;      // ราคาน้ำมันที่ใช้จริง (บาท/ลิตร) — จาก OR หรือ manual
  actualBoxes: number;           // จำนวนกล่องจริงของเที่ยว
  mountainAllowance?: number;    // ค่าน้ำมันขึ้นเขา (จาก Route Surcharge Master)
  routeSurcharge?: number;       // ค่าเส้นทางพิเศษอื่น
  localTaxAdjustment?: number;   // ภาษีบำรุงท้องที่ (ถ้ามี) — แยกต่างหากจากราคา OR
}

export interface FuelTripBreakdown {
  // input echo
  distanceMinKm: number;
  distanceMaxKm: number;
  storeCount: number;
  appliedFuelPrice: number;
  actualBoxes: number;
  // ค่ากลาง (เก็บ full precision — ห้ามเอาค่าที่ปัดแล้วไปคูณต่อ)
  bandAverageOneWayKm: number;
  estimatedRoundTripKm: number;
  speedKmh: number;
  travelHours: number;
  unloadingHours: number;
  driverAllowance: number;
  fuelCost: number;
  extraCharges: number;          // mountain + route + localTax
  tripRateBeforeRounding: number;
  boxRateUnrounded: number;
  boxRateDisplay: number;        // แสดง 2 ตำแหน่ง (ห้ามเอาไปคูณ)
  // ผลลัพธ์
  paymentMode: 'LUMP_SUM' | 'PER_BOX';
  finalPayment: number;          // ปัดครั้งเดียวตรงนี้
}

/**
 * คำนวณค่าจ้างรถร่วม 1 เที่ยว (BAND mode)
 * ⚠️ เวลา (travelHours) ใช้ "ปลายบนของช่วง" (distanceMax) = นโยบายอนุรักษนิยมด้านเวลา
 *    ส่วนน้ำมันใช้ "ค่ากลางของช่วง" — ตรงตาม Excel เดิม
 */
export function computeFuelTrip(input: FuelTripInput, policy: FuelTripPolicy = DEFAULT_FUEL_POLICY): FuelTripBreakdown {
  const { distanceMinKm, distanceMaxKm, storeCount, appliedFuelPrice, actualBoxes } = input;
  const mountain = input.mountainAllowance ?? 0;
  const route = input.routeSurcharge ?? 0;
  const tax = input.localTaxAdjustment ?? 0;

  const bandAverageOneWayKm = (distanceMinKm + distanceMaxKm) / 2;
  const estimatedRoundTripKm = bandAverageOneWayKm * 2;
  const speedKmh = distanceMaxKm <= policy.speedThresholdKm ? policy.nearSpeedKmh : policy.farSpeedKmh;
  // เวลาเดินทางใช้ระยะปลายบนของช่วง (ไป-กลับ) หารความเร็ว
  const travelHours = (distanceMaxKm * 2) / speedKmh;
  const unloadingHours = (storeCount * policy.unloadingMinutesPerStore) / 60;
  const driverAllowance = (travelHours + unloadingHours) * policy.driverHourlyRate;
  const fuelCost = (estimatedRoundTripKm / policy.fuelEfficiencyKmPerL) * appliedFuelPrice;
  const extraCharges = mountain + route + tax;
  const tripRateBeforeRounding = driverAllowance + fuelCost + extraCharges;
  const boxRateUnrounded = tripRateBeforeRounding / policy.baseBoxes;

  // เส้นแบ่งเหมา: actual <= threshold = เหมาทั้งเที่ยว, เกิน = เรทต่อกล่อง × กล่องจริง
  const paymentMode: 'LUMP_SUM' | 'PER_BOX' = actualBoxes <= policy.lumpSumBoxThreshold ? 'LUMP_SUM' : 'PER_BOX';
  const rawFinal = paymentMode === 'LUMP_SUM' ? tripRateBeforeRounding : boxRateUnrounded * actualBoxes;
  const finalPayment = roundHalfUp(rawFinal, policy.finalRoundingDecimals); // ปัดครั้งเดียวที่นี่

  return {
    distanceMinKm, distanceMaxKm, storeCount, appliedFuelPrice, actualBoxes,
    bandAverageOneWayKm, estimatedRoundTripKm, speedKmh,
    travelHours, unloadingHours, driverAllowance, fuelCost, extraCharges,
    tripRateBeforeRounding, boxRateUnrounded,
    boxRateDisplay: roundHalfUp(boxRateUnrounded, 2),
    paymentMode, finalPayment,
  };
}
