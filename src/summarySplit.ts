// ============================================================================
// แยกตาราง "สรุปรวมต่อทะเบียน" ตามพื้นที่พิเศษ (เช่น แม่ฮ่องสอน / แม่สอด) — pure function
// เจ้าของสั่ง 25 พ.ย.69 (ต้นแบบ: ไฟล์ทีม เชียงใหม่ ต.ค.16-31 / กำแพงเพชร ต.ค.1-15)
//
// กติกา (ถอดจากไฟล์ทีม ทานกับข้อมูลจริงตรง 8/8 แถวที่ทีมแยก):
//   · ใบที่มีอย่างน้อย 1 จุดส่งในพื้นที่ = ทั้งใบไปตารางพื้นที่ (รวมใบผสมหลายจังหวัด)
//   · ค่าน้ำมัน/รายการที่ไม่ผูกเลขใบ ไม่แยก -> อยู่ฝั่งที่ "รายได้มากกว่า" (เท่ากัน = ตารางหลัก)
//   · รายการ +/- ที่ผูกเลขใบกระจาย -> ตามใบนั้นไป
// กฎเหล็ก: แถวหลัก + แถวพื้นที่ = ยอดเดิมของ summarizeByVehicle ทุกช่อง (ตารางหลักรับเศษปัด)
// ============================================================================
import type { TripDocument, FuelEntry, DeductionEntry, Vehicle, VehicleSummary } from './types';
import { summarizeByVehicle, normPlate, normDoc, round2, placeContains } from './calc';
import { parseServiceAreas } from './serviceArea';

export interface SplitRow {
  plateNo: string;
  driverName: string;
  qty: number;          // จำนวนชิ้น "หลังหาร" (billingQty) — เจ้าของเคาะ 25 พ.ย.69
  trip: number;
  d1: number;
  fuel: number;
  inc: number;
  ded: number;
  net: number;
  note: string;         // ชื่อพื้นที่ (เฉพาะตารางพื้นที่)
}

// พื้นที่ที่ใบนี้วิ่งไป (ว่าง = ไม่มี) — ตั้งค่าว่าง = ไม่แยกตาราง
//   ไม่ใช้ inServiceArea เพราะตัวนั้น "ไม่ตั้งค่า = ผ่านหมด" และเทียบกลับด้าน (คำสั้นชนคำยาว)
export function tripZones(t: TripDocument, zoneText?: string): string[] {
  // ตัดคำนำหน้าที่คนพิมพ์มา ("จ.แม่ฮ่องสอน" / "อ.แม่สอด") ไม่งั้นเทียบไม่ติดแล้วไม่แยกตารางแบบเงียบ ๆ
  const areas = parseServiceAreas(zoneText)
    .map((a) => ({ prov: a.prov.replace(/^(จังหวัด|จ\.)\s*/, ''), dists: a.dists && a.dists.map((d) => d.replace(/^(อำเภอ|อ\.)\s*/, '')) }))
    .filter((a) => a.prov);
  if (!areas.length) return [];
  const out: string[] = [];
  const add = (s: string) => { if (!out.includes(s)) out.push(s); };
  for (const r of t.receipts || []) {
    for (const a of areas) {
      if (!placeContains(r.provinceRaw || '', a.prov)) continue;
      if (!a.dists) { add(a.prov); continue; }
      for (const d of a.dists) if (placeContains(r.districtRaw || '', d)) add(d);
    }
  }
  return out;
}

// ประสิทธิภาพในการบรรทุก = รายได้ ÷ จำนวนชิ้น (บาทต่อชิ้น) — เจ้าของเคาะ 25 พ.ย.69
//   (เดิมคิดกลับด้าน ชิ้น÷รายได้) · ชิ้น 0 -> 0 กัน Infinity/NaN
export const loadEfficiency = (trip: number, qty: number): number => (qty > 0 ? round2(trip / qty) : 0);

export function splitSummaries(
  cycleId: string,
  trips: TripDocument[],
  fuel: FuelEntry[],
  deductions: DeductionEntry[],
  vehicles: Vehicle[],
  zoneText?: string,
): { main: SplitRow[]; zone: SplitRow[] } {
  const all = summarizeByVehicle(cycleId, trips, fuel, deductions, vehicles);
  const cycleTrips = trips.filter((t) => t.cycleId === cycleId);
  const main: SplitRow[] = [];
  const zone: SplitRow[] = [];
  const qtyOf = (ts: TripDocument[]) => round2(ts.reduce((a, t) => a + Number(t.billingQty || 0), 0));
  const row = (s: VehicleSummary, qty: number, note = ''): SplitRow => ({
    plateNo: s.plateNo, driverName: s.driverName, qty, trip: s.totalTripAmount, d1: s.deduction1Percent,
    fuel: s.fuelTotal, inc: s.incomeAdd, ded: s.deductionTotal, net: s.netReceive, note,
  });

  for (const s of all) {
    const key = normPlate(s.plateNo);
    const pt = cycleTrips.filter((t) => normPlate(t.plateNo) === key);
    const labels: string[] = [];
    const zt: TripDocument[] = [];
    const mt: TripDocument[] = [];
    for (const t of pt) {
      const z = tripZones(t, zoneText);
      if (z.length) { zt.push(t); z.forEach((x) => { if (!labels.includes(x)) labels.push(x); }); }
      else mt.push(t);
    }
    if (!zt.length) { main.push(row(s, qtyOf(pt))); continue; }
    if (!mt.length) { zone.push(row(s, qtyOf(pt), labels.join(', '))); continue; }

    // มีทั้งสองฝั่ง -> คิดฝั่งพื้นที่ด้วยสูตรเดียวกับยอดรวม แล้วให้ตารางหลัก = ยอดรวม − ฝั่งพื้นที่
    const amt = (ts: TripDocument[]) => ts.reduce((a, t) => a + t.tripAmount, 0);
    const zoneIsHome = amt(zt) > amt(mt);
    const zDocs = new Set(zt.map((t) => normDoc(t.documentNo || '')).filter(Boolean));
    const mDocs = new Set(mt.map((t) => normDoc(t.documentNo || '')).filter(Boolean));
    const zDed = deductions.filter((d) => {
      if (d.cycleId !== cycleId || normPlate(d.plateNo) !== key) return false;
      const doc = normDoc(d.docNo || '');
      if (doc && zDocs.has(doc)) return true;
      if (doc && mDocs.has(doc)) return false;
      return zoneIsHome; // ไม่ผูกใบ / เลขใบไม่อยู่ในงวดนี้
    });
    const zFuel = zoneIsHome ? fuel.filter((f) => f.cycleId === cycleId && normPlate(f.plateNo) === key) : [];
    const z = summarizeByVehicle(cycleId, zt, zFuel, zDed, vehicles)[0];
    const zRow = row(z, qtyOf(zt), labels.join(', '));
    zone.push(zRow);
    main.push({
      plateNo: s.plateNo, driverName: s.driverName, qty: round2(qtyOf(pt) - zRow.qty),
      trip: round2(s.totalTripAmount - zRow.trip), d1: round2(s.deduction1Percent - zRow.d1),
      fuel: round2(s.fuelTotal - zRow.fuel), inc: round2(s.incomeAdd - zRow.inc),
      ded: round2(s.deductionTotal - zRow.ded), net: round2(s.netReceive - zRow.net), note: '',
    });
  }
  return { main, zone };
}
