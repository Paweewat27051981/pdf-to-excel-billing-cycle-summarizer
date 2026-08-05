// ============================================================================
// [ทดลอง] ดึงระยะทางถนนจริงจาก DOH TO TRAVEL (Longdo RouteService) — เรียกจาก server เท่านั้น
// 2 ขั้น (ดู memory doh-longdo-distance-api):
//   1) map.doh.go.th/mapsearch/search : ชื่ออำเภอ -> พิกัด (cert ไม่สมบูรณ์ ต้อง rejectUnauthorized:false)
//   2) api.longdo.com/RouteService/guide : พิกัด 2 จุด -> ระยะทาง (ต้องใส่ Referer dohgis.doh.go.th)
// ลูปหลายจุด: เรียง nearest-neighbor (คลัง -> จุดใกล้สุด -> ใกล้สุดถัดไป -> ... -> กลับคลัง)
// มี cache: พิกัดอำเภอ + ระยะคู่จุด (กันยิง API ซ้ำ) | retry 2 ครั้ง
// ============================================================================
import https from 'https';

const KEY = 'dohtotravel-test';
const REFERER = 'https://dohgis.doh.go.th/dohtotravel/';
// insecure agent ใช้ "เฉพาะ map.doh.go.th" (cert chain ไม่ครบ) — host อื่น (Longdo) verify ปกติ
const insecureAgent = new https.Agent({ rejectUnauthorized: false });
const DOH_HOST = 'map.doh.go.th';

export interface LatLon { lat: number; lon: number; }

function httpGet(url: string, headers: Record<string, string> = {}, timeoutMs = 15000): Promise<string> {
  // ปิด TLS verify เฉพาะ host ที่ cert เสียจริง ไม่ปิดทั้งหมด (กัน MITM host อื่น)
  const agent = new URL(url).hostname === DOH_HOST ? insecureAgent : undefined;
  return new Promise((resolve, reject) => {
    const req = https.get(url, { agent, headers, timeout: timeoutMs }, (res) => {
      let d = ''; res.on('data', (c) => (d += c)); res.on('end', () => resolve(d));
    });
    req.on('error', reject);
    req.on('timeout', () => { req.destroy(new Error('timeout')); });
  });
}
async function withRetry<T>(fn: () => Promise<T>, n = 2): Promise<T> {
  let e: unknown;
  for (let i = 0; i <= n; i++) { try { return await fn(); } catch (err) { e = err; if (i < n) await new Promise((r) => setTimeout(r, 400 * (i + 1))); } }
  throw e;
}

// ---- cache (in-memory ตลอดอายุ process) ----
const geoCache = new Map<string, LatLon | null>();     // "อ.เมือง|นครสวรรค์" -> พิกัด
const routeCache = new Map<string, number | null>();   // "lat,lon>lat,lon" -> กม.

/** ยิง DOH search 1 keyword -> พิกัด (ไม่ cache) */
async function dohSearch(keyword: string): Promise<LatLon | null> {
  const url = `https://map.doh.go.th/mapsearch/json/search?keyword=${encodeURIComponent(keyword)}&offset=0&locale=th&dataset=data2p,data2r,data2a,data2b,change,con,m2h&key=${KEY}`;
  const body = await withRetry(() => httpGet(url));
  const d = (JSON.parse(body).data || []).find((x: any) => x && typeof x.lat === 'number' && typeof x.lon === 'number');
  return d ? { lat: d.lat, lon: d.lon } : null;
}

// map ชื่อจังหวัดย่อ -> ชื่อเต็มที่ DOH รู้จัก (ข้อมูลใบใช้ชื่อย่อ)
const PROVINCE_ALIAS: Record<string, string> = {
  'อยุธยา': 'พระนครศรีอยุธยา',
};
/**
 * ค้นชื่ออำเภอ -> พิกัด (cache). key = อำเภอ+จังหวัด
 * ข้อมูลใบมีหลายรูปแบบ (มี "อ."/"เขต " นำหน้า, จังหวัดชื่อย่อ) — ลอง keyword หลายแบบ ตัวแรกที่เจอ = ใช้
 * (ปลายทางนอกพื้นที่สาขา เช่น กทม.ในใบต่างจังหวัด ถูกกรองที่ชั้น endpoint ด้วย serviceArea ก่อนถึงตรงนี้)
 */
export async function geocodeDistrict(district: string, province: string): Promise<LatLon | null> {
  const ck = `${district}|${province}`;
  if (geoCache.has(ck)) return geoCache.get(ck)!;
  // ล้างคำนำหน้าที่ข้อมูลใบใส่มา (อ., เขต, จ.) เหลือชื่อล้วน
  const d = district.replace(/^(อ\.|เขต\s*|อำเภอ\s*)/g, '').trim();
  const prov = province.trim();
  const provFull = PROVINCE_ALIAS[prov] || prov;             // แปลงชื่อย่อ -> เต็ม

  // รายการ keyword ที่จะลองตามลำดับ (dedupe อัตโนมัติ)
  const attempts: string[] = [];
  const add = (s: string) => { if (s && !attempts.includes(s)) attempts.push(s); };
  add(`อ.${d} จ.${provFull}`);           // มาตรฐาน
  if (d === 'เมือง') add(`อ.เมือง${provFull} จ.${provFull}`); // "เมือง" -> ชื่อเต็ม "เมือง<จังหวัด>"
  add(`${d} ${provFull}`);               // fallback สุดท้าย: ชื่อล้วน + จังหวัด (ครอบเคสที่ DOH ไม่รับ อ./จ.)

  try {
    let out: LatLon | null = null;
    for (const kw of attempts) { out = await dohSearch(kw); if (out) break; }
    geoCache.set(ck, out);
    return out;
  } catch { geoCache.set(ck, null); return null; }
}

/** ระยะถนนจริงระหว่าง 2 พิกัด (กม.) — cache + Referer */
export async function roadDistanceKm(from: LatLon, to: LatLon): Promise<number | null> {
  const ck = `${from.lat},${from.lon}>${to.lat},${to.lon}`;
  if (routeCache.has(ck)) return routeCache.get(ck)!;
  const url = `https://api.longdo.com/RouteService/json/route/guide?flon=${from.lon}&flat=${from.lat}&tlon=${to.lon}&tlat=${to.lat}&mode=c&type=127&restrict=0&locale=th&key=${KEY}&maxresult=1`;
  try {
    const body = await withRetry(() => httpGet(url, { Referer: REFERER }));
    const d = (JSON.parse(body).data || [])[0];
    const km = d && typeof d.distance === 'number' ? d.distance / 1000 : null;
    routeCache.set(ck, km);
    return km;
  } catch { routeCache.set(ck, null); return null; }
}

export interface LoopStop { district: string; province: string; loc: LatLon; }
export interface LoopResult {
  order: { district: string; province: string }[]; // ลำดับที่วิ่ง (nearest-neighbor)
  legs: number[];        // ระยะแต่ละขา (กม.) รวมขากลับคลัง
  totalKm: number;       // ระยะลูปรวม
  missing: string[];     // อำเภอที่หาพิกัด/ระยะไม่ได้
}

/**
 * เรียงลูปแบบ nearest-neighbor แล้วรวมระยะ (คลัง -> ใกล้สุด -> ใกล้สุดถัดไป -> ... -> กลับคลัง)
 * @param warehouse พิกัดคลังสาขา (จุดเริ่ม+จุดจบ)
 * @param stops รายอำเภอปลายทาง (พิกัดมาแล้ว)
 */
export async function nearestNeighborLoop(warehouse: LatLon, stops: LoopStop[]): Promise<LoopResult> {
  const order: { district: string; province: string }[] = [];
  const legs: number[] = [];
  const missing: string[] = [];
  const remaining = [...stops];
  let current = warehouse;

  while (remaining.length) {
    // หาจุดใกล้สุดจาก current (ตามระยะถนนจริง)
    let bestIdx = -1, bestKm = Infinity;
    for (let i = 0; i < remaining.length; i++) {
      const km = await roadDistanceKm(current, remaining[i].loc);
      if (km != null && km < bestKm) { bestKm = km; bestIdx = i; }
    }
    if (bestIdx === -1) { // หาระยะจากจุดที่เหลือไม่ได้เลย -> ข้าม
      remaining.forEach((s) => missing.push(`อ.${s.district}`));
      break;
    }
    const next = remaining.splice(bestIdx, 1)[0];
    order.push({ district: next.district, province: next.province });
    legs.push(bestKm);
    current = next.loc;
  }
  // ขากลับคลัง
  if (order.length) {
    const back = await roadDistanceKm(current, warehouse);
    if (back != null) legs.push(back); else missing.push('ขากลับคลัง');
  }
  const totalKm = legs.reduce((s, x) => s + x, 0);
  return { order, legs, totalKm, missing };
}
