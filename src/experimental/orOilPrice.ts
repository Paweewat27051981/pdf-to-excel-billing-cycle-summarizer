// ============================================================================
// [ทดลอง] ดึงราคาน้ำมันจาก OR/PTT OilPrice XML Web Service (SOAP)
// ⚠️ ข้อควรระวังจริง (พิสูจน์แล้ว):
//   - namespace ของ body element = "http://www.pttor.com" (ไม่ใช่ orapiweb!)
//   - ผลลัพธ์เป็น string ที่มี XML (<PTTOR_DS>) ซ้อนอยู่ + HTML entity (&lt; &gt;) -> ต้อง decode + parse 2 ชั้น
//   - service เป็น SOAP เท่านั้น (GET/POST form ปิด)
// เรียกจาก server เท่านั้น (browser เรียก cross-origin SOAP ไม่ได้)
// ============================================================================

export const OR_ENDPOINT = 'https://orapiweb.pttor.com/oilservice/OilPrice.asmx';
export const OR_NAMESPACE = 'http://www.pttor.com';

export interface OilPrice {
  product: string;      // ชื่อผลิตภัณฑ์ เช่น "ดีเซล", "ดีเซล B20"
  price: number;        // บาท/ลิตร
  priceDate: string;    // เช่น 2026-07-23T05:00
  location?: string;    // อำเภอ (เฉพาะราคารายจังหวัด <FUEL_PROVINCIAL>) เช่น "เมืองนครสวรรค์"
}

/** decode HTML entity พื้นฐาน (&lt; &gt; &amp; &quot; &#39;) */
function decodeEntities(s: string): string {
  return s
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/&amp;/g, '&'); // &amp; ท้ายสุด กัน double-decode
}

/** ดึงค่าใน tag แรกที่เจอ (ไม่สน namespace)
 * ⚠️ หลังชื่อ tag ต้องเป็น ">" หรือช่องว่าง(attribute) เท่านั้น — กัน <PRICE> ไปแมตช์ <PRICE_DATE> */
function pick(xml: string, tag: string): string {
  const m = new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/${tag}>`, 'i').exec(xml);
  return m ? m[1].trim() : '';
}

/**
 * parse SOAP response -> รายการราคาน้ำมัน (pure, ทดสอบได้โดยไม่ต้องต่อเน็ต)
 * รองรับทั้ง CurrentOilPrice / CurrentOilPriceProvincial / GetOilPrice* (โครง <FUEL> เหมือนกัน)
 */
export function parseOilPriceResult(soapXml: string): OilPrice[] {
  // 1) ดึงเนื้อใน <...Result>...</...Result> (ชื่อ tag ต่างกันตาม operation)
  const resultRaw = /<[A-Za-z]*OilPrice[A-Za-z]*Result>([\s\S]*?)<\/[A-Za-z]*OilPrice[A-Za-z]*Result>/i.exec(soapXml);
  const inner = resultRaw ? decodeEntities(resultRaw[1]) : decodeEntities(soapXml);
  // 2) parse ทีละ block — รองรับ 2 โครง: <FUEL> (กทม.) และ <FUEL_PROVINCIAL> (รายจังหวัด มี <LOCATION>=อำเภอ)
  //    ห้ามผูกกับลำดับ/ index
  const out: OilPrice[] = [];
  const blocks = inner.match(/<FUEL(?:_PROVINCIAL)?>[\s\S]*?<\/FUEL(?:_PROVINCIAL)?>/gi) || [];
  for (const b of blocks) {
    const product = pick(b, 'PRODUCT');
    const price = Number(pick(b, 'PRICE'));
    if (product && Number.isFinite(price)) {
      const location = pick(b, 'LOCATION');
      out.push({ product, price, priceDate: pick(b, 'PRICE_DATE'), ...(location ? { location } : {}) });
    }
  }
  return out;
}

/** เลือกราคา "ดีเซล" ธรรมดา (ตรงตัว ไม่ใช่ B20/Super Power) จากรายการ — mapping แก้ได้ */
export function pickDiesel(prices: OilPrice[], productName = 'ดีเซล'): OilPrice | null {
  // ตรงชื่อเป๊ะก่อน (กัน "ดีเซล B20" / "Super Power Diesel")
  return prices.find((p) => p.product.trim() === productName) || null;
}

/** สรุปราคาดีเซลระดับจังหวัด — เลือกตัวแทนตามอำเภอ (default "เมือง"), พร้อมช่วงราคาทุกอำเภอ */
export function summarizeProvinceDiesel(prices: OilPrice[], productName = 'ดีเซล', repDistrict = 'เมือง'): {
  representative: OilPrice | null; min: number | null; max: number | null; byLocation: OilPrice[];
} {
  const diesel = prices.filter((p) => p.product.trim() === productName);
  if (!diesel.length) return { representative: null, min: null, max: null, byLocation: [] };
  // ตัวแทน: อำเภอที่ระบุ (เช่น เมือง / แม่สอด) ก่อน ไม่งั้นตัวแรก (policy เปลี่ยนได้)
  const representative = diesel.find((d) => (d.location || '').includes(repDistrict)) || diesel[0];
  const vals = diesel.map((d) => d.price);
  return { representative, min: Math.min(...vals), max: Math.max(...vals), byLocation: diesel };
}

// map สาขาจริง -> ราคาน้ำมัน OR ที่ใช้ (policy แก้ได้) — ยืนยันกับผู้ใช้แล้ว
//  - สาขาภูมิภาค: ใช้ราคา อ.เมือง ของจังหวัดตัวเอง (kind='province')
//  - แม่สอด = อ.แม่สอด จ.ตาก (ไม่ใช่จังหวัด!)
//  - สาย3 = กทม./ปริมณฑล (ใช้ CurrentOilPrice ไม่ใช่รายจังหวัด, kind='bkk')
export interface BranchOilConfig { branch: string; kind: 'province' | 'bkk'; province?: string; repDistrict?: string; }
export const BRANCH_OIL_CONFIGS: BranchOilConfig[] = [
  { branch: 'สาย3', kind: 'bkk' },
  { branch: 'นครสวรรค์', kind: 'province', province: 'นครสวรรค์', repDistrict: 'เมือง' },
  { branch: 'กำแพงเพชร', kind: 'province', province: 'กำแพงเพชร', repDistrict: 'เมือง' },
  { branch: 'พิษณุโลก', kind: 'province', province: 'พิษณุโลก', repDistrict: 'เมือง' },
  { branch: 'แม่สอด', kind: 'province', province: 'ตาก', repDistrict: 'แม่สอด' },
  { branch: 'เชียงใหม่', kind: 'province', province: 'เชียงใหม่', repDistrict: 'เมือง' },
];

function soapEnvelope(operation: string, params: Record<string, string>): string {
  const inner = Object.entries(params).map(([k, v]) => `<${k}>${v}</${k}>`).join('');
  return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body>
    <${operation} xmlns="${OR_NAMESPACE}">${inner}</${operation}>
  </soap:Body>
</soap:Envelope>`;
}

async function callOr(operation: string, params: Record<string, string>, timeoutMs = 15000): Promise<OilPrice[]> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetch(OR_ENDPOINT, {
      method: 'POST',
      headers: {
        'Content-Type': 'text/xml; charset=utf-8',
        SOAPAction: `"https://orapiweb.pttor.com/${operation}"`,
      },
      body: soapEnvelope(operation, params),
      signal: controller.signal,
    });
    if (!res.ok) throw new Error(`OR SOAP ${operation} HTTP ${res.status}`);
    return parseOilPriceResult(await res.text());
  } finally {
    clearTimeout(timer);
  }
}

/** retry แบบ exponential backoff สูงสุด maxAttempts ครั้ง */
async function withRetry<T>(fn: () => Promise<T>, maxAttempts = 3): Promise<T> {
  let lastErr: unknown;
  for (let i = 0; i < maxAttempts; i++) {
    try { return await fn(); }
    catch (e) { lastErr = e; if (i < maxAttempts - 1) await new Promise((r) => setTimeout(r, 500 * 2 ** i)); }
  }
  throw lastErr;
}

/** ราคาปัจจุบัน กทม./ปริมณฑล */
export function fetchCurrentOilPrice(language = 'TH'): Promise<OilPrice[]> {
  return withRetry(() => callOr('CurrentOilPrice', { Language: language }));
}

/** ดึง+สรุปราคาดีเซลของ 1 สาขา ตาม config (province ใช้ อ.เมือง/แม่สอด, bkk ใช้ กทม.)
 * คืน { representative, min, max } — ใช้ร่วมกันทั้ง endpoint และ scheduler (ไม่ซ้ำโค้ด) */
export async function fetchBranchDiesel(
  cfg: BranchOilConfig, date?: { dd: string; mm: string; yyyy: string },
): Promise<{ representative: OilPrice | null; min: number | null; max: number | null; districts: number }> {
  if (cfg.kind === 'bkk') {
    // กทม./ปริมณฑล: ไม่มีรายอำเภอ -> ดีเซลตัวเดียว (ย้อนหลังใช้ GetOilPrice)
    const prices = date ? await fetchHistoricalOilPrice(date.dd, date.mm, date.yyyy) : await fetchCurrentOilPrice();
    const rep = pickDiesel(prices);
    return { representative: rep, min: rep?.price ?? null, max: rep?.price ?? null, districts: rep ? 1 : 0 };
  }
  const prices = date
    ? await fetchHistoricalProvincialOilPrice(cfg.province!, date.dd, date.mm, date.yyyy)
    : await fetchProvincialOilPrice(cfg.province!);
  const s = summarizeProvinceDiesel(prices, 'ดีเซล', cfg.repDistrict);
  return { representative: s.representative, min: s.min, max: s.max, districts: s.byLocation.length };
}

/** ราคาย้อนหลัง กทม./ปริมณฑล */
export function fetchHistoricalOilPrice(dd: string, mm: string, yyyy: string, language = 'TH'): Promise<OilPrice[]> {
  return withRetry(() => callOr('GetOilPrice', { Language: language, DD: dd, MM: mm, YYYY: yyyy }));
}

/** ราคาปัจจุบันรายจังหวัด */
export function fetchProvincialOilPrice(province: string, language = 'TH'): Promise<OilPrice[]> {
  return withRetry(() => callOr('CurrentOilPriceProvincial', { Language: language, Province: province }));
}

/** ราคาย้อนหลังรายจังหวัด (DD MM YYYY) */
export function fetchHistoricalProvincialOilPrice(province: string, dd: string, mm: string, yyyy: string, language = 'TH'): Promise<OilPrice[]> {
  return withRetry(() => callOr('GetOilPriceProvincial', { Language: language, DD: dd, MM: mm, YYYY: yyyy, Province: province }));
}
