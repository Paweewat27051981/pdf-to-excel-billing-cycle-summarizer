// ============================================================================
// [ทดลอง] Scheduler บันทึกราคาน้ำมัน OR อัตโนมัติ ทุกวัน 05:30 เวลาไทย (Asia/Bangkok)
// ไม่ใช้ cron lib — เช็คนาทีปัจจุบันในเขตเวลาไทยด้วย Intl (กัน timezone ของ server เพี้ยน)
// กันรันซ้ำในนาทีเดียวกันด้วย lastRunKey
// ============================================================================
import { fetchBranchDiesel, BRANCH_OIL_CONFIGS } from './orOilPrice.js';
import type { OilPriceRecord } from '../types.js';

const RUN_HH = 5, RUN_MM = 30; // 05:30 น.

/** เวลาไทยตอนนี้ -> { hh, mm, dateKey } โดยไม่พึ่ง timezone ของ server */
function bangkokNow(): { hh: number; mm: number; dateKey: string } {
  const parts = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Bangkok', hour: '2-digit', minute: '2-digit', year: 'numeric', month: '2-digit', day: '2-digit', hour12: false,
  }).formatToParts(new Date());
  const g = (t: string) => parts.find((p) => p.type === t)?.value || '';
  return { hh: Number(g('hour')), mm: Number(g('minute')), dateKey: `${g('year')}-${g('month')}-${g('day')}` };
}

async function saveAllProvinces(
  deps: { getDb: () => Promise<any>; saveRecord: (c: any, r: { id: string }) => Promise<void>; genId: (p: string) => string },
): Promise<{ saved: number; skipped: string[] }> {
  const db = await deps.getDb();
  if (!Array.isArray(db.oilPrices)) db.oilPrices = [];
  const now = new Date().toISOString();
  let saved = 0; const skipped: string[] = [];
  for (const c of BRANCH_OIL_CONFIGS) {
    try {
      const r = await fetchBranchDiesel(c);
      if (!r.representative) { skipped.push(c.branch); continue; }
      const priceDate = String(r.representative.priceDate || '').slice(0, 10);
      const existing = db.oilPrices.find((x: OilPriceRecord) => x.branch === c.branch && x.priceDate === priceDate);
      const rec: OilPriceRecord = {
        id: existing?.id || deps.genId('oil'),
        branch: c.branch, province: c.province || 'กทม./ปริมณฑล',
        location: r.representative.location || (c.kind === 'bkk' ? 'กทม./ปริมณฑล' : ''), product: r.representative.product,
        price: r.representative.price, priceDate, minPrice: r.min ?? undefined, maxPrice: r.max ?? undefined,
        source: 'OR', fetchedAt: now, savedBy: 'auto-05:30',
      };
      if (existing) Object.assign(existing, rec); else db.oilPrices.push(rec);
      await deps.saveRecord('oilPrices', rec);
      saved++;
    } catch (e: any) { skipped.push(`${c.branch} (${e.message})`); }
  }
  return { saved, skipped };
}

/** เริ่ม scheduler — เช็คทุกนาที ถ้าตรง 05:30 เวลาไทย (และยังไม่รันวันนี้) -> บันทึก */
export function startOilPriceScheduler(
  deps: { getDb: () => Promise<any>; saveRecord: (c: any, r: { id: string }) => Promise<void>; genId: (p: string) => string },
): void {
  let lastSuccessDate = ''; // ตั้งเมื่อบันทึก "สำเร็จ" เท่านั้น -> ถ้า 05:30 ล่ม นาทีถัดไปยังลองใหม่ได้
  let running = false;      // กันรันซ้อน (tick ทุกนาทีอาจทับกันถ้า OR ช้า)
  const tick = async () => {
    const { hh, mm, dateKey } = bangkokNow();
    // ยิงช่วง 05:30–05:39 (เผื่อ retry ในวันเดียวถ้ารอบแรกล่ม) จนกว่าจะสำเร็จ 1 ครั้ง/วัน
    const inWindow = hh === RUN_HH && mm >= RUN_MM && mm < RUN_MM + 10;
    if (!inWindow || running || lastSuccessDate === dateKey) return;
    running = true;
    try {
      const { saved, skipped } = await saveAllProvinces(deps);
      // ตั้ง done เฉพาะเมื่อ "ครบทุกสาขา ไม่มี skip" -> ถ้าบางสาขาล่ม ยัง retry ในนาทีถัดไป (จนหมด window)
      if (saved > 0 && skipped.length === 0) {
        lastSuccessDate = dateKey;
        console.log(`[oil-scheduler] ${dateKey} ${hh}:${mm} ไทย: บันทึกครบ ${saved} สาขา`);
      } else {
        console.error(`[oil-scheduler] ${dateKey} ${hh}:${mm}: บันทึก ${saved} สาขา แต่ยังเหลือ (${skipped.join(', ') || 'ทั้งหมด'}) — จะลองใหม่ในนาทีถัดไป`);
      }
    } catch (e: any) {
      console.error('[oil-scheduler] error (จะลองใหม่):', e.message);
    } finally { running = false; }
  };
  setInterval(tick, 60 * 1000); // เช็คทุก 1 นาที
  console.log('[oil-scheduler] เริ่มแล้ว — บันทึกราคาน้ำมันอัตโนมัติ 05:30 น. เวลาไทยทุกวัน');
}
