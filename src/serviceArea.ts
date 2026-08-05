// ============================================================================
// พื้นที่ให้บริการของสาขา (serviceAreaText) — parse + ตรวจ ปลายทางอยู่ในพื้นที่ไหม
// ใช้ร่วม frontend (ตรวจตอนบันทึกใบ) + backend (กรองปลายทางผิดในรายงานต้นทุนน้ำมัน)
// รูปแบบ text: บรรทัดละ "จังหวัด" (ทั้งจังหวัด) หรือ "จังหวัด: อำเภอ1, อำเภอ2" (เฉพาะอำเภอ)
// ============================================================================
const _normP = (s: string) => (s || '').replace(/\s/g, '');

export function parseServiceAreas(text?: string): { prov: string; dists: string[] | null }[] {
  return (text || '').split('\n').map((l) => l.trim()).filter(Boolean).map((line) => {
    const idx = line.search(/[:：]/);
    if (idx < 0) return { prov: line.trim(), dists: null };
    const dists = line.slice(idx + 1).split(/[,，]/).map((s) => s.trim()).filter(Boolean);
    return { prov: line.slice(0, idx).trim(), dists: dists.length ? dists : null };
  });
}

export function inServiceArea(areas: { prov: string; dists: string[] | null }[], prov: string, dist: string): boolean {
  if (!areas.length) return true; // ไม่ตั้งพื้นที่ = ไม่ตรวจ
  const np = _normP(prov), nd = _normP(dist);
  for (const a of areas) {
    const ap = _normP(a.prov);
    if (np && ap && (np.includes(ap) || ap.includes(np))) {
      if (!a.dists) return true; // ทั้งจังหวัด
      return a.dists.some((d) => { const ad = _normP(d); return nd && ad && (nd.includes(ad) || ad.includes(nd)); });
    }
  }
  return false;
}
