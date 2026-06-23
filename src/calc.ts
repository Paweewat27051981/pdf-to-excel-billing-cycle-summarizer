// ============================================================================
// Calculation engine — pure functions (ใช้ได้ทั้ง client และ server)
// หัวใจ: ตัวหารตามใบรับสินค้า, จับคู่ราคาขนส่ง, สูตรรับสุทธิ
// ============================================================================

import {
  RateMaster,
  ReceiverGroup,
  ReceiverGroupAlias,
  ProductConversionRule,
  ManualBoxSender,
  RoundingMethod,
  PriceType,
  Vehicle,
  ExtractedTripDocument,
  TripDocument,
  TripReceipt,
  ReceiptAdjustment,
  ExtractedReceipt,
  ExtractedReceiptItem,
  FuelEntry,
  DeductionEntry,
  VehicleSummary,
  SummaryLine,
  MoneyKind,
  CycleHalf,
} from './types';

// ---------------------------------------------------------------------------
// การปัดเศษ
// ---------------------------------------------------------------------------
export function applyRounding(value: number, method: RoundingMethod): number {
  switch (method) {
    case 'up':
      return Math.ceil(value);
    case 'down':
      return Math.floor(value);
    case 'half_up':
    default:
      // .5 ปัดขึ้น (ROUND_HALF_UP)
      return Math.floor(value + 0.5);
  }
}

// ---------------------------------------------------------------------------
// normalize ข้อความไทย/อังกฤษ สำหรับ keyword match
// ---------------------------------------------------------------------------
function norm(s: string): string {
  return (s || '').toString().toLowerCase().replace(/\s+/g, '').trim();
}

// เทียบทะเบียนรถแบบยืดหยุ่น: ตัดคำว่า "ตู้" (ชนิดรถ) และช่องว่างออกก่อนเทียบ
// เช่น "3ฒต-983 ตู้" (Master) = "3ฒต-983" (อ่านจาก PDF)
export function normPlate(s: string): string {
  return norm(s).replace(/ตู้/g, '');
}

export function textContains(haystack: string, needle: string): boolean {
  if (!needle) return false;
  return norm(haystack).includes(norm(needle));
}

// บรรทัดที่ "ชื่อสินค้ายังไม่ระบุ" เช่น "*** โปรดระบุ ***"
// หมายเหตุ: ยังนับเข้ายอดตามเอกสาร (PDF รวมบรรทัดพวกนี้ในยอดรวมสินค้า) แค่เตือนให้ระบุชื่อ
export function isUnspecifiedName(productName: string): boolean {
  const n = (productName || '').trim();
  if (!n) return true;
  if (/^\*{2,}/.test(n)) return true;            // ขึ้นต้นด้วย ** หรือ ***
  if (textContains(n, 'โปรดระบุ')) return true;
  return false;
}

// จับคู่ขนาดสินค้าแบบยืดหยุ่น: ชื่อจริงใน PDF มักไม่เขียน "42 กรัม" ตรง ๆ
// (เช่น "พริงเกิลส์ PIL 42n.(1x2×6)") จึง fallback ไปเทียบเฉพาะตัวเลขของขนาด
export function sizeMatches(productName: string, sizeKeyword: string): boolean {
  if (!sizeKeyword) return true;
  if (textContains(productName, sizeKeyword)) return true;
  const sizeDigits = sizeKeyword.match(/\d+/g);
  if (sizeDigits && sizeDigits.length) {
    const p = norm(productName);
    return sizeDigits.every((d) => p.includes(d));
  }
  return false;
}

// ---------------------------------------------------------------------------
// แยกรอบ 1-15 / 16-31 จากวันที่
// ---------------------------------------------------------------------------
export function halfOfDate(dateStr: string): CycleHalf | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec((dateStr || '').trim());
  if (!m) return null;
  const day = parseInt(m[3], 10);
  if (day < 1 || day > 31) return null;
  return day <= 15 ? 'first' : 'second';
}

export function isDateInCycle(
  dateStr: string,
  cycle: { year: number; month: number; half: CycleHalf }
): boolean {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec((dateStr || '').trim());
  if (!m) return false;
  const year = parseInt(m[1], 10);
  const month = parseInt(m[2], 10);
  const half = halfOfDate(dateStr);
  return year === cycle.year && month === cycle.month && half === cycle.half;
}

// ---------------------------------------------------------------------------
// effective date check
// ---------------------------------------------------------------------------
function isEffective(
  refDate: string,
  effectiveFrom: string,
  effectiveTo: string | null
): boolean {
  if (!refDate) return true;
  if (effectiveFrom && refDate < effectiveFrom) return false;
  if (effectiveTo && refDate > effectiveTo) return false;
  return true;
}

// ---------------------------------------------------------------------------
// จับคู่กลุ่มผู้รับจากชื่อ (ผ่าน alias)
// ---------------------------------------------------------------------------
export function matchReceiverGroup(
  receiverName: string,
  groups: ReceiverGroup[],
  aliases: ReceiverGroupAlias[]
): ReceiverGroup | null {
  const activeAliases = aliases.filter((a) => a.status === 'active');
  for (const alias of activeAliases) {
    if (textContains(receiverName, alias.aliasName)) {
      const g = groups.find(
        (gr) => gr.id === alias.receiverGroupId && gr.status === 'active'
      );
      if (g) return g;
    }
  }
  // ลองจับจากชื่อกลุ่มตรง ๆ ด้วย
  for (const g of groups) {
    if (g.status === 'active' && textContains(receiverName, g.groupName)) return g;
  }
  return null;
}

// ---------------------------------------------------------------------------
// หากฎตัวหารที่ตรงทุกเงื่อนไข
// เงื่อนไข: ผู้ส่งเข้าซีโน + กลุ่มผู้รับตรง + ชื่อสินค้าตรง + ขนาดตรง
// ---------------------------------------------------------------------------
export function findConversionRule(
  params: {
    senderName: string;
    receiverGroupId: string | null;
    productName: string;
    refDate: string;
  },
  rules: ProductConversionRule[]
): ProductConversionRule | null {
  for (const rule of rules) {
    if (rule.status !== 'active') continue;
    if (!isEffective(params.refDate, rule.effectiveFrom, rule.effectiveTo)) continue;
    if (!textContains(params.senderName, rule.senderKeyword)) continue;
    // ถ้ากฎไม่ระบุกลุ่มผู้รับ (ว่าง) = ใช้กับทุกผู้รับ; ถ้าระบุ ต้องตรงกลุ่ม
    if (rule.receiverGroupId && params.receiverGroupId !== rule.receiverGroupId) continue;
    if (!textContains(params.productName, rule.productKeyword)) continue;
    if (!sizeMatches(params.productName, rule.productSizeKeyword)) continue;
    return rule;
  }
  return null;
}

// ---------------------------------------------------------------------------
// จับคู่ราคาขนส่งจากจังหวัด/อำเภอ + effective date
// คืนราคาเหมา (flat) ก่อน ถ้าไม่มีค่อยใช้ราคาชิ้น (piece)
// ---------------------------------------------------------------------------
export interface RateMatch {
  rateMasterId: string;
  rateType: PriceType;
  rateValue: number;
}

export function matchRate(
  params: { provinceRaw: string; districtRaw: string; refDate: string },
  rates: RateMaster[]
): { flat?: RateMatch; piece?: RateMatch } {
  const candidates = rates.filter((r) => {
    if (r.status !== 'active') return false;
    if (!isEffective(params.refDate, r.effectiveFrom, r.effectiveTo)) return false;
    const provOk =
      textContains(params.provinceRaw, r.provinceName) ||
      textContains(params.provinceRaw, r.provinceShort) ||
      textContains(r.provinceName, params.provinceRaw);
    const distOk =
      !r.districtName ||
      textContains(params.districtRaw, r.districtName) ||
      textContains(r.districtName, params.districtRaw);
    return provOk && distOk;
  });

  const result: { flat?: RateMatch; piece?: RateMatch } = {};
  for (const r of candidates) {
    const match: RateMatch = {
      rateMasterId: r.id,
      rateType: r.priceType,
      rateValue: r.price,
    };
    if (r.priceType === 'flat' && !result.flat) result.flat = match;
    if (r.priceType === 'piece' && !result.piece) result.piece = match;
  }
  return result;
}

// ---------------------------------------------------------------------------
// คำนวณ 1 ใบรับสินค้า: ปรับจำนวนตามตัวหาร -> billingQty
//   special_qty = จำนวนสินค้าที่เข้ากฎ
//   converted_qty = ROUND(special_qty / divisor)
//   billing_qty = total_qty - special_qty + converted_qty
// ---------------------------------------------------------------------------
export function computeReceipt(
  extracted: ExtractedReceipt,
  ctx: {
    groups: ReceiverGroup[];
    aliases: ReceiverGroupAlias[];
    rules: ProductConversionRule[];
    rates: RateMaster[];
    manualBoxSenders: ManualBoxSender[];
    refDate: string;
    fallbackProvince: string;
    fallbackDistrict: string;
  },
  idFactory: () => string
): TripReceipt {
  const group = matchReceiverGroup(extracted.receiverName, ctx.groups, ctx.aliases);
  const groupId = group ? group.id : null;

  // นับทุกบรรทัดที่มีจำนวน (รวมบรรทัดที่ชื่อยังไม่ระบุ "*** โปรดระบุ ***") ให้ตรงกับยอดรวมในใบกระจาย
  const totalQty = extracted.items.reduce((sum, it) => sum + (it.quantity || 0), 0);

  // ผู้ส่งนี้ส่งเป็นชิ้น ต้องกรอกจำนวนกล่องเอง?
  const requiresManualBox = ctx.manualBoxSenders.some(
    (s) => s.status === 'active' && textContains(extracted.senderName, s.senderKeyword)
  );
  const manualBoxQty = typeof extracted.manualBoxQty === 'number' ? extracted.manualBoxQty : null;

  const adjustments: ReceiptAdjustment[] = [];
  let billingQty = totalQty;

  // ถ้าผู้ส่งต้องกรอกกล่องเอง: ใช้จำนวนกล่อง เป็น billingQty (ไม่ใช้ตัวหาร)
  if (requiresManualBox) {
    billingQty = manualBoxQty ?? 0;
  } else {
    // หารแยก "ทีละรายการ": แต่ละรายการที่เข้ากฎ -> qty ÷ divisor แล้วปัด (ROUND_HALF_UP)
    // billing = total - Σ(qty ที่เข้ากฎ) + Σ(ผลหารที่ปัดแล้วของแต่ละรายการ)
    for (const item of extracted.items) {
      const rule = findConversionRule(
        {
          senderName: extracted.senderName,
          receiverGroupId: groupId,
          productName: item.productName,
          refDate: ctx.refDate,
        },
        ctx.rules
      );
      if (!rule) continue;
      const specialQty = item.quantity || 0;
      const convertedQty = applyRounding(specialQty / rule.divisor, rule.roundingMethod);
      billingQty = billingQty - specialQty + convertedQty;
      adjustments.push({
        productName: item.productName,
        originalQty: specialQty,
        specialQty,
        divisor: rule.divisor,
        convertedQty,
        ruleId: rule.id,
        note: `${specialQty} ÷${rule.divisor} = ${convertedQty}`,
      });
    }
  }

  // ปลายทางของจุดส่งนี้ (ถ้าใบรับไม่ระบุ ใช้ของใบกระจาย)
  const provinceRaw = (extracted.provinceRaw || '').trim() || ctx.fallbackProvince;
  const districtRaw = (extracted.districtRaw || '').trim() || ctx.fallbackDistrict;
  const rm = matchRate({ provinceRaw, districtRaw, refDate: ctx.refDate }, ctx.rates);
  const flatPrice = rm.flat ? rm.flat.rateValue : null;
  const piecePrice = rm.piece ? rm.piece.rateValue : null;

  return {
    id: idFactory(),
    receiptNo: extracted.receiptNo,
    receiverName: extracted.receiverName,
    senderName: extracted.senderName,
    receiverGroupId: groupId,
    totalQty,
    billingQty,
    hasAdjustment: adjustments.length > 0,
    adjustments,
    items: extracted.items,
    provinceRaw,
    districtRaw,
    flatPrice,
    piecePrice,
    receiptAmount: 0, // คำนวณจริงในระดับใบกระจาย (หลังรู้ว่าเลือกเหมา/ชิ้น)
    requiresManualBox,
    manualBoxQty,
  };
}

// ---------------------------------------------------------------------------
// คำนวณ 1 ใบกระจาย: รวมใบรับ -> จับราคา -> ค่าเที่ยว + warnings
// ---------------------------------------------------------------------------
export function computeTripDocument(
  extracted: ExtractedTripDocument,
  ctx: {
    cycleId: string;
    cycle: { year: number; month: number; half: CycleHalf };
    vehicles: Vehicle[];
    rates: RateMaster[];
    groups: ReceiverGroup[];
    aliases: ReceiverGroupAlias[];
    rules: ProductConversionRule[];
    manualBoxSenders: ManualBoxSender[];
    fileName: string;
  },
  idFactory: () => string
): TripDocument {
  const warnings: string[] = [];
  const refDate = extracted.documentDate;

  // ตรวจวันที่
  if (!refDate || !halfOfDate(refDate)) {
    warnings.push('อ่านวันที่ไม่ได้ — ต้องแก้ไขก่อนคำนวณ');
  } else if (!isDateInCycle(refDate, ctx.cycle)) {
    warnings.push(`วันที่ ${refDate} อยู่นอกรอบที่เลือก (${ctx.cycle.year}-${String(ctx.cycle.month).padStart(2, '0')} รอบ ${ctx.cycle.half === 'first' ? '1-15' : '16-31'})`);
  }

  // จับรถ (ตัดคำว่า "ตู้" ออกก่อนเทียบ)
  const vehicle = ctx.vehicles.find(
    (v) => normPlate(v.plateNo) === normPlate(extracted.plateNo) && v.status === 'active'
  );
  if (!vehicle) {
    warnings.push(`ทะเบียนรถ "${extracted.plateNo}" ไม่อยู่ใน Master รายชื่อรถร่วม`);
  }

  // คำนวณใบรับ (แต่ละใบมีปลายทาง+ราคาของตัวเอง)
  const receipts: TripReceipt[] = extracted.receipts.map((r) =>
    computeReceipt(
      r,
      {
        groups: ctx.groups, aliases: ctx.aliases, rules: ctx.rules, rates: ctx.rates,
        manualBoxSenders: ctx.manualBoxSenders, refDate,
        fallbackProvince: extracted.provinceRaw, fallbackDistrict: extracted.districtRaw,
      },
      idFactory
    )
  );

  // เตือน: ผู้ส่งต้องกรอกจำนวนกล่อง แต่ยังไม่ได้กรอก
  for (const r of receipts) {
    if (r.requiresManualBox && (r.manualBoxQty == null || r.manualBoxQty <= 0)) {
      warnings.push(`ใบรับ ${r.receiptNo}: ผู้ส่ง "${r.senderName}" ส่งเป็นชิ้น — ต้องกรอกจำนวนกล่องก่อนบันทึก`);
    }
  }

  const totalQty = receipts.reduce((s, r) => s + r.totalQty, 0);
  const billingQty = receipts.reduce((s, r) => s + r.billingQty, 0);

  if (totalQty <= 0) warnings.push('จำนวนสินค้าเป็น 0 หรืออ่านไม่ได้');

  // เตือน: เข้าข่ายกลุ่มแม็คโคร/เซลส์ + สินค้าเข้าข่าย แต่ไม่มีกฎตัวหาร
  for (const r of receipts) {
    for (const it of r.items) {
      const looksSpecial =
        textContains(it.productName, 'ยูปี้') || textContains(it.productName, 'พริงเกิล');
      const hasRuleApplied = r.adjustments.some((a) => a.productName === it.productName);
      if (looksSpecial && r.receiverGroupId && !hasRuleApplied) {
        warnings.push(`สินค้า "${it.productName}" ในใบรับ ${r.receiptNo} เข้าข่ายตัวหารแต่ไม่เจอกฎใน Master`);
      }
    }
  }

  // ราคาที่มีให้เลือก (ระดับใบกระจาย — รวมจากทุกจุด)
  const anyFlat = receipts.some((r) => r.flatPrice != null);
  const anyPiece = receipts.some((r) => r.piecePrice != null);

  // เลือกแบบเดียวกันทั้งใบ: ตามที่ผู้ใช้เลือก ไม่งั้น default = เหมาก่อน
  let rateType: PriceType | null = extracted.rateChoice ?? null;
  if (rateType === 'flat' && !anyFlat) rateType = null;
  if (rateType === 'piece' && !anyPiece) rateType = null;
  if (!rateType) rateType = anyFlat ? 'flat' : anyPiece ? 'piece' : null;

  let tripAmount = 0;

  if (rateType === 'piece') {
    // ชิ้น: แต่ละจุด = billingQty × ราคาชิ้นของจุดนั้น แล้วรวม
    for (const r of receipts) {
      if (r.piecePrice != null) {
        r.receiptAmount = round2(r.billingQty * r.piecePrice);
        tripAmount += r.receiptAmount;
      } else {
        warnings.push(`ปลายทาง "${r.provinceRaw} ${r.districtRaw}" (ใบรับ ${r.receiptNo}) ไม่เจอราคาชิ้น`);
      }
    }
  } else if (rateType === 'flat') {
    // เหมา: คิดราคาเดียว = ราคาเหมา "สูงสุด" ในบรรดาปลายทางทั้งหมด
    let maxFlat: number | null = null;
    for (const r of receipts) {
      if (r.flatPrice != null) {
        r.receiptAmount = r.flatPrice;
        if (maxFlat === null || r.flatPrice > maxFlat) maxFlat = r.flatPrice;
      }
    }
    if (maxFlat === null) warnings.push('ไม่เจอราคาเหมาของปลายทางใดเลย');
    else tripAmount = maxFlat;
  } else {
    warnings.push('ไม่เจอ Master ราคาขนส่งของปลายทาง');
  }

  const rateOptions = {
    flat: anyFlat ? Math.max(...receipts.filter((r) => r.flatPrice != null).map((r) => r.flatPrice as number)) : null,
    piece: anyPiece ? round2(receipts.reduce((s, r) => s + (r.piecePrice != null ? r.billingQty * r.piecePrice : 0), 0)) : null,
  };

  return {
    id: idFactory(),
    branchId: '', // กำหนดในชั้น server ตามสาขาที่บันทึก
    cycleId: ctx.cycleId,
    documentNo: extracted.documentNo,
    documentDate: refDate,
    plateNo: extracted.plateNo,
    driverName: vehicle ? vehicle.driverName : '',
    provinceRaw: extracted.provinceRaw,
    districtRaw: extracted.districtRaw,
    rateMasterId: null,
    rateType,
    rateValue: rateType === 'flat' ? rateOptions.flat : null,
    rateOptions,
    totalQty,
    billingQty,
    tripAmount: round2(tripAmount),
    receipts,
    warnings,
    fileName: ctx.fileName,
    isVerified: false,
    createdAt: new Date().toISOString(),
  };
}

export function round2(n: number): number {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

// ---------------------------------------------------------------------------
// สรุปยอดต่อทะเบียน (สูตรรับสุทธิ)
//   รายได้ทั้งหมด = รวมค่าเที่ยว
//   หัก 1% = รายได้ทั้งหมด * 1%
//   รับสุทธิ = รายได้ - 1% - ค่าน้ำมัน + ค่าอัพเดทบิล
//             - โทรศัพท์ - GPS - ยืมเงิน - GPS รายปี - ประกัน - อื่นๆ
// ---------------------------------------------------------------------------
export function summarizeByVehicle(
  cycleId: string,
  trips: TripDocument[],
  fuel: FuelEntry[],
  deductions: DeductionEntry[],
  vehicles: Vehicle[]
): VehicleSummary[] {
  // รวมทะเบียนจากทุกแหล่ง โดย dedupe ตาม normPlate (เช่น "3ฒต-983" กับ "3ฒต-983 ตู้" = คันเดียวกัน)
  const plateKeys = new Map<string, string>(); // normKey -> display plate (เลือกชื่อจาก Master ก่อน)
  const addPlate = (p: string) => {
    const key = normPlate(p);
    if (!key) return;
    const master = vehicles.find((v) => normPlate(v.plateNo) === key);
    if (master) plateKeys.set(key, master.plateNo);
    else if (!plateKeys.has(key)) plateKeys.set(key, p);
  };
  trips.filter((t) => t.cycleId === cycleId).forEach((t) => addPlate(t.plateNo));
  fuel.filter((f) => f.cycleId === cycleId).forEach((f) => addPlate(f.plateNo));
  deductions.filter((d) => d.cycleId === cycleId).forEach((d) => addPlate(d.plateNo));

  const out: VehicleSummary[] = [];
  for (const [key, plate] of plateKeys) {
    const v = vehicles.find((x) => normPlate(x.plateNo) === key);
    const plateTrips = trips.filter((t) => t.cycleId === cycleId && normPlate(t.plateNo) === key);
    const totalTripAmount = round2(plateTrips.reduce((s, t) => s + t.tripAmount, 0));
    const deduction1Percent = round2(totalTripAmount * 0.01);

    const fuelTotal = round2(
      fuel
        .filter((f) => f.cycleId === cycleId && normPlate(f.plateNo) === key)
        .reduce((s, f) => s + f.amount, 0)
    );

    const plateDeductions = deductions.filter(
      (d) => d.cycleId === cycleId && normPlate(d.plateNo) === key
    );

    // รวมตามประเภท (category) แบบ data-driven
    const lineMap = new Map<string, SummaryLine>();
    for (const d of plateDeductions) {
      // รองรับข้อมูลเก่าที่ยังไม่มี kind/categoryId
      const kind: MoneyKind = d.kind ?? (d.type === 'bill_update' ? 'income' : 'deduction');
      const categoryId = d.categoryId ?? `cat-${d.type ?? 'other'}`;
      const key = categoryId;
      const existing = lineMap.get(key);
      if (existing) existing.amount = round2(existing.amount + d.amount);
      else lineMap.set(key, { categoryId, label: d.label, kind, amount: round2(d.amount) });
    }
    const lines = [...lineMap.values()];
    const incomeAdd = round2(lines.filter((l) => l.kind === 'income').reduce((s, l) => s + l.amount, 0));
    const deductionTotal = round2(lines.filter((l) => l.kind === 'deduction').reduce((s, l) => s + l.amount, 0));

    const netReceive = round2(
      totalTripAmount - deduction1Percent - fuelTotal + incomeAdd - deductionTotal
    );

    out.push({
      plateNo: plate,
      driverName: v ? v.driverName : '',
      totalTripAmount,
      deduction1Percent,
      fuelTotal,
      incomeAdd,
      deductionTotal,
      lines,
      netReceive,
    });
  }
  return out.sort((a, b) => a.plateNo.localeCompare(b.plateNo));
}
