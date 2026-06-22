// ============================================================================
// Domain: ระบบคำนวณค่าเที่ยว + ค่าน้ำมันรถร่วม
// PDF ใบกระจาย -> Review -> คำนวณรอบ 1-15 / 16-31 -> Export Excel
// ============================================================================

export type RecordStatus = 'active' | 'inactive';
export type RoundingMethod = 'half_up' | 'up' | 'down'; // .5 ปัดขึ้น / ปัดขึ้นเสมอ / ปัดลง
export type PriceType = 'flat' | 'piece'; // ราคาเหมา / ราคาชิ้น
export type CycleHalf = 'first' | 'second'; // 1-15 / 16-31

// ---------------------------------------------------------------------------
// 1) รอบการคำนวณ (เดือน + ครึ่งเดือน)
// ---------------------------------------------------------------------------
export interface BillingCycle {
  id: string;
  name: string;          // เช่น "พ.ค. 69 รอบ 1-15"
  year: number;          // ค.ศ. เช่น 2026
  month: number;         // 1-12
  half: CycleHalf;       // first = 1-15, second = 16-31
  startDate: string;     // YYYY-MM-DD
  endDate: string;       // YYYY-MM-DD
  status: 'open' | 'closed';
  createdAt: string;
}

// ---------------------------------------------------------------------------
// 2) Master: รถร่วม + คนขับ
// ---------------------------------------------------------------------------
export interface Vehicle {
  id: string;
  plateNo: string;       // ทะเบียนรถ
  driverName: string;    // ชื่อคนขับ
  vehicleType: string;   // ประเภทรถ
  status: RecordStatus;
}

// ---------------------------------------------------------------------------
// 3) Master: ราคาขนส่ง + ประวัติ
// ---------------------------------------------------------------------------
export interface RateMaster {
  id: string;
  destinationName: string; // ปลายทาง เช่น "อ.เมือง จ.นว"
  provinceName: string;    // จังหวัด เช่น "นครสวรรค์"
  provinceShort: string;   // ตัวย่อจังหวัด เช่น "นว"
  districtName: string;    // อำเภอ เช่น "เมือง"
  priceType: PriceType;    // เหมา / ชิ้น
  price: number;
  effectiveFrom: string;   // YYYY-MM-DD
  effectiveTo: string | null;
  status: RecordStatus;
  remark?: string;
  createdBy: string;
  createdAt: string;
  updatedBy?: string;
  updatedAt?: string;
}

export interface RateMasterHistory {
  id: string;
  rateMasterId: string;
  oldPrice: number;
  newPrice: number;
  changedBy: string;
  changedAt: string;
  changeReason: string;
}

// ---------------------------------------------------------------------------
// 4) Master: กลุ่มผู้รับสินค้า + ชื่อพ้อง (alias)
// เช่น กลุ่ม "แม็คโคร/เซลส์" มี alias: แม็คโคร, MK, CP AXTRA, ซีพี แอ็กซ์ตร้า
// ---------------------------------------------------------------------------
export interface ReceiverGroup {
  id: string;
  groupName: string;
  status: RecordStatus;
}

export interface ReceiverGroupAlias {
  id: string;
  receiverGroupId: string;
  aliasName: string;
  status: RecordStatus;
}

// ---------------------------------------------------------------------------
// 5) Master: เงื่อนไขแปลงจำนวนสินค้า (ตัวหาร)
// ---------------------------------------------------------------------------
export interface ProductConversionRule {
  id: string;
  ruleName: string;
  senderKeyword: string;        // คำที่ต้องเจอในชื่อผู้ส่ง เช่น "ซีโน"
  receiverGroupId: string;      // กลุ่มผู้รับที่ใช้กฎ (เช่น แม็คโคร/เซลส์)
  productKeyword: string;       // เช่น "ยูปี้" หรือ "พริงเกิล"
  productSizeKeyword: string;   // เช่น "14 กรัม" / "42 กรัม"
  divisor: number;              // ตัวหาร เช่น 3
  roundingMethod: RoundingMethod;
  applyLevel: 'receipt';        // ต้องคำนวณแยกตามเลขที่ใบรับสินค้า
  status: RecordStatus;
  effectiveFrom: string;
  effectiveTo: string | null;
  remark?: string;
}

// ---------------------------------------------------------------------------
// 5.1) Master: ผู้ส่งที่ส่งเป็น "ชิ้น" ต้องกรอกจำนวนกล่องเอง
// ---------------------------------------------------------------------------
export interface ManualBoxSender {
  id: string;
  senderKeyword: string;  // คำในชื่อผู้ส่ง เช่น "คอนซูเมอร์"
  note?: string;
  status: RecordStatus;
}

// ---------------------------------------------------------------------------
// 6) ผลการอ่าน PDF ใบกระจาย (จาก AI ก่อน Review)
// ---------------------------------------------------------------------------
export interface ExtractedReceiptItem {
  productName: string;   // รายการสินค้า
  quantity: number;      // จำนวน
  unit?: string;         // หน่วยนับ เช่น กล่อง, หีบ, ลัง
}

export interface ExtractedReceipt {
  receiptNo: string;             // เลขที่ใบรับสินค้า
  receiverName: string;          // ผู้รับสินค้า
  senderName: string;            // ผู้ส่งสินค้า
  items: ExtractedReceiptItem[];
  provinceRaw?: string;          // ปลายทางของจุดส่งนี้ (ถ้าว่าง ใช้ของใบกระจาย)
  districtRaw?: string;
  manualBoxQty?: number;         // จำนวนกล่องที่ผู้ใช้กรอกเอง (เฉพาะผู้ส่งที่ส่งเป็นชิ้น)
}

export interface ExtractedTripDocument {
  documentNo: string;     // เลขที่ใบกระจาย
  documentDate: string;   // YYYY-MM-DD (วันที่ออก)
  plateNo: string;        // ทะเบียนรถ
  provinceRaw: string;    // จังหวัด (ดิบ)
  districtRaw: string;    // อำเภอ (ดิบ)
  receipts: ExtractedReceipt[];
  rateChoice?: PriceType; // ผู้ใช้เลือกใช้ราคาเหมา/ชิ้น (ถ้าปลายทางมีทั้งคู่)
}

// ---------------------------------------------------------------------------
// 7) ข้อมูลที่บันทึกลง DB (หลัง Review) + ผลคำนวณ
// ---------------------------------------------------------------------------

// การปรับจำนวนระดับ "รายการสินค้าในใบรับ" ที่เข้ากฎตัวหาร
export interface ReceiptAdjustment {
  productName: string;
  originalQty: number;     // จำนวนจริงของรายการนั้น
  specialQty: number;      // จำนวนสินค้าที่เข้ากฎ (= originalQty)
  divisor: number;
  convertedQty: number;    // ROUND(specialQty / divisor)
  ruleId: string;
  note: string;            // เช่น "100-7+2 = 95"
}

// ใบรับสินค้า (ระดับคำนวณค่าเที่ยวจริง + ปลายทาง/ราคาต่อจุด)
export interface TripReceipt {
  id: string;
  receiptNo: string;
  receiverName: string;
  senderName: string;
  receiverGroupId: string | null;  // กลุ่มที่จับคู่ได้
  totalQty: number;                // จำนวนรวมจริงทุกรายการในใบรับ
  billingQty: number;              // จำนวนคิดค่าเที่ยว (หลังปรับตัวหาร)
  hasAdjustment: boolean;          // มีรายการตัวหารหรือไม่
  adjustments: ReceiptAdjustment[];
  items: ExtractedReceiptItem[];   // รายการดิบทั้งหมด
  provinceRaw: string;             // ปลายทางของจุดส่งนี้
  districtRaw: string;
  flatPrice: number | null;        // ราคาเหมาของปลายทางนี้
  piecePrice: number | null;       // ราคาชิ้นของปลายทางนี้
  receiptAmount: number;           // ค่าเที่ยวของจุดนี้ (ตามแบบที่เลือกทั้งใบ)
  requiresManualBox: boolean;      // ผู้ส่งนี้ส่งเป็นชิ้น ต้องกรอกจำนวนกล่องเอง
  manualBoxQty: number | null;     // จำนวนกล่องที่กรอก (ใช้เป็น billingQty)
}

// ใบกระจาย (1 คัน 1 เที่ยว)
export interface TripDocument {
  id: string;
  cycleId: string;
  documentNo: string;
  documentDate: string;     // YYYY-MM-DD
  plateNo: string;
  driverName: string;       // จับจาก vehicle master
  provinceRaw: string;
  districtRaw: string;
  rateMasterId: string | null;
  rateType: PriceType | null;
  rateValue: number | null; // ราคาที่ใช้ (เหมา หรือ ราคา/ชิ้น)
  rateOptions: { flat: number | null; piece: number | null }; // ราคาที่เลือกได้สำหรับปลายทางนี้
  totalQty: number;         // รวมจำนวนจริงทุกใบรับ
  billingQty: number;       // รวมจำนวนคิดค่าเที่ยว
  tripAmount: number;       // ค่าเที่ยว = billingQty * piece OR flat
  receipts: TripReceipt[];
  warnings: string[];       // รายการแจ้งเตือน/ต้องตรวจสอบ
  fileName: string;
  isVerified: boolean;
  createdAt: string;
}

// ---------------------------------------------------------------------------
// 8) ค่าน้ำมัน (แยกตามทะเบียน)
// ---------------------------------------------------------------------------
export interface FuelEntry {
  id: string;
  cycleId: string;
  plateNo: string;
  refNo: string;        // เลขใบสั่งเติม
  date: string;         // YYYY-MM-DD
  amount: number;       // ค่าน้ำมัน
  note?: string;
}

// ---------------------------------------------------------------------------
// 9) รายการหัก (ต่อทะเบียน/ต่อรอบ)
// ค่าอัพเดทบิล(+), โทรศัพท์, GPS, ยืมเงิน, GPS รายปี, ประกัน
// ---------------------------------------------------------------------------
export type DeductionType =
  | 'bill_update'   // ค่าอัพเดทบิล (บวกกลับ) [legacy]
  | 'phone' | 'gps' | 'loan' | 'gps_yearly' | 'insurance' | 'other';

// ทิศทางของเงิน: income = รายได้เพิ่ม (+), deduction = หักออก (-)
export type MoneyKind = 'income' | 'deduction';

// Master ประเภทรายการเงิน (เพิ่ม/แก้ได้จากหน้าจอ ไม่ต้องแก้โค้ด)
export interface MoneyCategory {
  id: string;
  name: string;        // ชื่อแสดงใน dropdown เช่น "ค่าโทรศัพท์", "ค่าอัพเดทบิล"
  kind: MoneyKind;
  status: RecordStatus;
  builtin?: boolean;   // ประเภทของระบบ (ลบไม่ได้)
}

export interface DeductionEntry {
  id: string;
  cycleId: string;
  plateNo: string;
  categoryId: string;  // อ้างอิง MoneyCategory
  kind: MoneyKind;     // สำเนา kind ไว้คำนวณเร็ว
  label: string;       // ชื่อแสดง (สำเนาจาก category ตอนบันทึก)
  amount: number;      // จำนวนเงิน (บวกเสมอ, sign กำหนดโดย kind)
  note?: string;
  type?: DeductionType; // legacy
}

// ---------------------------------------------------------------------------
// 10) สรุปยอดต่อทะเบียน (ผลคำนวณ)
// ---------------------------------------------------------------------------
export interface SummaryLine {
  categoryId: string;
  label: string;
  kind: MoneyKind;
  amount: number;
}

export interface VehicleSummary {
  plateNo: string;
  driverName: string;
  totalTripAmount: number;    // รวมค่าเที่ยวทั้งหมด (รายได้ทั้งหมด)
  deduction1Percent: number;  // หัก 1%
  fuelTotal: number;          // ค่าน้ำมัน
  incomeAdd: number;          // รวมรายได้เพิ่ม (+) เช่น ค่าอัพเดทบิล
  deductionTotal: number;     // รวมรายการหัก (-)
  lines: SummaryLine[];       // รายละเอียดแยกประเภท (income/deduction)
  netReceive: number;         // รวมรับสุทธิ
}

// ---------------------------------------------------------------------------
// Database root
// ---------------------------------------------------------------------------
export interface DatabaseState {
  cycles: BillingCycle[];
  vehicles: Vehicle[];
  rateMasters: RateMaster[];
  rateMasterHistory: RateMasterHistory[];
  receiverGroups: ReceiverGroup[];
  receiverGroupAliases: ReceiverGroupAlias[];
  conversionRules: ProductConversionRule[];
  manualBoxSenders: ManualBoxSender[];
  moneyCategories: MoneyCategory[];
  tripDocuments: TripDocument[];
  fuelEntries: FuelEntry[];
  deductions: DeductionEntry[];
}
