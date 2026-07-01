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
// 1.5) Master: สาขา (แต่ละสาขาแยกข้อมูล ราคา/รถ/กฎ + รหัสผ่านเข้าใช้งาน)
// ---------------------------------------------------------------------------
export interface Branch {
  id: string;
  name: string;          // ชื่อสาขา เช่น "นครสวรรค์"
  password: string;      // รหัสผ่านเข้าใช้งานสาขา (ไม่ส่งกลับใน /api/state)
  isHQ?: boolean;        // true = สำนักงานใหญ่ (เห็นทุกสาขา + จัดการสาขา)
  minBoxes?: number | null; // จำนวนกล่องขั้นต่ำต่อเที่ยว (default ของสาขา) — ต่ำกว่านี้ขึ้นเตือน
  // กลุ่มราคา (รถคนละกลุ่มใช้ราคาคนละชุด) + ขั้นต่ำต่อกลุ่ม เช่น เชียงใหม่ กลุ่ม 1=190, กลุ่ม 2=ไม่บังคับ
  rateGroups?: { name: string; minBoxes?: number | null }[];
  // พื้นที่ให้บริการ (ตรวจตอนบันทึก): บรรทัดละ "จังหวัด" (ทั้งจังหวัด) หรือ "จังหวัด: อำเภอ1, อำเภอ2" (เฉพาะอำเภอ) — ว่าง=ไม่ตรวจ
  serviceAreaText?: string;
  // true = "เก็บสินค้าคืน" คิด = ครึ่งของราคาชิ้นปกติของปลายทาง อัตโนมัติ (เช่น นครสวรรค์)
  //        ไม่มีราคาชิ้น -> เตือน+บันทึกไม่ได้. false/ไม่ตั้ง = ใช้ตาราง collect_back แยก (พิษณุโลก/กำแพงเพชร)
  collectBackHalfPiece?: boolean;
  status: RecordStatus;
}

// ---------------------------------------------------------------------------
// 2) Master: รถร่วม + คนขับ
// ---------------------------------------------------------------------------
export interface Vehicle {
  id: string;
  branchId: string;      // สาขาที่รถประจำ
  plateNo: string;       // ทะเบียนรถ
  driverName: string;    // ชื่อคนขับ
  vehicleType: string;   // ประเภทรถ
  rateGroup?: string;    // กลุ่มราคา (รถกลุ่มนี้ใช้ราคาชุดนี้) — ว่าง=ปกติ
  status: RecordStatus;
}

// ---------------------------------------------------------------------------
// 3) Master: ราคาขนส่ง + ประวัติ
// ---------------------------------------------------------------------------
export interface RateMaster {
  id: string;
  branchId: string;        // สาขาเจ้าของราคานี้
  destinationName: string; // ปลายทาง เช่น "อ.เมือง จ.นว"
  provinceName: string;    // จังหวัด เช่น "นครสวรรค์"
  provinceShort: string;   // ตัวย่อจังหวัด เช่น "นว"
  districtName: string;    // อำเภอ เช่น "เมือง"
  priceType: PriceType;    // เหมา / ชิ้น
  price: number;
  // ประเภทสินค้า: normal=งานปกติ, collect_back=เก็บสินค้าคืน(คิดชิ้น), peat_mass=Peat mass(คิดชิ้น), fixed_addon=ค่าเหมาบวกเพิ่มตายตัว(เช่น ท่าสองยาง +700)
  productCategory?: 'normal' | 'collect_back' | 'peat_mass' | 'fixed_addon';
  pieceThreshold?: number | null; // จุดตัดจำนวน: <=จุดตัด ใช้เหมา, >จุดตัด ใช้ชิ้น (เฉพาะปลายทางที่มีทั้ง 2 ราคา)
  // ราคาขั้นบันไดตามจำนวนกล่อง (เช่น CP All ลำพูน 1-150=500, 151+=1200) + ผูกชื่อผู้รับ/ผู้ส่ง/สินค้า
  minQty?: number | null;          // จำนวนกล่องขั้นต่ำที่ใช้ราคานี้
  maxQty?: number | null;          // จำนวนกล่องสูงสุด (null=ไม่จำกัด)
  receiverKeyword?: string;        // ใช้เฉพาะผู้รับที่มีคำนี้ (เช่น "ซีพี ออลล์")
  senderKeyword?: string;          // ใช้เฉพาะผู้ส่งที่มีคำนี้ (เช่น "คูห์เน่")
  productKeyword?: string;         // ใช้เฉพาะสินค้าที่มีคำนี้ (เช่น "สินค้า AD")
  rateGroup?: string;              // กลุ่มราคา (ตามรถ) — ว่าง=ใช้ทุกกลุ่ม
  effectiveFrom: string;   // YYYY-MM-DD
  effectiveTo: string | null;
  status: RecordStatus;
  remark?: string;
  createdBy: string;
  createdAt: string;
  updatedBy?: string;
  updatedAt?: string;
}

// ราคาเฉพาะรอบ (ทับราคาหลักเฉพาะรอบที่ระบุ) — รอบอื่นใช้ราคาหลัก
export interface RateOverride {
  id: string;
  branchId: string;
  cycleId: string;
  rateMasterId: string;          // อ้างอิงราคาหลัก (RateMaster) ที่ถูกทับ
  price: number;
  pieceThreshold?: number | null;
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
  branchId: string;
  groupName: string;
  status: RecordStatus;
}

export interface ReceiverGroupAlias {
  id: string;
  branchId: string;
  receiverGroupId: string;
  aliasName: string;
  status: RecordStatus;
}

// ---------------------------------------------------------------------------
// 5) Master: เงื่อนไขแปลงจำนวนสินค้า (ตัวหาร)
// ---------------------------------------------------------------------------
export interface ProductConversionRule {
  id: string;
  branchId: string;
  ruleName: string;
  senderKeyword: string;        // คำที่ต้องเจอในชื่อผู้ส่ง เช่น "ซีโน"
  receiverKeyword?: string;     // คำที่ต้องเจอในชื่อผู้รับ เช่น "สบายใจรีสอร์ท2" (ว่าง=ทุกผู้รับ, หลายคำคั่น |)
  provinceKeyword?: string;     // คำที่ต้องเจอในจังหวัดปลายทาง เช่น "นครปฐม" (ว่าง=ทุกจังหวัด, หลายคำคั่น |)
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
  branchId: string;
  senderKeyword: string;  // คำในชื่อผู้ส่ง เช่น "คอนซูเมอร์"
  note?: string;
  status: RecordStatus;
}

// ---------------------------------------------------------------------------
// 5.2) Master: แก้ปลายทาง — ใบกระจายระบุจังหวัด/อำเภอผิด ให้คิดราคาตามปลายทางจริง
//      เช่น เจอคำว่า "เจ๊ไพร" (ในชื่อผู้รับ/บรรทัดโน้ต) -> จริงๆ ส่ง พิษณุโลก/เมือง
// ---------------------------------------------------------------------------
export interface DestinationOverride {
  id: string;
  branchId: string;
  keyword: string;        // คำที่ต้องเจอ (ในชื่อผู้รับ หรือ บรรทัดโน้ต *...*) เช่น "เจ๊ไพร"
  province: string;       // จังหวัดจริง เช่น "พิษณุโลก"
  district: string;       // อำเภอจริง เช่น "เมือง"
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
  docNote?: string;       // โน้ตท้ายใบ (เช่น "วิ่งย่อยไม่เกิน 13 จุด") — ใช้เป็น keyword จับราคาแบบมีเงื่อนไข
}

// ---------------------------------------------------------------------------
// 7) ข้อมูลที่บันทึกลง DB (หลัง Review) + ผลคำนวณ
// ---------------------------------------------------------------------------

// การปรับจำนวนระดับ "รายการสินค้าในใบรับ" ที่เข้ากฎตัวหาร
export interface ReceiptAdjustment {
  productName: string;
  originalQty: number;     // จำนวนจริงของรายการนั้น (กลุ่ม = ผลรวม)
  specialQty: number;      // จำนวนสินค้าที่เข้ากฎ (= originalQty)
  divisor: number;
  convertedQty: number;    // ROUND(specialQty / divisor)
  ruleId: string;
  note: string;            // เช่น "1+4 = 5 ÷3 = 2"
  items?: string[];        // ชื่อสินค้าทุกตัวในกลุ่ม (รวมตระกูลเดียวกันในใบรับแล้วหาร)
}

// ใบรับสินค้า (ระดับคำนวณค่าเที่ยวจริง + ปลายทาง/ราคาต่อจุด)
export interface TripReceipt {
  id: string;
  receiptNo: string;
  receiverName: string;
  senderName: string;
  receiverGroupId: string | null;  // กลุ่มที่จับคู่ได้
  totalQty: number;                // จำนวนรวมจริงทุกรายการในใบรับ (ทุกประเภท)
  billingQty: number;              // จำนวนคิดค่าเที่ยว "งานปกติ" (หลังปรับตัวหาร)
  normalQty: number;               // จำนวนงานปกติ (ก่อนหาร)
  collectQty: number;              // จำนวน "เก็บสินค้าคืน"
  collectPrice: number | null;     // ราคาเก็บคืน/ชิ้น (ตามจังหวัด)
  collectFlatPrice: number | null; // ราคาเก็บคืน "เหมา" ของปลายทาง (ใช้เมื่อเก็บคืนอย่างเดียว เทียบกับชิ้นแล้วเลือกสูงกว่า)
  peatQty: number;                 // จำนวน "Peat mass"
  peatPrice: number | null;        // ราคา Peat mass/ชิ้น (ตามอำเภอ, เมื่อ Peat อย่างเดียว)
  hasAdjustment: boolean;          // มีรายการตัวหารหรือไม่
  adjustments: ReceiptAdjustment[];
  divisorSkipped?: { productName: string; qty: number; divisor: number; groupTotal: number }[]; // เข้ากฎตัวหารแต่รวมไม่ถึงตัวหาร (ไฮไลท์ ไม่หาร)
  items: ExtractedReceiptItem[];   // รายการดิบทั้งหมด
  provinceRaw: string;             // ปลายทางของจุดส่งนี้
  districtRaw: string;
  flatPrice: number | null;        // ราคาเหมาของปลายทางนี้
  piecePrice: number | null;       // ราคาชิ้นของปลายทางนี้
  pieceThreshold: number | null;   // จุดตัดจำนวน (ถ้ามี) ของปลายทางนี้
  receiptAmount: number;           // ค่าเที่ยวของจุดนี้ (ตามแบบที่เลือกทั้งใบ)
  requiresManualBox: boolean;      // ผู้ส่งนี้ส่งเป็นชิ้น ต้องกรอกจำนวนกล่องเอง
  manualBoxQty: number | null;     // จำนวนกล่องที่กรอก (ใช้เป็น billingQty)
  destCorrected?: boolean;         // ปลายทางถูกแก้ (ใบกระจายระบุผิด) — คิดราคาตามปลายทางจริง
  origProvince?: string;           // จังหวัดเดิมในใบกระจาย (ก่อนแก้)
  origDistrict?: string;           // อำเภอเดิมในใบกระจาย (ก่อนแก้)
  destFixKeyword?: string;         // คีย์เวิร์ดที่ทำให้แก้ (เช่น "เจ๊ไพร")
  destNote?: string;               // บรรทัดโน้ต "*ส่ง...*" ที่เจอ (เตือนให้ตรวจปลายทาง แม้ยังไม่มีกฎ)
}

// ใบกระจาย (1 คัน 1 เที่ยว)
export interface TripDocument {
  id: string;
  branchId: string;
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
  breakdown: { normal: number; collect: number; peat: number; addon?: number }; // แยกยอด งานปกติ/เก็บคืน/Peat mass/ค่าเหมาบวกเพิ่ม
  addonByDest?: { prov: string; dist: string; amount: number }[]; // ค่าเหมาบวกเพิ่มตายตัว แยกตามปลายทาง (เช่น ท่าสองยาง +700)
  receipts: TripReceipt[];
  warnings: string[];       // รายการแจ้งเตือน/ต้องตรวจสอบ
  docNote?: string;         // โน้ตท้ายใบ (เก็บไว้เพื่อ recalculate + จับราคาแบบมีเงื่อนไข)
  fileName: string;
  isVerified: boolean;
  createdAt: string;
}

// ---------------------------------------------------------------------------
// 8) ค่าน้ำมัน (แยกตามทะเบียน)
// ---------------------------------------------------------------------------
export interface FuelEntry {
  id: string;
  branchId: string;
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
  branchId: string;
  name: string;        // ชื่อแสดงใน dropdown เช่น "ค่าโทรศัพท์", "ค่าอัพเดทบิล"
  kind: MoneyKind;
  status: RecordStatus;
  builtin?: boolean;   // ประเภทของระบบ (ลบไม่ได้)
}

export interface DeductionEntry {
  id: string;
  branchId: string;
  cycleId: string;
  plateNo: string;
  categoryId: string;  // อ้างอิง MoneyCategory
  kind: MoneyKind;     // สำเนา kind ไว้คำนวณเร็ว
  label: string;       // ชื่อแสดง (สำเนาจาก category ตอนบันทึก)
  amount: number;      // จำนวนเงิน (บวกเสมอ, sign กำหนดโดย kind)
  docNo?: string;      // เลขที่ใบกระจายอ้างอิง (เช่น JB0626075363)
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
// การตั้งค่าระบบ (เลือกได้จากหน้าเว็บ)
// ---------------------------------------------------------------------------
export interface AppSettings {
  geminiModel: string; // รุ่นโมเดล Gemini ที่ใช้อ่าน PDF
}

// ---------------------------------------------------------------------------
// Database root
// ---------------------------------------------------------------------------
export interface DatabaseState {
  settings: AppSettings;
  branches: Branch[];
  cycles: BillingCycle[];
  vehicles: Vehicle[];
  rateMasters: RateMaster[];
  rateOverrides: RateOverride[];
  rateMasterHistory: RateMasterHistory[];
  receiverGroups: ReceiverGroup[];
  receiverGroupAliases: ReceiverGroupAlias[];
  conversionRules: ProductConversionRule[];
  manualBoxSenders: ManualBoxSender[];
  destinationOverrides: DestinationOverride[];
  moneyCategories: MoneyCategory[];
  tripDocuments: TripDocument[];
  fuelEntries: FuelEntry[];
  deductions: DeductionEntry[];
}
