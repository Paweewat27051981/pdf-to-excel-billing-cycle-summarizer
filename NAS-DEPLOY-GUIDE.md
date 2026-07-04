# คู่มือย้ายแอป NEOSIAM มารันบน Synology NAS (Container Manager)

> เป้าหมาย: รันแอปบน NAS ควบคู่กับ Render ก่อน ทดสอบให้ชัวร์ แล้วค่อยสลับมาใช้ NAS
> ข้อมูลจริงอยู่บน Firebase (คลาวด์) — NAS แค่ทำหน้าที่ "เซิร์ฟเวอร์" เท่านั้น ไม่ต้องย้ายข้อมูล

---

## ⚠️ กฎเหล็กช่วงรันขนาน (Render + NAS พร้อมกัน)
ทั้งสองเครื่องต่อ **Firebase ก้อนเดียวกัน** และแต่ละเครื่องมี cache ในหน่วยความจำของตัวเอง
👉 **ช่วงทดสอบ ให้กรอก/แก้ข้อมูลจริงที่ "Render ที่เดียว" เท่านั้น** ส่วน NAS เปิดไว้ "ดูอย่างเดียว"
เมื่อพร้อมสลับ (cutover) ค่อยเปลี่ยนมากรอกที่ NAS ที่เดียว แล้วเลิกใช้ Render
(ถ้ากรอกสองที่พร้อมกัน cache จะไม่ตรงกัน เขียนทับกันได้)

---

## Phase 0 — เตรียมของ 2 อย่าง
1. **`serviceAccountKey.json`** (กุญแจ Firebase)
   - เอาจาก Firebase Console > Project Settings > Service accounts > Generate new private key
   - หรือถ้าเก็บไว้ที่ไหนแล้วก็ใช้ไฟล์เดิม
2. **`GEMINI_API_KEY`** — ค่าเดียวกับที่ตั้งไว้บน Render (Render > Environment)

## Phase 1 — ติดตั้ง Container Manager
- DSM > **Package Center** > ค้นหา **Container Manager** > Install

## Phase 2 — build ที่ PC แล้วเอาขึ้น NAS
> ใช้แนวทางเดียวกับที่ deploy logistics KPI สำเร็จ: **build ที่ PC ก่อน** container บน NAS แค่ "รัน"
> (ไม่ build หนักๆ บน NAS = ไม่กิน RAM)

1. ที่เครื่อง PC (PowerShell): build แบบ subpath `/neosiam` — *ผมทำให้แล้ว*
   ```powershell
   $env:VITE_BASE_PATH='/neosiam/'; npm run build
   ```
   *(ถ้า build ใน Git Bash ค่า path จะเพี้ยน — ใช้ PowerShell เท่านั้น)*
2. สร้างโฟลเดอร์บน NAS เช่น `/volume1/docker/neosiam` (ผ่าน File Station)
3. อัปโหลดไฟล์เหล่านี้เข้าโฟลเดอร์นั้น (ไม่ต้องเอา `node_modules` ขึ้น):
   - `dist/` (ทั้งโฟลเดอร์) · `package.json` · `package-lock.json`
   - `Dockerfile` · `docker-compose.yml` · `.dockerignore`
   - `serviceAccountKey.json` (จาก Phase 0)
   - `.env` — ก๊อปจาก `.env.nas.example` แล้วเติม `GEMINI_API_KEY`

## Phase 3 — สร้าง Container (build เบา + run)
- Container Manager > **Project** > **Create**
  - Project name: `neosiam`
  - Path: เลือกโฟลเดอร์ `/volume1/docker/neosiam`
  - ระบบจะเจอ `docker-compose.yml` เอง > **Build**
- ตอน build จะแค่ `npm install` (เบา ไม่ได้ build เว็บใหม่) แล้วรัน container ฟังพอร์ต **3000**
- เช็ก log ในแท็บ Project ว่ามี `✅ ใช้ Firebase Realtime Database` และ `Server running`
  *(สถานะ container ควรขึ้น healthy ภายใน ~1 นาที)*

## Phase 4 — เปิดออกเน็ต (พอร์ตเฉพาะ 8444)
> ⚠️ Synology reverse proxy **ไม่มีช่อง Path** → ทำ subpath บนพอร์ต DSM 5001 ไม่ได้
> ต้องใช้ **พอร์ตเฉพาะ** แบบเดียวกับ logistics ที่ใช้ 8443 → เราใช้ **8444** ให้ neosiam

**4.1 สร้าง Reverse Proxy** (ทำผ่าน DSM จากที่ไหนก็ได้) — ✅ ทำแล้ว
- Control Panel > **Login Portal** > **Advanced** > **Reverse Proxy** > Create
  - **Source**: Protocol `HTTPS`, Hostname `neosiam.dscloud.biz`, Port `8444`
  - **Destination**: Protocol `HTTP`, Hostname `localhost`, Port `3000`
  - Advanced Settings: timeout 300 ทั้ง 3 ช่อง + HTTP 1.1

**4.2 เปิดพอร์ต 8444 ที่เราเตอร์** (ต้องอยู่วงแลนเดียวกับ NAS — เราเตอร์ไม่รองรับ UPnP)
- เข้า `192.168.1.1` (หน้าเราเตอร์) > เมนู Port Forwarding / Virtual Server
- ดูกฎ 8443 เดิมของ logistics เป็นตัวอย่าง แล้วเพิ่มกฎใหม่:

| ช่อง | ค่า |
|---|---|
| External / WAN Port | `8444` |
| Internal IP | `192.168.1.82` |
| Internal Port | `8444` |
| Protocol | `TCP` |

- เปิดเว็บ **https://neosiam.dscloud.biz:8444/neosiam** → ต้องเห็นหน้าล็อกอิน NEOSIAM

## Phase 5 — ทดสอบ (ยังไม่เลิก Render)
- ล็อกอิน เปิดดูรายงานต่อทะเบียน/งวด ว่าข้อมูลครบ ตรงกับ Render
- ลองความไว (ไม่มี cold-start ควรติดทันที)
- **ยังไม่กรอกข้อมูลจริงที่ NAS** (ดูกฎเหล็กด้านบน)

## Phase 6 — สลับมาใช้ NAS (cutover)
1. บอกพนักงานทุกสาขาเปลี่ยนมาเข้า URL ใหม่ `https://neosiam.dscloud.biz:5001/neosiam`
2. แก้ URL ในสคริปต์ backup (`backup-apps-script.gs` > `API_URL`) ให้ชี้ NAS = `https://neosiam.dscloud.biz:5001/neosiam/api/state`
3. เก็บ Render ไว้เป็น **ตัวสำรอง** สัก 1-2 งวด ค่อยปิด

---

## อัปเดตโค้ดใหม่ในอนาคต (ง่ายมาก — dist ถูก mount จากดิสก์)
> compose ตั้ง `- ./dist:/app/dist:ro` ไว้แล้ว → container อ่าน dist จากโฟลเดอร์บน NAS ตรงๆ
> **ไม่ต้อง Build/rebuild อีก** (เลี่ยงปัญหา Docker COPY cache ที่เอา dist ใหม่เข้า image ไม่ได้)

1. ที่ PC (PowerShell): `$env:VITE_BASE_PATH='/neosiam/'; npm run build`
2. เอาโฟลเดอร์ `dist` ใหม่ทับใน `docker/neosiam` (ลบตัวเก่าก่อน แล้วอัป zip + Extract)
3. Container Manager > **Container** `neosiam` > **Stop → Start** (แค่ restart ~5 วิ) — เสิร์ฟโค้ดใหม่ทันที
4. เปิดเว็บ Ctrl+Shift+R (หรือ InPrivate) เช็กผล

## เผื่อพลาด
- **Build ไม่ผ่าน / RAM ไม่พอ**: build อิมเมจบนเครื่อง PC แล้ว export/import เข้า NAS แทน (บอกผมได้ ผมช่วยทำ)
- **เข้าเว็บไม่ได้**: เช็ก container รันอยู่ไหม + Reverse Proxy ปลายทางเป็น `localhost:3000`
- **จอขาว/ข้อมูลไม่มา**: เช็ก log ว่าขึ้น `ใช้ Firebase` ไหม ถ้าไม่ แปลว่า serviceAccountKey.json ไม่ถูกอ่าน
