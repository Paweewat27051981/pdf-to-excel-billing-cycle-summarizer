# ============================================================
# NEOSIAM — Docker image สำหรับ Synology NAS (Container Manager)
# แนวทางเดียวกับที่ deploy logistics KPI สำเร็จ:
#   build dist/ ที่เครื่อง PC ก่อน (npm run build) แล้ว container แค่ "รัน"
#   -> ไม่ build หนักๆ บน NAS = ปลอดภัยกับ RAM
# ก่อน deploy: รัน `npm run build` ที่ PC ให้มีโฟลเดอร์ dist/ ก่อนเสมอ
# ============================================================
FROM node:20-slim
WORKDIR /app

# ติดตั้งเฉพาะ dependency ที่ใช้ตอนรัน (express, firebase-admin, exceljs ฯลฯ)
COPY package*.json ./
RUN npm install --omit=dev

# เอาโค้ดที่ build เสร็จแล้วเข้าไป (dist/server.cjs = เซิร์ฟเวอร์, dist/ ที่เหลือ = หน้าเว็บ)
COPY dist/ ./dist/

ENV NODE_ENV=production
ENV PORT=3000
EXPOSE 3000
CMD ["node", "dist/server.cjs"]
