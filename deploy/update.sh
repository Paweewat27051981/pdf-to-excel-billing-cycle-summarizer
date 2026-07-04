#!/usr/bin/env bash
# ============================================================
# NEOSIAM — อัปเดต/รันแอปบน Oracle VM (ใช้ซ้ำได้ทุกครั้งที่แก้โค้ด)
# วิธีใช้: จากโฟลเดอร์โปรเจกต์ รัน:  bash deploy/update.sh
# ============================================================
set -e
cd "$(dirname "$0")/.."

echo "== ดึงโค้ดล่าสุด =="
git pull

echo "== ติดตั้ง dependency + build =="
npm install
npm run build

echo "== สั่ง pm2 รัน/รีสตาร์ท =="
if pm2 describe neosiam > /dev/null 2>&1; then
  pm2 restart neosiam --update-env
else
  pm2 start deploy/ecosystem.config.cjs
  pm2 save
  # ตั้งให้ pm2 เด้งเองหลังเครื่องรีบูต (รันคำสั่งที่มันพิมพ์ออกมา 1 ครั้ง)
  pm2 startup systemd -u "$USER" --hp "$HOME" || true
fi

echo "✅ เสร็จ — ดู log: pm2 logs neosiam"
