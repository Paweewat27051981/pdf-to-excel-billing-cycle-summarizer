#!/usr/bin/env bash
# ============================================================
# NEOSIAM — ติดตั้งเครื่อง Oracle Cloud (Ubuntu 22.04 ARM) ครั้งเดียว
# วิธีใช้: SSH เข้า VM แล้วรัน:  bash setup-vm.sh
# ============================================================
set -e

echo "== 1) อัปเดตระบบ + Node.js 20 =="
sudo apt-get update -y
curl -fsSL https://deb.nodesource.com/setup_20.x | sudo -E bash -
sudo apt-get install -y nodejs git build-essential

echo "== 2) pm2 (คุมให้แอปรันตลอด + รีสตาร์ทเองเวลาเครื่องรีบูต) =="
sudo npm install -g pm2

echo "== 3) Caddy (reverse proxy + ต่อ HTTPS ฟรีอัตโนมัติ) =="
sudo apt-get install -y debian-keyring debian-archive-keyring apt-transport-https curl
curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/gpg.key' | sudo gpg --dearmor -o /usr/share/keyrings/caddy-stable-archive-keyring.gpg
curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/debian.deb.txt' | sudo tee /etc/apt/sources.list.d/caddy-stable.list
sudo apt-get update -y
sudo apt-get install -y caddy

echo "== 4) เปิด firewall port 80/443 (Oracle Ubuntu บล็อกไว้ default) =="
sudo apt-get install -y iptables-persistent netfilter-persistent
sudo iptables -I INPUT 1 -p tcp --dport 80  -j ACCEPT
sudo iptables -I INPUT 1 -p tcp --dport 443 -j ACCEPT
sudo netfilter-persistent save

echo ""
echo "✅ ติดตั้งพื้นฐานเสร็จ — ขั้นต่อไป: clone repo, ใส่ .env + serviceAccountKey.json, แล้ว bash deploy/update.sh"
