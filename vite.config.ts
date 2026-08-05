import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import path from 'path';
import {defineConfig} from 'vite';

export default defineConfig(() => {
  return {
    // รันใต้ subpath ได้ (เช่น NAS: VITE_BASE_PATH=/neosiam/) — ไม่ตั้ง = root สำหรับ Render
    base: process.env.VITE_BASE_PATH || '/',
    plugins: [react(), tailwindcss()],
    resolve: {
      alias: {
        '@': path.resolve(__dirname, '.'),
      },
    },
    server: {
      // HMR is disabled in AI Studio via DISABLE_HMR env var.
      // Do not modifyâfile watching is disabled to prevent flickering during agent edits.
      hmr: process.env.DISABLE_HMR !== 'true',
      // Disable file watching when DISABLE_HMR is true to save CPU during agent edits.
      // ignore .claude/ (permission auto-add) + db.json + scratchpad — กันหน้า reload ตัวเองรัวๆ ตอน dev (กดปุ่มไม่ติด)
      // ignore ไฟล์ฝั่ง server + config + ข้อมูล — Vite watch เฉพาะ frontend (src/) กันหน้า reload รัวๆ ตอน dev (กดปุ่มไม่ติด)
      watch: process.env.DISABLE_HMR === 'true' ? null : {
        ignored: [
          '**/.claude/**', '**/db.json', '**/Excel/**', '**/node_modules/**',
          '**/server.ts', '**/server-db.ts', '**/experimental-routes.ts', // ฝั่ง server (tsx รันแยก ไม่เกี่ยว frontend)
          '**/*.test.mts', '**/scratchpad/**',
        ],
      },
    },
  };
});
