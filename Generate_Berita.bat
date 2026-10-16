@echo off
echo =========================================
echo Menjalankan Server Lokal Agora Vada...
echo Jangan tutup jendela CMD ini selama ngedit!
echo =========================================

:: Buka browser otomatis ke alamat lokal
start http://localhost:3000

:: Paksa Next.js menggunakan Webpack (bypass Turbopack)
call npx next dev --webpack

pause