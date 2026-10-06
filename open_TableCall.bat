@echo off
chcp 65001 >nul
echo กำลังเปิดระบบ ร้านอาหารทดสอบ POS ในเว็บเบราว์เซอร์ของคุณ...
start "" "%~dp0index.html"
exit
