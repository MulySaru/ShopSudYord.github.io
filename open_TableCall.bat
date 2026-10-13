@echo off
chcp 65001 >nul
echo กำลังเปิดระบบ ร้านก๋วยเตี๋ยวเรือ รงค์เรือ ในเว็บเบราว์เซอร์ของคุณ...
start "" "%~dp0index.html"
exit
