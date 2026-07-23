@echo off
:: Flask uygulamasını arka planda başlatır
start /b python app.py
:: Sunucunun açılması için 2 saniye bekle
timeout /t 2 /nobreak >nul
:: Kontrol Paneli ve Ekran pencerelerini Chrome ile (app modunda) açar
start chrome http://127.0.0.1:5000/panel
start chrome http://127.0.0.1:5000/ekran
exit