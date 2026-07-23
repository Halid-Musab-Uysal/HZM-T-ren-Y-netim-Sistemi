@echo off
echo HZM Sistemi Kapatiliyor...
:: Python (Flask) sunucusunu sonlandırır
taskkill /f /im python.exe /t
:: Eğer Chrome pencerelerini kapatmak istersen aşağıdaki satırı aktif et:
:: taskkill /f /im chrome.exe /t
echo Sistem Durduruldu.
timeout /t 2
exit