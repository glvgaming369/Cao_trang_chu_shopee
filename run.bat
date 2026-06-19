@echo off
title Shopee Scraper Pro
echo Dang khoi chay ung dung Shopee Scraper...
python shopee_scraper_gui.py
if %errorlevel% neq 0 (
    echo.
    echo [LOI] Co loi xay ra khi chay ung dung!
    pause
)
