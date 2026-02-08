@echo off
chcp 65001 >nul

REM Переходим в папку проекта
cd /d "C:\Users\l_trofimov\Desktop\Эпид мониторинг"

REM Запуск программы
py microbio_app.py

pause
