@echo off
chcp 65001 >nul
REM Сборка "Агрегатор таблиц.exe" одной командой.
REM Требуется: .venv с установленными зависимостями (requirements.txt).
.venv\Scripts\python.exe -m PyInstaller --onefile --console --clean --name "Агрегатор таблиц" --collect-all rich main.py
echo.
echo Готовый файл: dist\Агрегатор таблиц.exe
pause
