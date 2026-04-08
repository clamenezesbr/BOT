@echo off
echo Instalando dependencias...
pip install opencv-python numpy pyautogui mss keyboard pywin32 pyinstaller

echo.
echo Limpando build antigo...
rmdir /s /q build
rmdir /s /q dist
del *.spec

echo Gerando novo executavel...
pyinstaller --onefile fishing_bot.py

echo.
echo Build finalizado! Veja em /dist
pause
