@echo off
title FRIOPACKING — Servidor OC Mailer + SharePoint
echo.
echo  ============================================================
echo   FRIOPACKING S.A. — Servidor de Correos OC + SharePoint
echo   Mantener esta ventana abierta al usar el dashboard
echo  ============================================================
echo.

cd /d "%~dp0"

echo  Verificando dependencias...
python -m pip install --quiet flask flask-cors pywin32 2>nul
echo  Dependencias OK.
echo.

python oc_mailer.py

echo.
echo  El servidor se detuvo. Presiona cualquier tecla para cerrar.
pause >nul
