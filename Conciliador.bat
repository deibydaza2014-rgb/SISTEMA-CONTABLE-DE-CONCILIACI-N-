@echo off
title Sistema Contable de Conciliacion 3 Vias - ALLTECH y PARMOR
echo ==========================================================
echo  SISTEMA DE CONCILIACION BANCARIA 3 VIAS EN EJECUCION
echo  Empresas: ALLTECH SOLUTIONS S.A.S. y PARMOR SAS
echo  Autor: Deiby Andres Vasquez Daza
echo ==========================================================
echo.
echo 1. Abriendo navegador web en http://localhost:8080 ...
cd /d "%~dp0"

:: Abrir navegador
start http://localhost:8080

echo 2. Manteniendo el motor contable encendido...
echo (NO cierres esta ventana mientras uses el programa)
echo.

:: Ejecutar servidor en primer plano para que NUNCA se cierre solo
python conciliador_alltech.py --server --port 8080

pause
