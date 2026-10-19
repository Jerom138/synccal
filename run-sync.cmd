@echo off
rem Lancement de la synchronisation pour le planificateur de taches Windows.
rem Passe par un fichier .cmd plutot que par une ligne schtasks : les guillemets
rem et le && d une commande inline y sont une source classique d erreurs.
cd /d "%~dp0"
if not exist "logs" mkdir "logs"
echo ---------- %DATE% %TIME% ---------->> "logs\schedule.log"
python sync.py --apply >> "logs\schedule.log" 2>&1
exit /b %ERRORLEVEL%
