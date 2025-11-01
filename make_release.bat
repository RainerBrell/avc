echo off
cls
setlocal enabledelayedexpansion
rem 2025.11.01: macht ein Comit, danach ein tag und pusht es nach github 

rem Eingabe der Commit-Nachricht
set /p commitMsg=Gib die Commit-Nachricht ein: 

rem Eingabe der Versionsnummer (Tag)
set /p version=Gib die Versionsnummer ein (z.B. 2025.08.16s): 

rem Eingabe des Release-Titels
set /p releaseTitle=Gib den Release-Titel ein: 

rem Eingabe der Release-Beschreibung
set /p releaseNotes=Gib die Release-Beschreibung ein: 

cls
echo Commitnachricht    : %commitMsg%
echo Versionsnummer     : %version%
echo Releasetitel       : %releaseTitle%
echo Releasebeschreibung: %releaseNotes%
Echo Wichtig: Das fertige Addon muss erstellt worden sein und die Versionnummer gleich dem Tag, also evtl. mit Buchstabe.
pause

echo.
echo === Änderungen werden committet ===
git add .
git commit -m "%commitMsg%"
pause 
echo.
echo === Tag wird erstellt ===
git tag -a %version% -m "%releaseTitle%"
pause 

echo.
echo === Änderungen und Tag werden gepusht ===
git push origin master
git push origin %version%
pause 
echo.
echo === Release wird auf GitHub erstellt ===
gh release create %version% "avc-%version%.nvda-addon" --title "%releaseTitle%" --notes "%releaseNotes%"
pause 
echo.
echo Fertig: %version%
pause