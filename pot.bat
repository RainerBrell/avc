@echo off
cls
rem 29. Juni 2025 
set Ord=addon\
set name=avc
set pot=avc.pot

Echo POT Datei erstellen:
scons pot 
Echo POT-Datei mit bestehenden PO-Dateien verschmelzen:
For /R %ord%locale\ %%f In (*.po) Do msgmerge -U --backup=none %%f %pot% 
pause 