@echo off

REM Homepath = C:\Users\beep
REM strHomePath&"\Desktop\webcam must exist
REM repeat 500 times
REM vbs includes 30 seconds pause starting 10 am.
REM you will save 250 minutes = little bit more than 2 hours


REM {Variable} IN (Startzahl, Schrittweite, Endzahl) DO (Aktion)

FOR /L %%N IN (1,1,50) DO (
cscript.exe %homepath%\Desktop\webcam\webcam_download.vbs
)

echo %%N Wiederholungen
pause
