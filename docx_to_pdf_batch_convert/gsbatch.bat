@echo off
setlocal

Powershell.exe -executionpolicy remotesigned -File .\doctopdf.ps1

for %%I in (*.pdf) do (
    gswin64c -dNOPAUSE -dBATCH -sDEVICE=jpeg -dTextAlphaBits=4 -r300 -dFirstPage=1 -dLastPage=1 -sOutputFile="%%~nI_p%%02d.jpg" "%%~I"
)

del *.pdf

if not exist "Vorschaubilder" mkdir Vorschaubilder

move *.jpg Vorschaubilder