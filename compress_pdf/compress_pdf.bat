@echo off
setlocal

for %%I in (*.pdf) do (
gswin64c -dNOPAUSE -dBATCH -sDEVICE=png16m -dTextAlphaBits=4 -r300 -sOutputFile="%%~nI_p%%02d.png" "%%~I"
)

magick convert *.png -quality 100 compressed.pdf

del *.png