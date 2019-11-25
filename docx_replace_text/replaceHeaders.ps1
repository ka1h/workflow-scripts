$documents_path = $PSScriptRoot

$word_app = New-Object -ComObject Word.Application

# This filter will find .doc as well as .docx documents
Get-ChildItem -Path $documents_path -Filter *.doc? | ForEach-Object {
       $replaceWith = "Lorem Ipsum dolor sit amet!"
       $replace = [Microsoft.Office.Interop.Word.WdReplace]::wdReplaceAll
       $findWrap = [Microsoft.Office.Interop.Word.WdFindWrap]::wdFindContinue

       $find = $header.Range.find
       $find.Execute($header.Range.Text, 
                     $false, #match case
                     $false, #match whole word
                     $false, #match wildcards
                     $false, #match soundslike
                     $false, #match all word forms
                     $true,  #forward
                     $findWrap, 
                     $null,      #format
                     $replaceWith,
                     $replace)
}

$word_app.Quit()