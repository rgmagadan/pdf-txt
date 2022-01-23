$ruta = Read-Host "Introduce la ruta completa del archivo:"
$word = New-Object -comobject Word.Application
$doc = $word.Documents.Open($ruta)
$Format = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatUnicodeText
$doc.SaveAs($ruta+".txt", $format)
$doc.Close()
$word.Quit()