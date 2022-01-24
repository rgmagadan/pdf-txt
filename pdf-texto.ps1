Write-Host "Trabajando..."
$contador = 0
$word = New-Object -comobject Word.Application
$Format = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatUnicodeText
foreach( $item in Get-ChildItem $args[0])
{
$doc = $word.Documents.Open($item.fullname)
$doc.SaveAs($item.fullname+".txt", $format)
$doc.Close()
$nombre = $item.fullname+".txt"
Get-Content $nombre | Add-Content file.txt -Encoding UTF8
Remove-Item $nombre
$contador++
}
$word.Quit()
Write-Host ("Listo. Se han convertido " + $contador.ToString() + " archivos.")