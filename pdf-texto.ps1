
$word = New-Object -comobject Word.Application
$Format = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatUnicodeText
foreach( $item in Get-ChildItem $args[0])
{
$doc = $word.Documents.Open($item.fullname)
$doc.SaveAs($item.fullname+".txt", $format)
$doc.Close()
}
$word.Quit()