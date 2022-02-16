param([string]$path)
$word = New-Object -comobject Word.Application
$Format = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatUnicodeText
    $doc = $word.Documents.Open($path)
    $doc.SaveAs($path+".txt", $format)
    $doc.Close()
    $word.Quit()