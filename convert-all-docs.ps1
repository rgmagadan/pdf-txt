$contador = 0
$word = New-Object -comobject Word.Application
$Format = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatUnicodeText
Write-Host "Trabajando..."
foreach( $item in Get-ChildItem $args[0])
{
    if ($item.fullname -match '.+\.(doc|rtf|docx|pdf)$'){
        $doc = $word.Documents.Open($item.fullname)
        $doc.SaveAs($item.fullname+".txt", $format)
        $doc.Close()
        $texto = $item.fullname+".txt"
        Add-Content -Value "EOF" -Path $texto
        Get-Content $texto | Out-File -Append file.txt -Encoding UTF8 -ErrorAction Stop
        $contador++
Remove-Item $texto
    }
}
Write-Host ("Listo. Se han convertido " + $contador.ToString() + " archivos.")
$word.Quit()