$word = New-Object -ComObject Word.Application
$word.Visible = $false
$files = Get-ChildItem -Path (Get-Location) -Filter *.docx

foreach ($file in $files) {
    $doc = $word.Documents.Open($file.FullName)
    $pdfPath = $file.FullName -replace "\.docx$", ".pdf"
    $doc.SaveAs([ref] $pdfPath, [ref] 17)  # 17 es formato PDF
    $doc.Close()
}

$word.Quit()
