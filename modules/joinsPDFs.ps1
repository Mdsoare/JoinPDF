# Carrega as DLLs local
Add-Type -Path (Resolve-Path ".\BouncyCastle.Cryptography.dll")
Add-Type -Path (Resolve-Path ".\itextsharp.dll")

$outputFile = "Resultado_Final.pdf"
$pdfs = Get-ChildItem *.pdf -Exclude $outputFile

# Inicia o documento
$doc = New-Object iTextSharp.text.Document
$stream = [System.IO.File]::Create((Join-Path $pwd $outputFile))
$copy = New-Object iTextSharp.text.pdf.PdfCopy($doc, $stream)
$doc.Open()

foreach ($file in $pdfs) {
    Write-Output "Processando: $($file.Name)"
    $reader = New-Object iTextSharp.text.pdf.PdfReader($file.FullName)
    for ($i = 1; $i -le $reader.NumberOfPages; $i++) {
        $page = $copy.GetImportedPage($reader, $i)
        $copy.AddPage($page)
    }
    $reader.Close()
}

$doc.Close()
$stream.Close()
Write-Output "Pronto! Arquivo $outputFile criado com sucesso." -ForegroundColor Green