# Define os caminhos
$diretorio = $PSScriptRoot
if(!$diretorio) { $diretorio = $pwd } # Garante que pega a pasta atual
$outputFile = Join-Path $diretorio "ResultadoFinal.pdf"
$arquivos = Get-ChildItem -Path "$diretorio\*.pdf" -Exclude "ResultadoFinal.pdf" | Sort-Object Name

try {
    Write-Host "Iniciando o Word para processar os arquivos..." -ForegroundColor Cyan
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false  # O Word não vai aparecer na tela
    
    # Cria um documento novo que servirá de base
    $documentoFinal = $word.Documents.Add()
    $selecao = $word.Selection

    foreach ($arquivo in $arquivos) {
        Write-Host "Mesclando: $($arquivo.Name)"
        # Insere o conteúdo do PDF no documento do Word
        $selecao.InsertFile($arquivo.FullName)
        # Adiciona uma quebra de página para o próximo certificado não colar no anterior
        $selecao.InsertBreak(7) # 7 = wdPageBreak
    }

    # Salva o resultado final como PDF
    $documentoFinal.ExportAsFixedFormat($outputFile, 17) # 17 = Formato PDF
    
    Write-Host "`nSucesso! O arquivo foi gerado em: $outputFile" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao processar: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    # Fecha o Word obrigatoriamente para não travar o PC
    if ($documentoFinal) { $documentoFinal.Close($false) }
    if ($word) { $word.Quit() }
    
    # Remove o objeto da memória
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}