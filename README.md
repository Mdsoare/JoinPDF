# PDF Merge Tool (PowerShell + iTextSharp)

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://docs.microsoft.com/en-us/powershell/)
[![Security Compliance](https://img.shields.io/badge/Security-Local%20Only%20%2F%20Zero%20Trust-green.svg)](#segurança-e-compliance-devsecops)
[![License](https://img.shields.io/badge/License-MIT-lightgrey.svg)](LICENSE)

Utilitário em PowerShell para fusão (*merge*) automatizada de arquivos PDF localmente em ambientes Windows, com foco em **privacidade de dados (LGPD)**, **desempenho** e **segurança da informação**.

---

## 📋 Sumário
- [Visão Geral](#visão-geral)
- [Por que usar esta solução? (Perspectiva DevSecOps)](#por-que-usar-esta-solução-perspectiva-devsecops)
- [Pré-requisitos](#pré-requisitos)
- [Estrutura do Repositório](#estrutura-do-repositório)
- [Guia de Uso Passo a Passo](#guia-de-uso-passo-a-passo)
- [Funcionamento do Código](#funcionamento-do-código)
- [Segurança e Compliance (DevSecOps)](#segurança-e-compliance-devsecops)
- [Solução de Problemas (Troubleshooting)](#solução-de-problemas-troubleshooting)
- [Boas Práticas e Recomendações](#boas-práticas-e-recomendações)

---

## 🔍 Visão Geral

Em ambientes corporativos, a unificação de documentos PDF que contêm informações sensíveis (PII, relatórios financeiros, logs de auditoria ou documentos jurídicos) muitas vezes leva usuários a recorrerem a ferramentas web gratuitas. Essa prática representa um alto risco de vazamento de dados (*Data Exfiltration*) e não conformidade com a **LGPD (Lei Geral de Proteção de Dados)**.

Esta ferramenta resolve esse problema permitindo a manipulação de PDFs **100% offline**, utilizando bibliotecas consolidadas (`iTextSharp` e `BouncyCastle`) diretamente via Scripting em PowerShell.

---

## 🔒 Por que usar esta solução? (Perspectiva DevSecOps)

- **Air-Gapped & Local Processing:** Nenhum dado sai da máquina local. Não há chamadas HTTP/HTTPS ou conexões de rede ativas.
- **Conformidade com LGPD:** Evita a transferência não autorizada de dados pessoais para servidores de terceiros ou cloud não homologada.
- **Sem Dependência de Instalação Administrativa:** Não exige permissões de Administrador local (`Run as Admin`) para rodar, desde que a execução de scripts locais seja autorizada por política.
- **Zero Custos de Licenciamento SaaS:** Elimina a necessidade de licenças pagas de softwares de terceiros para tarefas simples de manipulação de PDF.

---

## 🛠️ Pré-requisitos

1. **Sistema Operacional:** Windows 10/11 ou Windows Server 2016+
2. **PowerShell:** Versão 5.1 ou superior (já nativo no Windows).
3. **Bibliotecas .NET (incluídas no diretório):**
   - `BouncyCastle.Cryptography.dll`
   - `itextsharp.dll`

---

## 📁 Estrutura do Repositório

```text
.
├── .gitignore                     # Aquivo de configuração do git
├── BouncyCastle.Cryptography.dll  # Biblioteca de criptografia / dependência iTextSharp
├── itextsharp.dll                 # Motor de manipulação e montagem de PDFs (.NET)
├── joinPDFs_alternative.ps1       # Script de automação PowerShell alternativo
├── joinPDFs.ps1                   # Script de automação PowerShell
└── README.md                      # Documentação do projeto
```

---

## 🚀 Guia de Uso Passo a Passo

### 1º Passo: Copiar os arquivos necessários
Copie os **três arquivos** abaixo para a pasta onde estão armazenados os arquivos PDF que você deseja juntar:
- `BouncyCastle.Cryptography.dll`
- `itextsharp.dll`
- `joinPDFs.ps1`

---

### 2º Passo: Abrir o PowerShell no diretório
1. Navegue até a pasta onde estão os arquivos.
2. Mantenha a tecla **`Shift`** pressionada.
3. Clique com o **botão direito do mouse** em um espaço em branco da pasta.
4. Selecione a opção **"Abrir janela do PowerShell aqui"** (ou *"Abrir no Terminal"* no Windows 11).

---

### 3º Passo: Executar o script
Execute o comando abaixo no PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\joinPDFs.ps1
```

> **Nota:** A flag `-ExecutionPolicy Bypass` garante a execução do script no escopo da sessão atual sem alterar a política de execução global do sistema (*Execution Policy*).

---

### 4º Passo: Localizar o resultado
Ao finalizar a execução, o arquivo unificado será gerado no mesmo diretório com o nome:
- **`Resultado_Final.pdf`**

O script ignorará automaticamente o arquivo `Resultado_Final.pdf` em execuções subsequentes na mesma pasta, evitando loops ou reprocessamento do arquivo consolidado.

---

## 💻 Funcionamento do Código

O script `joinPDFs.ps1` utiliza o ecossistema .NET nativo e as DLLs do iTextSharp para realizar a mesclagem diretamente em memória:

```powershell
# 1. Carregamento dinamico das DLLs locais .NET
Add-Type -Path (Resolve-Path ".\BouncyCastle.Cryptography.dll")
Add-Type -Path (Resolve-Path ".\itextsharp.dll")

# 2. Definição do arquivo de saída e listagem dos PDFs (excluindo o arquivo final)
$outputFile = "Resultado_Final.pdf"
$pdfs = Get-ChildItem *.pdf -Exclude $outputFile

# 3. Inicialização dos objetos iTextSharp e Stream de escrita
$doc = New-Object iTextSharp.text.Document
$stream = [System.IO.File]::Create((Join-Path $pwd $outputFile))
$copy = New-Object iTextSharp.text.pdf.PdfCopy($doc, $stream)
$doc.Open()

# 4. Iteração página a página entre todos os PDFs encontrados
foreach ($file in $pdfs) {
    Write-Host "Processando: $($file.Name)"
    $reader = New-Object iTextSharp.text.pdf.PdfReader($file.FullName)
    for ($i = 1; $i -le $reader.NumberOfPages; $i++) {
        $page = $copy.GetImportedPage($reader, $i)
        $copy.AddPage($page)
    }
    $reader.Close()
}

# 5. Encerramento dos handlers de arquivo e desalocação de memória
$doc.Close()
$stream.Close()
Write-Host "Pronto! Arquivo $outputFile criado com sucesso." -ForegroundColor Green
```

---

## 🛡️ Segurança e Compliance (DevSecOps)

### Análise de Segurança dos Binários (DLLs)
Para garantir a integridade da cadeia de suprimentos de software (*Software Supply Chain Security*), valide as assinaturas/hashes das DLLs fornecidas no repositório antes da primeira execução:

```powershell
# Verificação de Hash SHA256 no PowerShell
Get-FileHash .\itextsharp.dll, .\BouncyCastle.Cryptography.dll -Algorithm SHA256 | Format-List
```

### Controles de Segurança Implementados
- **Isolamento de Processo:** Sem tráfego de rede ou chamadas de API externas.
- **Sanitização de Escopo de Execução:** O uso do `-ExecutionPolicy Bypass` limita-se ao processo filho e não afeta a postura global de execução do repositório/SO.
- **Tratamento de Arquivo de Saída:** O filtro `-Exclude $outputFile` previne corrupção de arquivo por auto-ingestão ou substituição acidental durante a leitura.

---

## ⚠️ Solução de Problemas (Troubleshooting)

| Erro / Problema | Causa Provável | Solução |
| :--- | :--- | :--- |
| `Add-Type : Could not load file or assembly...` | DLLs bloqueadas pelo Windows após download ou caminho incorreto. | Clique com o botão direito nas DLLs > Propriedades > Marque **"Desbloquear"** (Unblock) > Aplicar. |
| `O arquivo já está sendo usado por outro processo` | O PDF resultante ou um dos arquivos de entrada está aberto em um leitor de PDF. | Feche o leitor de PDF (Acrobat, Edge, Chrome, etc.) e execute novamente. |
| `A execução de scripts foi desabilitada neste sistema` | Restrição de política local do PowerShell. | Certifique-se de incluir a flag `-ExecutionPolicy Bypass` ao chamar o script. |
| `Arquivo corrompido ao abrir` | Algum PDF de origem está protegido por senha ou corrompido. | Remova a proteção por senha do PDF de origem antes de rodar a mesclagem. |

---

## 💡 Boas Práticas e Recomendações

1. **Ordem de Fusão:** O script mescla os arquivos com base na **ordem alfabética dos nomes dos arquivos**. Para garantir uma ordem específica, renomeie os arquivos com prefixos numéricos (ex: `01_Introducao.pdf`, `02_Relatorio.pdf`, `03_Anexos.pdf`).
2. **Ambiente Corporativo:** Em pipelines CI/CD ou automações em lote, recomenda-se armazenar as DLLs em um repositório centralizado de artefatos (Nexus, JFrog Artifactory ou Azure Artifacts) e realizar o download via hash verificado.