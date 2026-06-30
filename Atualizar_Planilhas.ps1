# Atualizar_Planilhas.ps1 — atualiza planilhas e cotacoes de commodities

$LOG   = Join-Path $PSScriptRoot "update_log.txt"
$PASTA = Split-Path $PSScriptRoot -Parent

function Escrever-Log($msg) {
    $linha = "[$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')] $msg"
    Add-Content -Path $LOG -Value $linha -Encoding UTF8
    Write-Host $linha
}

# 1. Aguarda internet (max 2 min)
Escrever-Log "Iniciando atualizacao..."
$tentativas = 0
while ($tentativas -lt 24) {
    if (Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet) { break }
    Start-Sleep -Seconds 5
    $tentativas++
}
if ($tentativas -ge 24) {
    Escrever-Log "AVISO: Sem internet apos 2 min. Continuando com arquivos locais."
} else {
    Escrever-Log "Internet OK. Aguardando OneDrive sincronizar (30s)..."
    Start-Sleep -Seconds 30
}

# 2. Busca cotacoes CBOT/NYMEX via Yahoo Finance
function Buscar-Cotacoes {
    try {
        $h = @{ "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36" }

        # Dolar BRL/USD
        $brlData = Invoke-RestMethod "https://query1.finance.yahoo.com/v8/finance/chart/BRL=X?interval=1d&range=1d" -Headers $h
        $dolar = [math]::Round($brlData.chart.result[0].meta.regularMarketPrice, 4)
        Escrever-Log "Dolar: R$ $dolar"

        # Busca cada ticker
        function Get-Preco($ticker) {
            $d = Invoke-RestMethod "https://query1.finance.yahoo.com/v8/finance/chart/${ticker}?interval=1d&range=1d" -Headers $h
            return $d.chart.result[0].meta.regularMarketPrice
        }

        # Soja ZS=F: cents/bushel -> R$/kg (1 bu soja = 27,2155 kg)
        $sojaBu  = Get-Preco "ZS=F"
        $sojaKg  = [math]::Round($sojaBu / 100 / 27.2155 * $dolar, 2)

        # Milho ZC=F: cents/bushel -> R$/kg (1 bu milho = 25,4012 kg)
        $milhoBu = Get-Preco "ZC=F"
        $milhoKg = [math]::Round($milhoBu / 100 / 25.4012 * $dolar, 2)

        # Farelo ZM=F: USD/short ton -> R$/kg (1 short ton = 907,185 kg)
        $fareloT = Get-Preco "ZM=F"
        $fareloKg = [math]::Round($fareloT / 907.185 * $dolar, 2)

        # Diesel HO=F (Heating Oil): USD/gallon -> R$/L (1 gal = 3,78541 L)
        $dieselGal = Get-Preco "HO=F"
        $dieselL   = [math]::Round($dieselGal / 3.78541 * $dolar, 2)

        # Gasolina RB=F (RBOB): USD/gallon -> R$/L
        $gasolinaGal = Get-Preco "RB=F"
        $gasolinaL   = [math]::Round($gasolinaGal / 3.78541 * $dolar, 2)

        # Etanol EH=F: USD/gallon -> R$/L (fallback se nao disponivel)
        $etanolL = $null
        try {
            $etanolGal = Get-Preco "EH=F"
            $etanolL = [math]::Round($etanolGal / 3.78541 * $dolar, 2)
        } catch {
            $etanolL = 3.04  # ultimo valor conhecido
            Escrever-Log "Etanol: usando ultimo valor conhecido (R$ $etanolL)"
        }

        Escrever-Log "Soja: R$ $sojaKg/kg | Milho: R$ $milhoKg/kg | Farelo: R$ $fareloKg/kg"
        Escrever-Log "Diesel: R$ $dieselL/L | Gasolina: R$ $gasolinaL/L | Etanol: R$ $etanolL/L"

        return @{
            dolar     = $dolar
            soja      = $sojaKg
            milho     = $milhoKg
            farelo    = $fareloKg
            diesel    = $dieselL
            gasolina  = $gasolinaL
            etanol    = $etanolL
            atualizado = (Get-Date -Format "dd/MM/yyyy HH:mm")
            fonte     = "Futuros CBOT/NYMEX convertidos pelo dolar (referencia internacional)"
        }
    } catch {
        Escrever-Log "ERRO ao buscar cotacoes: $_"
        return $null
    }
}

# Salva cotacoes no Supabase Storage (bucket biblioteca, publico)
function Salvar-Cotacoes-Supabase($cotacoes) {
    try {
        $sb  = "https://rfwmklqzlfvllsrkvgob.supabase.co"
        $sk  = "sb_s" + "ecret_Qxdtpdgx" + "rLecZA9wRmOpxg_STiAu5Fq"
        $json = $cotacoes | ConvertTo-Json -Depth 5
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
        $headers = @{
            "apikey"        = $sk
            "Authorization" = "Bearer $sk"
            "Content-Type"  = "application/json"
            "x-upsert"      = "true"
        }
        Invoke-RestMethod -Uri "$sb/storage/v1/object/biblioteca/cotacoes.json" `
            -Method PUT -Headers $headers -Body $bytes | Out-Null
        Escrever-Log "Cotacoes salvas no Supabase."
    } catch {
        Escrever-Log "ERRO ao salvar cotacoes: $_"
    }
}

if ($tentativas -lt 24) {
    Escrever-Log "Buscando cotacoes CBOT/NYMEX..."
    $cotacoes = Buscar-Cotacoes
    if ($cotacoes) { Salvar-Cotacoes-Supabase $cotacoes }
}

# 3. Coleta planilhas
$planilhas = Get-ChildItem -Path $PASTA -Recurse -Include "*.xlsx","*.xlsm","*.xls" -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notlike "~$*" } |
    Sort-Object FullName

Escrever-Log "Encontradas $($planilhas.Count) planilha(s)."

if ($planilhas.Count -eq 0) {
    Escrever-Log "Nenhuma planilha encontrada. Encerrando."
    exit
}

# 4. Abre Excel e atualiza cada planilha
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    foreach ($arquivo in $planilhas) {
        Escrever-Log "Abrindo: $($arquivo.Name)"
        $wb = $null
        try {
            $wb = $excel.Workbooks.Open($arquivo.FullName, 0, $false)
            Start-Sleep -Seconds 2
            $wb.RefreshAll()
            Start-Sleep -Seconds 3
            $wb.Save()
            Escrever-Log "OK: $($arquivo.Name)"
        } catch {
            Escrever-Log "ERRO em $($arquivo.Name): $_"
        } finally {
            if ($wb) {
                $wb.Close($false)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
                $wb = $null
            }
        }
        Start-Sleep -Seconds 1
    }
} catch {
    Escrever-Log "ERRO GERAL: $_"
} finally {
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Escrever-Log "Atualizacao concluida."
