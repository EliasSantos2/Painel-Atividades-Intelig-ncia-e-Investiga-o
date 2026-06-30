# ============================================================
# Atualizar_Planilhas.ps1
# Atualiza todas as planilhas Excel da pasta de investigação
# Dispara automaticamente no logon quando há internet
# ============================================================

$LOG = "$PSScriptRoot\update_log.txt"
$PASTA = "C:\Users\risco\OneDrive\Núcleo de Inteligência\03 - INVESTIGAÇÕES\I-38 CASO RONDONÓPOLIS"

function Write-Log($msg) {
    $linha = "[$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')] $msg"
    Add-Content -Path $LOG -Value $linha -Encoding UTF8
    Write-Host $linha
}

# ── 1. Aguarda internet (máx. 2 min) ─────────────────────────────────────────
Write-Log "Iniciando verificação de conectividade…"
$tentativas = 0
while ($tentativas -lt 24) {
    if (Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet) { break }
    Start-Sleep -Seconds 5
    $tentativas++
}
if ($tentativas -ge 24) {
    Write-Log "AVISO: Sem internet após 2 min. Atualizando apenas arquivos locais."
}

# ── 2. Aguarda OneDrive sincronizar (30 s extras se havia internet) ───────────
if ($tentativas -lt 24) { Start-Sleep -Seconds 30 }

# ── 3. Coleta planilhas (exclui arquivos temporários ~$) ─────────────────────
$planilhas = Get-ChildItem -Path $PASTA -Recurse -Include "*.xlsx","*.xlsm","*.xls" -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notlike "~$*" } |
    Sort-Object FullName

Write-Log "Encontradas $($planilhas.Count) planilha(s)."

if ($planilhas.Count -eq 0) {
    Write-Log "Nenhuma planilha encontrada. Encerrando."
    exit
}

# ── 4. Abre Excel invisível e atualiza cada arquivo ──────────────────────────
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    foreach ($arquivo in $planilhas) {
        Write-Log "Abrindo: $($arquivo.Name)"
        $wb = $null
        try {
            $wb = $excel.Workbooks.Open($arquivo.FullName, 0, $false)
            # Aguarda abertura
            Start-Sleep -Seconds 2

            # Atualiza todas as conexões e tabelas dinâmicas
            $wb.RefreshAll()
            Start-Sleep -Seconds 3

            $wb.Save()
            Write-Log "OK: $($arquivo.Name)"
        } catch {
            Write-Log "ERRO em $($arquivo.Name): $_"
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
    Write-Log "ERRO GERAL: $_"
} finally {
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        $excel = $null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Log "Atualização concluída."
