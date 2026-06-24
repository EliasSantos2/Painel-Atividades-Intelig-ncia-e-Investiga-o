# ============================================================
# CARREGA OS INFRATORES NO SUPABASE (upload automatico)
# ------------------------------------------------------------
# PRE-REQUISITO: a tabela 'infratores' precisa existir.
#   1) No Supabase -> SQL Editor -> rode o arquivo:
#      "15 - DASHBOARD\Planilha\CRIAR_TABELA_infratores.sql"
#
# USO (rode no PowerShell, na pasta do projeto):
#   .\carregar_infratores_supabase.ps1 -ServiceKey "COLE_AQUI_SUA_SERVICE_ROLE"
#
# A chave service_role esta em: Supabase -> Settings -> API -> 'service_role secret'
# ATENCAO: nao compartilhe essa chave nem a salve em arquivo versionado.
# ============================================================
param(
  [Parameter(Mandatory=$true)][string]$ServiceKey,
  [string]$Csv = "C:\Users\risco\OneDrive\Núcleo de Inteligência\15 - DASHBOARD\Planilha\infratores_supabase.csv",
  [switch]$LimparAntes
)

$ErrorActionPreference = 'Stop'
$url = 'https://rfwmklqzlfvllsrkvgob.supabase.co'

if(-not (Test-Path $Csv)){ Write-Host "CSV nao encontrado: $Csv" -ForegroundColor Red; exit 1 }

$headers = @{ apikey=$ServiceKey; Authorization="Bearer $ServiceKey"; 'Content-Type'='application/json' }

# 0) Testa se a tabela existe
try{
  Invoke-RestMethod -Uri "$url/rest/v1/infratores?select=id&limit=1" -Headers $headers -TimeoutSec 25 | Out-Null
}catch{
  $msg = $_.Exception.Message
  if("$msg" -match '404' -or "$msg" -match 'Not Found'){
    Write-Host "A tabela 'infratores' ainda NAO existe." -ForegroundColor Red
    Write-Host "Rode primeiro o SQL: 15 - DASHBOARD\Planilha\CRIAR_TABELA_infratores.sql no SQL Editor do Supabase." -ForegroundColor Yellow
    exit 1
  }
  Write-Host "Falha ao consultar (verifique a service_role key): $msg" -ForegroundColor Red
  exit 1
}

# 1) (Opcional) limpa a tabela antes de recarregar
if($LimparAntes){
  Write-Host "Limpando registros anteriores..."
  $h2 = $headers.Clone(); $h2['Prefer']='return=minimal'
  Invoke-RestMethod -Method Delete -Uri "$url/rest/v1/infratores?id=gt.0" -Headers $h2 | Out-Null
}

# 2) Le o CSV e monta payload (la/lo numericos ou null)
$rows = Import-Csv -Path $Csv
$payload = foreach($r in $rows){
  $la = if($r.la -and $r.la -ne ''){ [double]$r.la } else { $null }
  $lo = if($r.lo -and $r.lo -ne ''){ [double]$r.lo } else { $null }
  [pscustomobject]@{
    nome=$r.nome; status=$r.status; org=$r.org; cidade=$r.cidade; trecho=$r.trecho; dt=$r.dt
    la=$la; lo=$lo; mae=$r.mae; pai=$r.pai; cpf=$r.cpf; bo=$r.bo
  }
}
Write-Host "Registros a enviar: $($payload.Count)"

# 3) Envia em lotes
$hIns = $headers.Clone(); $hIns['Prefer']='return=minimal'
$batch = 500; $total = $payload.Count; $ok = 0
for($i=0; $i -lt $total; $i += $batch){
  $end = [Math]::Min($i+$batch-1, $total-1)
  $chunk = @($payload[$i..$end])
  $body = $chunk | ConvertTo-Json -Depth 4
  if($chunk.Count -eq 1){ $body = "[$body]" }
  Invoke-RestMethod -Method Post -Uri "$url/rest/v1/infratores" -Headers $hIns -Body ([System.Text.Encoding]::UTF8.GetBytes($body))
  $ok += $chunk.Count
  Write-Host ("  enviados {0}/{1}" -f $ok,$total)
}
Write-Host "Concluido! $ok registros carregados no Supabase." -ForegroundColor Green
