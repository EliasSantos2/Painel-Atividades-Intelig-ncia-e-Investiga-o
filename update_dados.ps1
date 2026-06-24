# ============================================================
# ATUALIZAÇÃO AUTOMÁTICA — Dashboard Indicadores da Inteligência
# Executa diariamente às 06:00 via Agendador de Tarefas Windows
# ============================================================
param([switch]$Silent)

$ErrorActionPreference = 'Stop'
$log = "$PSScriptRoot\update_log.txt"
function Log($msg) { $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'; "$ts  $msg" | Tee-Object -FilePath $log -Append | Write-Host }

Log "=== INICIO DA ATUALIZACAO ==="

Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Web

$boletinsPath  = "C:\Users\risco\OneDrive\Núcleo de Inteligência\15 - DASHBOARD\Planilha\Controle de Boletins (2).xlsx"
$processosPath = "C:\Users\risco\OneDrive\Núcleo de Inteligência\15 - DASHBOARD\Planilha\Controle geral de processos.xlsx"
$circPath      = "C:\Users\risco\OneDrive\Núcleo de Inteligência\15 - DASHBOARD\Planilha\Indicadores de Circulacao.xlsx"
$outDir        = $PSScriptRoot

# ── COPIA ARQUIVOS (evita lock do Excel) ────────────────────
$tmpBol  = "$env:TEMP\bol_update.xlsx"
$tmpProc = "$env:TEMP\proc_update.xlsx"
Copy-Item $boletinsPath  $tmpBol  -Force
Copy-Item $processosPath $tmpProc -Force
Log "Arquivos copiados para temp"

# ── HELPERS ─────────────────────────────────────────────────
function ReadSS($zip) {
    $e = $zip.Entries | Where-Object { $_.FullName -eq 'xl/sharedStrings.xml' }
    $s = $e.Open(); $r = New-Object System.IO.StreamReader($s); $x = $r.ReadToEnd(); $r.Close()
    [regex]::Matches($x, '<t[^>]*>([^<]*)</t>') | ForEach-Object { [System.Web.HttpUtility]::HtmlDecode($_.Groups[1].Value) }
}
function ParseSheet($zip, $path, $strings) {
    $sh=$zip.Entries|Where-Object{$_.FullName -eq $path}
    $ss=$sh.Open(); $sr=New-Object System.IO.StreamReader($ss); $xml=$sr.ReadToEnd(); $sr.Close()
    $rows=[regex]::Matches($xml,'<row[^>]*r="(\d+)"[^>]*>(.*?)</row>'); $result=@()
    foreach ($row in $rows) {
        $rn=[int]$row.Groups[1].Value; if($rn -eq 1){continue}
        $cells=[regex]::Matches($row.Groups[2].Value,'<c r="([A-Z]+)\d+"(?:[^>]*t="([^"]*)")?[^>]*>(?:.*?<v>([^<]*)</v>)?')
        $rec=@{}
        foreach($c in $cells){$col=$c.Groups[1].Value;$t=$c.Groups[2].Value;$v=$c.Groups[3].Value
            if($t -eq 's' -and $v -ne ''){try{$v=$strings[[int]$v]}catch{$v=''}}; $rec[$col]=$v}
        if($rec.Count -gt 0){$result+=$rec}
    }
    $result
}

# ── LÊ PROCESSOS ─────────────────────────────────────────────
$zip2=[System.IO.Compression.ZipFile]::OpenRead($tmpProc)
$ss2=ReadSS $zip2
$procs2=ParseSheet $zip2 'xl/worksheets/sheet2.xml' $ss2
$zip2.Dispose()
Log "Processos lidos: $($procs2.Count)"

$boLookup=@{}
foreach($p in $procs2){
    $caso=$p['E']
    if($caso -match 'BO\s+([A-Z0-9\-\/]+)'){
        $boNum=$matches[1].Trim()
        $dpRaw=if($p['G']){($p['G'] -split "`n")[0].Trim()}else{''}
        $andRaw=if($p['J']){($p['J'] -split "`n"|Where-Object{$_.Trim() -ne ''}|Select-Object -First 1).Trim()}else{''}
        if($andRaw.Length -gt 100){$andRaw=$andRaw.Substring(0,100)+'...'}
        $sigRaw=$p['M']; $sigCat='Sem Info'
        if($sigRaw -and $sigRaw.Trim() -ne ''){
            if($sigRaw -match 'não foi solicitada|não solicitada|não houve|não foi apreendido|não foram apreendidos'){$sigCat='Não'}
            elseif($sigRaw -match 'quebra de sigilo|apreendido|deferindo|requerendo|solicitada'){$sigCat='Sim'}
            else{$sigCat='Não'}
        }
        if(-not $boLookup.ContainsKey($boNum)){
            $boLookup[$boNum]=[ordered]@{proc=if($p['D']){$p['D']}else{''}; num=if($p['F']){$p['F']}else{''}
                dp=$dpRaw; vara=if($p['I']){($p['I'] -split "`n")[0].Trim()}else{''}; and=$andRaw; sig=$sigCat; ano=if($p['A']){$p['A']}else{''}}
        }
    }
}
Log "BOs com processo: $($boLookup.Count)"

# ── LÊ BOLETINS (Excel COM — array bulk read) ─────────────────────────────────
$xlApp = New-Object -ComObject Excel.Application
$xlApp.Visible = $false; $xlApp.DisplayAlerts = $false
$wbBol = $xlApp.Workbooks.Open($tmpBol)
$wsBol = $wbBol.Sheets.Item(1)

# Ler coluna T inteira para encontrar as linhas com coordenadas (uma chamada COM)
$colTArr = $wsBol.Range("T1:T30000").Value2  # array 1D ou 2D
$rowsWithCoord = New-Object System.Collections.Generic.List[int]
for($i=2; $i -le 30000; $i++){
    $v = if($colTArr -is [System.Array]){ $colTArr[$i,1] }else{ $null }
    if($v -ne $null -and "$v".Trim() -ne ''){ $rowsWithCoord.Add($i) }
}
$lastRow = if($rowsWithCoord.Count -gt 0){ $rowsWithCoord[$rowsWithCoord.Count-1] }else{ 2 }
$lastCol = 38
Log "Planilha: $($rowsWithCoord.Count) linhas com T preenchido, última = $lastRow"

# Leitura em bulk de todas as colunas necessárias até a última linha com coord
$bulkRange = $wsBol.Range($wsBol.Cells.Item(1,1), $wsBol.Cells.Item($lastRow, $lastCol))
$data = $bulkRange.Value2  # 2D array [1..lastRow, 1..lastCol]

$wbBol.Close($false); $xlApp.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xlApp) | Out-Null
Log "Array lido: $($data.GetLength(0)) x $($data.GetLength(1)) (linhas x colunas)"

# Mapear índices de colunas (1-based)
$colB=2;$colC=3;$colD=4;$colE=5;$colF=6;$colG=7;$colH=8;$colI=9;$colJ=10
$colK=11;$colL=12;$colM=13;$colN=14;$colO=15;$colP=16;$colQ=17;$colR=18;$colS=19
$colT=20;$colU=21;$colV=22;$colW=23;$colY=25;$colZ=26;$colAA=27;$colAB=28
$colAC=29;$colAD=30;$colAE=31;$colAF=32;$colAG=33;$colAH=34;$colAI=35;$colAJ=36;$colAK=37;$colAL=38

function CvTime($arr,$r,$c){
    $v = Cv $arr $r $c
    if($v -eq ''){ return '' }
    try{ $tot=[double]$v*24; $h=[math]::Floor($tot); $mi=[math]::Floor(($tot-$h)*60); return "{0:D2}:{1:D2}" -f $h,$mi }catch{ return $v }
}

function Cv($arr,$r,$c){ if($c -le $arr.GetUpperBound(1)){ $v=$arr[$r,$c]; if($v -ne $null){"$v".Trim()}else{''} }else{''} }

$coords=New-Object System.Collections.Generic.List[object]

for($r=2; $r -le $lastRow; $r++){
    $coord = Cv $data $r $colT
    if(-not $coord){ continue }
    if(-not ($coord -match '^\s*-?\d+[\.,]\d+[\s,]+\s*-?\d+[\.,]\d+\s*$')){ continue }
    $parts = ($coord -split '[,\s]+' | Where-Object{$_ -ne ''})
    try{
        $lat=[double]$parts[0]; $lng=[double]$parts[1]
        if($lat -gt -35 -and $lat -lt 5 -and $lng -gt -75 -and $lng -lt -30){
            $dataFmt=''; $horaFmt=''
            $rawI = Cv $data $r $colI
            if($rawI){try{$d=[DateTime]::FromOADate([double]$rawI); $dataFmt=$d.ToString('dd/MM/yyyy')}catch{}}
            $rawJ = Cv $data $r $colJ
            if($rawJ){try{$tot=[double]$rawJ*24; $h=[math]::Floor($tot); $m=[math]::Floor(($tot-$h)*60); $horaFmt="{0:D2}:{1:D2}" -f $h,$m}catch{}}
            $boVal = Cv $data $r $colAC
            if($boVal -eq '0'){ $boVal='' }
            $inqData=$null
            if($boVal -and $boLookup.ContainsKey($boVal)){$inqData=$boLookup[$boVal]}
            $rec=[ordered]@{la=$lat;lo=$lng
                bo  =$boVal
                cl  =Cv $data $r $colC
                cr  =Cv $data $r $colD
                art =Cv $data $r $colE
                tp  =Cv $data $r $colV
                sb  =Cv $data $r $colW
                nat =Cv $data $r $colU
                esp =Cv $data $r $colZ
                dt  =$dataFmt; hr=$horaFmt
                mes =Cv $data $r $colH
                cid =Cv $data $r $colK
                uf  =Cv $data $r $colL
                circ=Cv $data $r $colF
                km  =Cv $data $r $colO
                comp=Cv $data $r $colP
                emp =Cv $data $r $colQ
                aut =Cv $data $r $colAD
                ras =Cv $data $r $colB
                perf=Cv $data $r $colG
                thp =Cv $data $r $colAH
                tre =Cv $data $r $colM
                sbn =Cv $data $r $colN
                maq =Cv $data $r $colR
                re  =Cv $data $r $colS
                qtd =Cv $data $r $colY
                cpf =Cv $data $r $colAA
                prot=Cv $data $r $colAB
                recv=Cv $data $r $colAE
                hi  =CvTime $data $r $colAF
                hf  =CvTime $data $r $colAG
                cgr =Cv $data $r $colAI
                grau=Cv $data $r $colAJ
                tipos=Cv $data $r $colAK
                val =Cv $data $r $colAL
                inq =$inqData}
            $coords.Add($rec)
        }
    } catch{}
}
Log "Coordenadas extraídas: $($coords.Count)"

# ── SALVA JSON ────────────────────────────────────────────────
$json = $coords | ConvertTo-Json -Compress -Depth 4

# Atualiza Mapa_Ocorrencias.html (const DATA)
$mapHtml = Get-Content "$outDir\Mapa_Ocorrencias.html" -Raw -Encoding utf8
$mapHtml = [regex]::Replace($mapHtml, 'const DATA = \[.*?\];', "const DATA = $json;", [System.Text.RegularExpressions.RegexOptions]::Singleline)
$mapHtml | Out-File "$outDir\Mapa_Ocorrencias.html" -Encoding utf8
Log "Mapa_Ocorrencias.html atualizado ($($coords.Count) pontos GPS)"

# Atualiza Indicadores_Inteligencia.html (const COORDS)
$indHtml = Get-Content "$outDir\Indicadores_Inteligencia.html" -Raw -Encoding utf8
$indHtml = [regex]::Replace($indHtml, 'const COORDS = \[.*?\];', "const COORDS = $json;", [System.Text.RegularExpressions.RegexOptions]::Singleline)

# Monta PROCS
$procsArr = @()
foreach ($p in $procs2) {
    $caso = $p['E']
    if ($caso -match 'BO\s+([A-Z0-9\-\/]+)') {
        $dpRaw  = if ($p['G']) { ($p['G'] -split "`n")[0].Trim() } else { '' }
        $andRaw = if ($p['J']) { ($p['J'] -split "`n" | Where-Object { $_.Trim() -ne '' } | Select-Object -First 1).Trim() } else { '' }
        if ($andRaw.Length -gt 60) { $andRaw = $andRaw.Substring(0,60) + '...' }
        $sigRaw = $p['M']; $sigCat = 'Sem Info'
        if ($sigRaw -and $sigRaw.Trim() -ne '') {
            if ($sigRaw -match 'não foi solicitada|não solicitada|não houve|não foi apreendido|não foram apreendidos') { $sigCat = 'Não' }
            elseif ($sigRaw -match 'quebra de sigilo|apreendido|deferindo|requerendo|solicitada') { $sigCat = 'Sim' }
            else { $sigCat = 'Não' }
        }
        $procsArr += [ordered]@{
            proc = if ($p['D']) { $p['D'] } else { '' }
            dp   = $dpRaw
            vara = if ($p['I']) { ($p['I'] -split "`n")[0].Trim() } else { '' }
            and  = $andRaw
            sig  = $sigCat
            ano  = if ($p['A']) { $p['A'] } else { '' }
        }
    }
}
$procsJson = $procsArr | ConvertTo-Json -Compress -Depth 3
$indHtml = [regex]::Replace($indHtml, 'const PROCS = \[.*?\];', "const PROCS = $procsJson;", [System.Text.RegularExpressions.RegexOptions]::Singleline)

# ── COTAÇÕES (internet) — futuros CBOT/NYMEX convertidos por USD-BRL ──────────
# Sem chave de API. Valores de referência internacional (não é preço spot BR).
$usdbrl = 5.20
try{ $fx = Invoke-RestMethod -Uri 'https://economia.awesomeapi.com.br/json/last/USD-BRL' -TimeoutSec 20; $usdbrl=[double]($fx.USDBRL.bid) }catch{ Log "FX USD-BRL falhou, usando $usdbrl" }
function YPrice($sym){ try{ $r=Invoke-RestMethod -Uri "https://query1.finance.yahoo.com/v8/finance/chart/$sym" -Headers @{'User-Agent'='Mozilla/5.0'} -TimeoutSec 20; return [double]$r.chart.result[0].meta.regularMarketPrice }catch{ return 0 } }
$zs=YPrice 'ZS=F'   # soja  cents/bushel (27,2155 kg/bu)
$zc=YPrice 'ZC=F'   # milho cents/bushel (25,401 kg/bu)
$zm=YPrice 'ZM=F'   # farelo USD/short ton (907,18474 kg)
$ho=YPrice 'HO=F'   # diesel (oleo aquec.) USD/galao (3,785411784 L)
$rb=YPrice 'RB=F'   # gasolina USD/galao
$eh=YPrice 'EH=F'   # etanol USD/galao
$GAL=3.785411784
$unit=@{
    soja   = if($zs){ ($zs/100)/27.2155*$usdbrl }else{0}      # R$/kg
    milho  = if($zc){ ($zc/100)/25.401*$usdbrl }else{0}       # R$/kg
    farelo = if($zm){ $zm/907.18474*$usdbrl }else{0}          # R$/kg
    diesel = if($ho){ $ho/$GAL*$usdbrl }else{0}               # R$/L
    gasolina = if($rb){ $rb/$GAL*$usdbrl }else{0}             # R$/L
    etanol = if($eh){ $eh/$GAL*$usdbrl }else{0}               # R$/L
}
Log ("Cotacoes (USD-BRL={0:N2}): soja={1:N3} milho={2:N3} farelo={3:N3} diesel={4:N3} gasolina={5:N3} etanol={6:N3}" -f $usdbrl,$unit.soja,$unit.milho,$unit.farelo,$unit.diesel,$unit.gasolina,$unit.etanol)
function ProdDe($txt){
    $t=("$txt").ToUpper()
    if($t -match 'FARELO'){return 'farelo'}
    if($t -match 'SOJA'){return 'soja'}
    if($t -match 'MILHO'){return 'milho'}
    if($t -match 'DIESEL|[ÓO]LEO DIESEL'){return 'diesel'}
    if($t -match 'ETANOL|[ÁA]LCOOL'){return 'etanol'}
    if($t -match 'GASOLINA'){return 'gasolina'}
    return ''
}

# ── APREENSÕES (TIPO = APREENSÃO) — Produto Recuperado ─────────────────────────
$apre = New-Object System.Collections.Generic.List[object]
for($r=2; $r -le $lastRow; $r++){
    $tipoVal = Cv $data $r $colV
    if($tipoVal -ne 'APREENSÃO'){ continue }
    $dataFmt=''
    $rawI = Cv $data $r $colI
    if($rawI){ try{ $d=[DateTime]::FromOADate([double]$rawI); $dataFmt=$d.ToString('dd/MM/yyyy') }catch{} }
    # Coordenadas (col T) — para o mapa de calor das apreensões
    $aLat=$null; $aLng=$null
    $aCoord = Cv $data $r $colT
    if($aCoord -and ($aCoord -match '^\s*-?\d+[\.,]\d+[\s,]+\s*-?\d+[\.,]\d+\s*$')){
        $ap = ($aCoord -split '[,\s]+' | Where-Object{$_ -ne ''})
        try{ $la=[double]$ap[0]; $lo=[double]$ap[1]
            if($la -gt -35 -and $la -lt 5 -and $lo -gt -75 -and $lo -lt -30){ $aLat=$la; $aLng=$lo } }catch{}
    }
    $subV = Cv $data $r $colW
    $espV = Cv $data $r $colZ
    $qtdV = Cv $data $r $colY
    # Identifica produto (esp tem prioridade; senao tenta subtipo)
    $prod = ProdDe $espV
    if(-not $prod){ $prod = ProdDe $subV }
    # Valor estimado = quantidade x cotacao unitaria (somente os 6 produtos)
    $valNum = 0.0
    if($prod -and $unit.ContainsKey($prod)){
        $qn = 0.0; $qclean = ("$qtdV" -replace '[^\d,\.]','') -replace ',','.'
        if([double]::TryParse($qclean,[ref]$qn)){ $valNum = [math]::Round($qn * [double]$unit[$prod], 2) }
    }
    $apre.Add([ordered]@{
        sub =$subV                # SUBTIPO = categoria
        esp =$espV                # ESPECIFICAÇÃO = produto específico
        prod=$prod                # produto normalizado (soja/milho/farelo/diesel/etanol/gasolina)
        qtd =$qtdV                # QUANTIDADE (kg/litros)
        val =$valNum              # VALOR estimado (R$) pela cotação atual
        tre =Cv $data $r $colM    # TRECHO
        cid =Cv $data $r $colK
        emp =Cv $data $r $colQ
        dt  =$dataFmt
        bo  =Cv $data $r $colAC
        la  =$aLat; lo=$aLng
    })
}
$apreJson = $apre | ConvertTo-Json -Compress -Depth 3
if(-not $apreJson.StartsWith('[')){ $apreJson = "[$apreJson]" }  # 1 item vira objeto; força array
$indHtml = [regex]::Replace($indHtml, 'const APREENSOES = \[.*?\];', "const APREENSOES = $apreJson;", [System.Text.RegularExpressions.RegexOptions]::Singleline)
Log "Apreensões extraídas: $($apre.Count)"

# Cotações para a UI (preços unitários + dólar + data)
$cot=[ordered]@{
    usdbrl=[math]::Round($usdbrl,4)
    data=(Get-Date -Format 'dd/MM/yyyy HH:mm')
    unidades=[ordered]@{
        soja=[math]::Round($unit.soja,4); milho=[math]::Round($unit.milho,4); farelo=[math]::Round($unit.farelo,4)
        diesel=[math]::Round($unit.diesel,4); gasolina=[math]::Round($unit.gasolina,4); etanol=[math]::Round($unit.etanol,4)
    }
}
$cotJson=$cot | ConvertTo-Json -Compress -Depth 4
$indHtml = [regex]::Replace($indHtml, 'const COTACOES = \{.*?\};', "const COTACOES = $cotJson;", [System.Text.RegularExpressions.RegexOptions]::Singleline)

# ── INDICADORES DE CIRCULAÇÃO (aba "MES A MES DASH" do acompanhamento de vandalismo) ──
# Fonte: pasta "Indicadores de Segurança Baixada" — pega o arquivo do ano mais recente.
try{
  $circDir = "C:\Users\risco\OneDrive\Núcleo de Inteligência\15 - DASHBOARD\Planilha\Indicadores de Segurança Baixada"
  $circFile = $null
  if(Test-Path $circDir){
    $circFile = Get-ChildItem -Path $circDir -Recurse -Include *.xlsm,*.xlsx -ErrorAction SilentlyContinue |
      Where-Object{ $_.Name -match 'ACOMPANHAMENTO' -and $_.Name -match '20\d\d' } |
      Sort-Object { if($_.Name -match '(20\d\d)'){[int]$matches[1]}else{0} } -Descending |
      Select-Object -First 1 -ExpandProperty FullName
  }
  if($circFile){
    $tmpCirc="$env:TEMP\circ_update.xlsm"; Copy-Item $circFile $tmpCirc -Force
    $xlC=New-Object -ComObject Excel.Application; $xlC.Visible=$false; $xlC.DisplayAlerts=$false
    $wbC=$xlC.Workbooks.Open($tmpCirc,0,$true)   # ReadOnly
    $wsC=$null; foreach($s in $wbC.Sheets){ if($s.Name -eq 'MES A MES DASH'){ $wsC=$s; break } }
    if($wsC -ne $null){
      $nC=$wsC.UsedRange.Rows.Count
      $dC=$wsC.Range("A1").Resize($nC,16).Value2
      $mesIdx=@{ 'JAN'=0;'FEV'=1;'MAR'=2;'ABR'=3;'MAI'=4;'JUN'=5;'JUL'=6;'AGO'=7;'SET'=8;'OUT'=9;'NOV'=10;'DEZ'=11 }
      $keysC=@('thp','vandalismo','circulacao','sucesso','abertura','descarga','taxaabert')
      $indC=@{}; foreach($k in $keysC){ $indC[$k]=@{ anos=@{} } }
      $anosSet=@{}
      function EnsureC($k,$ano){ if(-not $indC[$k].anos.ContainsKey("$ano")){ $indC[$k].anos["$ano"]=@($null)*12 } }
      function NumC($v){ if($v -eq $null){return $null}; $s="$v".Trim(); if($s -eq '' -or $s -eq '-'){return $null}; try{ return [double]$v }catch{ return $null } }
      for($r=2;$r -le $nC;$r++){
        $ano="$($dC[$r,1])".Trim(); $mes="$($dC[$r,2])".Trim().ToUpper()
        if($ano -notmatch '^20\d\d$' -or [int]$ano -lt 2023 -or -not $mesIdx.ContainsKey($mes)){ continue }
        $suc=NumC $dC[$r,5]; if($suc -eq $null){ continue }   # mês sem dado real -> projeção
        $mi=$mesIdx[$mes]; $anosSet[$ano]=1; foreach($k in $keysC){ EnsureC $k $ano }
        $indC['circulacao'].anos["$ano"][$mi]=NumC $dC[$r,3]
        $indC['vandalismo'].anos["$ano"][$mi]=NumC $dC[$r,4]
        $indC['sucesso'].anos["$ano"][$mi]=[math]::Round($suc*100,2)
        $thpv=NumC $dC[$r,10]; $indC['thp'].anos["$ano"][$mi]=$(if($thpv -ne $null){[math]::Round($thpv,2)}else{$null})
        $indC['descarga'].anos["$ano"][$mi]=NumC $dC[$r,12]
        $indC['abertura'].anos["$ano"][$mi]=NumC $dC[$r,14]
        $tav=NumC $dC[$r,15]; $indC['taxaabert'].anos["$ano"][$mi]=$(if($tav -ne $null){[math]::Round($tav*100,3)}else{$null})
      }
      $anoAtualC=($anosSet.Keys | ForEach-Object{[int]$_} | Measure-Object -Maximum).Maximum
      $objC=[ordered]@{ anoAtual=$anoAtualC; meses=@('Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'); ind=$indC }
      $circJson=$objC | ConvertTo-Json -Depth 6 -Compress
      $indHtml = [regex]::Replace($indHtml, 'const CIRCULACAO = \{.*?\};', "const CIRCULACAO = $circJson;", [System.Text.RegularExpressions.RegexOptions]::Singleline)
      Log "Circulação atualizada de '$([System.IO.Path]::GetFileName($circFile))' (ano atual: $anoAtualC)"
    } else { Log "Aba 'MES A MES DASH' não encontrada — circulação mantida" }
    $wbC.Close($false); $xlC.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xlC)|Out-Null
  } else { Log "Arquivo de circulação não encontrado — circulação mantida" }
}catch{ Log "Falha ao processar circulação: $($_.Exception.Message)" }

$indHtml | Out-File "$outDir\Indicadores_Inteligencia.html" -Encoding utf8
Log "Indicadores_Inteligencia.html atualizado (COORDS: $($coords.Count), PROCS: $($procsArr.Count), APREENSOES: $($apre.Count))"

# ── GIT PUSH ──────────────────────────────────────────────────
Set-Location $outDir
$dt = Get-Date -Format 'yyyy-MM-dd HH:mm'
$prev = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
git add Mapa_Ocorrencias.html Indicadores_Inteligencia.html 2>&1 | Out-Null
$commitMsg = "auto: atualizacao diaria $dt"
git commit -m $commitMsg 2>&1 | Out-Null
git push 2>&1 | Out-Null
$ErrorActionPreference = $prev
Log "Push concluido"
Log "=== FIM DA ATUALIZACAO ==="
