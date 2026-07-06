# extrair_indicadores.ps1
# Extrai indicadores das planilhas XLSM e gera dados_indicadores.js
# Execute: powershell -ExecutionPolicy Bypass -File extrair_indicadores.ps1

$ErrorActionPreference = 'Continue'
$BASE = "C:\Users\risco\Rumo\Indicadores de Segurança - Resultado Geral"
$OUT  = Join-Path $PSScriptRoot "dados_indicadores.js"
$TRECHOS = @(
    @{ id='t0'; dir='Trecho 0' },
    @{ id='t1'; dir='Trecho 1' },
    @{ id='t2'; dir='Trecho 2' },
    @{ id='t3'; dir='Trecho 3' }
)
$MESES_NOME = @('Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez')
$MESES_IDX = @{ 'JAN'=0;'FEV'=1;'MAR'=2;'ABR'=3;'MAI'=4;'MAIO'=4;'JUN'=5;'JUL'=6;'AGO'=7;'SET'=8;'OUT'=9;'NOV'=10;'DEZ'=11 }

function Expand-Xlsm([string]$xlsm,[string]$tmpBase) {
    $key=([System.IO.Path]::GetFileNameWithoutExtension($xlsm)) -replace '[^a-zA-Z0-9]','_'
    $d=Join-Path $tmpBase $key
    if (Test-Path $d) { $null=Remove-Item $d -Recurse -Force -ErrorAction SilentlyContinue }
    $null=New-Item -ItemType Directory -Path $d -Force
    $z="$d.zip"; Copy-Item $xlsm $z -Force
    Expand-Archive $z -DestinationPath $d -Force
    $null=Remove-Item $z -Force -ErrorAction SilentlyContinue
    return $d
}
function Load-SS([string]$xlDir) {
    $p=Join-Path $xlDir 'sharedStrings.xml'; if(-not(Test-Path $p)){return @()}
    $raw=[System.IO.File]::ReadAllText($p,[System.Text.Encoding]::UTF8)
    $arr=[System.Collections.Generic.List[string]]::new()
    foreach ($m in [regex]::Matches($raw,'<t[^>]*>([^<]*)</t>')) { $arr.Add($m.Groups[1].Value) }
    return $arr.ToArray()
}
function Load-Sheets([string]$xlDir) {
    $wb=[System.IO.File]::ReadAllText((Join-Path $xlDir 'workbook.xml'),[System.Text.Encoding]::UTF8)
    $re=[System.IO.File]::ReadAllText((Join-Path $xlDir '_rels\workbook.xml.rels'),[System.Text.Encoding]::UTF8)
    $rm=@{}
    foreach ($r in [regex]::Matches($re,'Id="([^"]+)"[^>]+Target="(worksheets/[^"]+)"')) { $rm[$r.Groups[1].Value]=$r.Groups[2].Value }
    $res=@{}
    foreach ($s in [regex]::Matches($wb,'name="([^"]+)"[^>]+r:id="([^"]+)"')) {
        if ($rm.ContainsKey($s.Groups[2].Value)) { $res[$s.Groups[1].Value]=$rm[$s.Groups[2].Value] }
    }
    return $res
}
function Get-Str([string]$fc,[string[]]$ss) {
    if (-not $fc) { return $null }
    $ta=[regex]::Match($fc,'<c[^>]+\bt="([^"]+)"'); $ct=if($ta.Success){$ta.Groups[1].Value}else{''}
    $v=[regex]::Match($fc,'<v>([^<]*)</v>'); if(-not $v.Success){return $null}
    $r=$v.Groups[1].Value.Trim(); if($r -eq ''){return $null}
    if ($ct -eq 's') { $i=[int]$r; if($i -ge 0 -and $i -lt $ss.Count){return $ss[$i]}; return $null }
    return $r
}
function Get-Num([string]$fc) {
    if (-not $fc) { return $null }
    $ta=[regex]::Match($fc,'<c[^>]+\bt="([^"]+)"'); $ct=if($ta.Success){$ta.Groups[1].Value}else{''}
    if ($ct -eq 's' -or $ct -eq 'b' -or $ct -eq 'e') { return $null }
    $v=[regex]::Match($fc,'<v>([^<]*)</v>'); if(-not $v.Success){return $null}
    $r=$v.Groups[1].Value.Trim(); if($r -eq ''){return $null}
    try { return [double]$r } catch { return $null }
}
function Parse([string]$f,[string[]]$ss,[string]$cC,[string]$cV,[string]$cT,[string]$cVA,[string]$cVT) {
    if(-not(Test-Path $f)){Write-Host "  [ERR] $f"; return $null}
    $xml=[System.IO.File]::ReadAllText($f,[System.Text.Encoding]::UTF8)
    $aC=New-Object double[] 12;$aV=New-Object double[] 12;$aT=New-Object double[] 12
    $aVA=New-Object double[] 12;$aVT=New-Object double[] 12;$dias=New-Object int[] 12
    $opts=[System.Text.RegularExpressions.RegexOptions]::Singleline
    $rowRx=[regex]::new('<row r="(\d+)"[^>]*>(.*?)</row>',$opts)
    $cellRx=[regex]::new('<c r="([A-Z]+)\d+"[^>]*>(?:(?!</c>).)*</c>',$opts)
    foreach ($rm in $rowRx.Matches($xml)) {
        $rn=[int]$rm.Groups[1].Value; if($rn -lt 3){continue}
        $cm=@{}; foreach($c in $cellRx.Matches($rm.Groups[2].Value)){$cm[$c.Groups[1].Value]=$c.Value}
        if(-not $cm.ContainsKey('A')){continue}
        $code=Get-Str $cm['A'] $ss; if($null -eq $code -or $code -eq ''){continue}
        $ms=[regex]::Match($code.Trim(),'[A-Z]+$'); if(-not $ms.Success){continue}
        $mSuf=$ms.Value.ToUpper(); if(-not $MESES_IDX.ContainsKey($mSuf)){continue}
        $mi=$MESES_IDX[$mSuf]
        $vC=if($cm.ContainsKey($cC)){Get-Num $cm[$cC]}else{$null}; if($null -eq $vC){continue}
        $vV=if($cm.ContainsKey($cV)){Get-Num $cm[$cV]}else{$null}
        $vT=if($cm.ContainsKey($cT)){Get-Num $cm[$cT]}else{$null}
        $vVA=if($cm.ContainsKey($cVA)){Get-Num $cm[$cVA]}else{$null}
        $vVT=if($cm.ContainsKey($cVT)){Get-Num $cm[$cVT]}else{$null}
        $dias[$mi]++; $aC[$mi]+=$vC
        if($null -ne $vV){$aV[$mi]+=$vV}; if($null -ne $vT){$aT[$mi]+=$vT}
        if($null -ne $vVA){$aVA[$mi]+=$vVA}; if($null -ne $vVT){$aVT[$mi]+=$vVT}
    }
    $rc=@();$rv=@();$rt=@();$rva=@();$rtx=@()
    for($i=0;$i-lt 12;$i++){
        # filtra meses sem dados reais: zero dias, ou circ total <= numero de dias (template com 1/dia)
        if($dias[$i]-eq 0 -or $aC[$i]-le $dias[$i]){$rc+='null';$rv+='null';$rt+='null';$rva+='null';$rtx+='null'}
        else{
            $rc+=[string][int]$aC[$i];$rv+=[string][int]$aV[$i]
            $rt+=[string][math]::Round($aT[$i]*24,2);$rva+=[string][int]$aVA[$i]
            if($aVT[$i]-gt 0){$rtx+=[string][math]::Round(($aVA[$i]/$aVT[$i])*100,2)}else{$rtx+='null'}
        }
    }
    return @{c=$rc;v=$rv;t=$rt;va=$rva;tx=$rtx;dias=$dias}
}
function FA([string[]]$a){'['+($a -join ',')+']'}
$NUL='{ circ:[null,null,null,null,null,null,null,null,null,null,null,null], vand:[null,null,null,null,null,null,null,null,null,null,null,null], thp:[null,null,null,null,null,null,null,null,null,null,null,null], vagab:[null,null,null,null,null,null,null,null,null,null,null,null], txab:[null,null,null,null,null,null,null,null,null,null,null,null] }'
function TJ($r){if($null -eq $r){return $NUL};"{ circ:$(FA $r.c), vand:$(FA $r.v), thp:$(FA $r.t), vagab:$(FA $r.va), txab:$(FA $r.tx) }"}

# Main
$tmpBase=Join-Path $env:TEMP "xlsm_rumo_x"
New-Item -ItemType Directory -Path $tmpBase -Force -ErrorAction SilentlyContinue | Out-Null
Write-Host "=== Extraindo indicadores ===" ; Write-Host ""
# usa chaves planas "ano_id" para evitar problema de escopo com hashtable aninhada
$RES = [System.Collections.Hashtable]::new()

foreach ($t in $TRECHOS) {
    $du=$t.dir.ToUpper()
    # ── 2025 xlsm: contém DADOS 2024 e DADOS/DANOS ANO CORRENTE (2025) ──
    $xlsm=Join-Path $BASE "$($t.dir)\2025\$du ACOMPANHAMENTO VANDALISMO 2025.xlsm"
    if (-not (Test-Path $xlsm)) { Write-Host "[$($t.id)] NAO ENCONTRADO (2025): $xlsm" }
    else {
        Write-Host "[$($t.id)] $(Split-Path $xlsm -Leaf)"
        $d=Expand-Xlsm $xlsm $tmpBase; $xl=Join-Path $d 'xl'
        $ss=Load-SS $xl; $sh=Load-Sheets $xl
        Write-Host "  SS:$($ss.Count) Abas:$($sh.Count)"
        # 2024
        $s24=$null; foreach($n in $sh.Keys){if($n -match 'DADOS 2024|DANOS 2024'){$s24=Join-Path $xl $sh[$n];Write-Host "  2024:'$n'";break}}
        if($null -eq $s24){Write-Host '  [?] 2024 nao encontrada';$RES["2024_$($t.id)"]=$null}
        else{$r=Parse $s24 $ss 'D' 'H' 'L' 'N' 'O';$RES["2024_$($t.id)"]=$r;$m=@(0..11|Where-Object{$r.dias[$_]-gt 0}|ForEach-Object{$MESES_NOME[$_]});Write-Host "  2024($($m.Count)):$($m -join ',')"}
        # 2025
        $s25=$null
        if($t.id -eq 't0'){
            foreach($n in $sh.Keys){if($n -match 'DADOS 2025'){$s25=Join-Path $xl $sh[$n];Write-Host "  2025:'$n'";break}}
            if($null -ne $s25){$r=Parse $s25 $ss 'E' 'I' 'M' 'O' 'P';$RES["2025_$($t.id)"]=$r
                $m=@(0..11|Where-Object{$r.dias[$_]-gt 0}|ForEach-Object{$MESES_NOME[$_]});Write-Host "  2025($($m.Count)):$($m -join ',')"}
            else{$RES["2025_$($t.id)"]=$null;Write-Host '  [?] 2025 nao encontrada'}
        } else {
            foreach($n in $sh.Keys){if($n -match 'DANOS ANO CORRENTE|DADOS ANO CORRENTE'){$s25=Join-Path $xl $sh[$n];Write-Host "  2025:'$n'";break}}
            if($null -ne $s25){$r=Parse $s25 $ss 'D' 'H' 'L' 'N' 'O';$RES["2025_$($t.id)"]=$r
                $m=@(0..11|Where-Object{$r.dias[$_]-gt 0}|ForEach-Object{$MESES_NOME[$_]});Write-Host "  2025($($m.Count)):$($m -join ',')"}
            else{$RES["2025_$($t.id)"]=$null;Write-Host '  [?] 2025 nao encontrada'}
        }
    }
    # ── 2026 xlsm: DADOS/DANOS ANO CORRENTE com colunas E,I,M,O,P ──
    $xlsm26=Join-Path $BASE "$($t.dir)\2026\$du ACOMPANHAMENTO VANDALISMO 2026.xlsm"
    if (-not (Test-Path $xlsm26)) { Write-Host "  [?] 2026 nao encontrado"; $RES["2026_$($t.id)"]=$null }
    else {
        Write-Host "  [2026] $(Split-Path $xlsm26 -Leaf)"
        $d26=Expand-Xlsm $xlsm26 $tmpBase; $xl26=Join-Path $d26 'xl'
        $ss26=Load-SS $xl26; $sh26=Load-Sheets $xl26
        $s26=$null; foreach($n in $sh26.Keys){if($n -match 'DADOS ANO CORRENTE|DANOS ANO CORRENTE'){$s26=Join-Path $xl26 $sh26[$n];Write-Host "  2026:'$n'";break}}
        if($null -eq $s26){Write-Host '  [?] 2026 nao encontrada';$RES["2026_$($t.id)"]=$null}
        else{$r=Parse $s26 $ss26 'E' 'I' 'M' 'O' 'P';$RES["2026_$($t.id)"]=$r
            $m=@(0..11|Where-Object{$r.dias[$_]-gt 0}|ForEach-Object{$MESES_NOME[$_]});Write-Host "  2026($($m.Count)):$($m -join ',')"}
    }
    Write-Host ''
}

$ts=(Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
$t0_24=TJ $RES["2024_t0"]; $t1_24=TJ $RES["2024_t1"]; $t2_24=TJ $RES["2024_t2"]; $t3_24=TJ $RES["2024_t3"]
$t0_25=TJ $RES["2025_t0"]; $t1_25=TJ $RES["2025_t1"]; $t2_25=TJ $RES["2025_t2"]; $t3_25=TJ $RES["2025_t3"]
$t0_26=TJ $RES["2026_t0"]; $t1_26=TJ $RES["2026_t1"]; $t2_26=TJ $RES["2026_t2"]; $t3_26=TJ $RES["2026_t3"]
$jsTxt = @"
// Gerado automaticamente por extrair_indicadores.ps1
// Data: $ts
// NAO editar manualmente.

const DADOS_IND = {
  meses: ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'],
  trechos: ['Trecho 0','Trecho 1','Trecho 2','Trecho 3'],
  anos: [2024, 2025, 2026],
  geradoEm: '$ts',
  data: {
    '2024': {
      't0': $t0_24,
      't1': $t1_24,
      't2': $t2_24,
      't3': $t3_24
    },
    '2025': {
      't0': $t0_25,
      't1': $t1_25,
      't2': $t2_25,
      't3': $t3_25
    },
    '2026': {
      't0': $t0_26,
      't1': $t1_26,
      't2': $t2_26,
      't3': $t3_26
    }
  }
};
"@
$utf8bom=New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($OUT,$jsTxt,$utf8bom)
Write-Host "JS gerado: $OUT"
Write-Host "`n=== Concluido ==="