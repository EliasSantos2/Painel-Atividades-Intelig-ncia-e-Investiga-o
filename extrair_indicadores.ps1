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
        if($dias[$i]-eq 0){$rc+='null';$rv+='null';$rt+='null';$rva+='null';$rtx+='null'}
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
$R=@{}; $R['2024']=@{}; $R['2025']=@{}

foreach ($t in $TRECHOS) {
    $du=$t.dir.ToUpper()
    $xlsm=Join-Path $BASE "$($t.dir)\2025\$du ACOMPANHAMENTO VANDALISMO 2025.xlsm"
    if (-not (Test-Path $xlsm)) { Write-Host "[$($t.id)] NAO ENCONTRADO: $xlsm"; $R['2024'][$t.id]=$null;$R['2025'][$t.id]=$null; continue }
    Write-Host "[$($t.id)] $(Split-Path $xlsm -Leaf)"
    $d=Expand-Xlsm $xlsm $tmpBase; $xl=Join-Path $d 'xl'
    $ss=Load-SS $xl; $sh=Load-Sheets $xl
    Write-Host "  SS:$($ss.Count) Abas:$($sh.Count)"
    # 2024
    $s24=$null; foreach($n in $sh.Keys){if($n -match 'DADOS 2024|DANOS 2024'){$s24=Join-Path $xl $sh[$n];Write-Host "  2024:'$n'";break}}
    if($null -eq $s24){Write-Host '  [?] 2024 nao encontrada';$R['2024'][$t.id]=$null}
    else{$r=Parse $s24 $ss 'D' 'H' 'L' 'N' 'O';$R['2024'][$t.id]=$r;$m=@();for($i=0;$i-lt 12;$i++){if($r.dias[$i]-gt 0){$m+=$MESES_NOME[$i]}};Write-Host "  2024($($m.Count)):$($m -join ',')"}
    # 2025
    $s25=$null
    if($t.id -eq 't0'){
        foreach($n in $sh.Keys){if($n -match 'DADOS 2025'){$s25=Join-Path $xl $sh[$n];Write-Host "  2025:'$n' (cols E,I,M,O,P)";break}}
        if($null -ne $s25){$r=Parse $s25 $ss 'E' 'I' 'M' 'O' 'P'
            $R['2025'][$t.id]=$r;$m=@();for($i=0;$i-lt 12;$i++){if($r.dias[$i]-gt 0){$m+=$MESES_NOME[$i]}};Write-Host "  2025($($m.Count)):$($m -join ',')"}
        else{$R['2025'][$t.id]=$null;Write-Host '  [?] 2025 nao encontrada'}
    } else {
        foreach($n in $sh.Keys){if($n -match 'DANOS ANO CORRENTE'){$s25=Join-Path $xl $sh[$n];Write-Host "  2025:'$n'";break}}
        if($null -ne $s25){$r=Parse $s25 $ss 'D' 'H' 'L' 'N' 'O'
            $R['2025'][$t.id]=$r;$m=@();for($i=0;$i-lt 12;$i++){if($r.dias[$i]-gt 0){$m+=$MESES_NOME[$i]}};Write-Host "  2025($($m.Count)):$($m -join ',')"}
        else{$R['2025'][$t.id]=$null;Write-Host '  [?] 2025 nao encontrada'}
    }
    Write-Host ''
}

$ts=(Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
$t0_24=TJ $R['2024']['t0']; $t1_24=TJ $R['2024']['t1']; $t2_24=TJ $R['2024']['t2']; $t3_24=TJ $R['2024']['t3']
$t0_25=TJ $R['2025']['t0']; $t1_25=TJ $R['2025']['t1']; $t2_25=TJ $R['2025']['t2']; $t3_25=TJ $R['2025']['t3']
$jsTxt = @"
// Gerado automaticamente por extrair_indicadores.ps1
// Data: $ts
// NAO editar manualmente.

const DADOS_IND = {
  meses: ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'],
  trechos: ['Trecho 0','Trecho 1','Trecho 2','Trecho 3'],
  anos: [2024, 2025],
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
    }
  }
};
"@
$utf8bom=New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($OUT,$jsTxt,$utf8bom)
Write-Host "JS gerado: $OUT"
Write-Host "`n=== Concluido ==="