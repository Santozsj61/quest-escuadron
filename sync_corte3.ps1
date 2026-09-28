$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 1. Read SEED_PDVS and SEED_ZONES from index.html
$html = Get-Content 'C:\Users\analista.retail\Downloads\quest-escuadron\index.html' -Raw -Encoding UTF8

$pdvMatches = [regex]::Matches($html, '\{\s*id:\s*(\d+),\s*co:\s*''([^'']*)'',\s*zoneId:\s*(\d+),\s*name:\s*''([^'']*)'',\s*leader:\s*''([^'']*)''\s*\}')
$seedPdvs = @()
foreach ($m in $pdvMatches) {
    $storeId = [int]$m.Groups[1].Value
    $pCO = $m.Groups[2].Value.Trim()
    $pZone = [int]$m.Groups[3].Value
    $pName = $m.Groups[4].Value.Trim()
    $pLdr = $m.Groups[5].Value.Trim()

    $pPrefix = if ($pName.ToLower().StartsWith('fq')) { 'fq' } else { 'q' }
    $pNum = $null
    if ($pName.ToLower() -match '^(fq|q|n)\s*0*(\d+)') {
        $pPrefix = $matches[1].ToLower()
        $pNum = [int]$matches[2]
    } elseif ($pCO) {
        $pNum = [int]$pCO
    }

    $seedPdvs += [PSCustomObject]@{
        Id = $storeId
        CO = $pCO
        ZoneId = $pZone
        Name = $pName
        Leader = $pLdr
        Prefix = $pPrefix
        Num = $pNum
    }
}

$zoneMatches = [regex]::Matches($html, '\{\s*id:\s*(\d+),\s*name:\s*''([^'']*)'',\s*leader:\s*''([^'']*)'',\s*emoji:\s*''([^'']*)''\s*\}')
$seedZones = @()
foreach ($m in $zoneMatches) {
    $seedZones += [PSCustomObject]@{
        Id = [int]$m.Groups[1].Value
        Name = $m.Groups[2].Value.Trim()
        Leader = $m.Groups[3].Value.Trim()
    }
}

$matchCache = @{}

function Find-PDVMatch($rawText, $rawLeader) {
    if (-not $rawText) { return $null }
    $cacheKey = "$rawText|$rawLeader"
    if ($matchCache.ContainsKey($cacheKey)) {
        return $matchCache[$cacheKey]
    }

    $textClean = $rawText.Trim().ToLower()
    $leaderClean = if ($rawLeader) { $rawLeader.Trim().ToLower() } else { '' }

    $targetPrefix = $null
    $targetNum = $null

    if ($textClean -match '^(fq|q|n)\s*0*(\d+)') {
        $targetPrefix = $matches[1].ToLower()
        $targetNum = [int]$matches[2]
    } else {
        if ($textClean.Contains('franquicia') -or $textClean.StartsWith('fq')) { $targetPrefix = 'fq' }
        elseif ($textClean.StartsWith('q')) { $targetPrefix = 'q' }
    }

    $res = $null

    if ($targetNum -ne $null) {
        $candidates = @()
        foreach ($p in $seedPdvs) {
            if ($p.Num -eq $targetNum) {
                if (-not $targetPrefix -or $p.Prefix -eq $targetPrefix) {
                    $candidates += $p
                }
            }
        }

        if ($candidates.Count -eq 1) {
            $res = $candidates[0]
        } elseif ($candidates.Count -gt 1) {
            if ($leaderClean) {
                foreach ($c in $candidates) {
                    $l = $c.Leader.ToLower()
                    if ($leaderClean.Contains($l) -or $l.Contains($leaderClean)) {
                        $res = $c
                        break
                    }
                }
            }
            if (-not $res) { $res = $candidates[0] }
        }
    }

    if (-not $res) {
        foreach ($p in $seedPdvs) {
            $pn = $p.Name.ToLower()
            if ($textClean -and ($pn -eq $textClean -or $textClean.Contains($pn) -or $pn.Contains($textClean))) {
                $res = $p
                break
            }
        }
    }

    $matchCache[$cacheKey] = $res
    return $res
}

# 2. Read full 23-30 Presupuesto from Consursos (14).xlsx
Write-Host "Reading Consursos (14).xlsx for total 23-30 budget..."
$wb14 = $excel.Workbooks.Open('C:\Users\analista.retail\Downloads\Consursos (14).xlsx')
$sh14 = $wb14.Sheets.Item(1)
$used14 = $sh14.UsedRange
$rows14Count = $used14.Rows.Count
$cols14Count = $used14.Columns.Count
$valArray14 = $used14.Value2
$wb14.Close($false)

$colDia = -1; $colDescCO = -1; $colLider = -1; $colPptoVal = -1
for ($c = 1; $c -le $cols14Count; $c++) {
    $h = [string]$valArray14[1, $c]
    $h = $h.Trim().ToLower()
    if ($h -like '*d*a de fecha*') { $colDia = $c }
    if ($h.Contains('desc. c.o') -or $h.Contains('desc c.o')) { $colDescCO = $c }
    if ($h -like '*l*der*') { $colLider = $c }
    if ($h -like '*ppto*valor*') { $colPptoVal = $c }
}

$pptoTotalMap = @{}
for ($r = 2; $r -le $rows14Count; $r++) {
    $dia = [string]$valArray14[$r, $colDia]
    if ($dia -eq 'Total') { continue }
    $desc = ([string]$valArray14[$r, $colDescCO]).Trim()
    $lider = ([string]$valArray14[$r, $colLider]).Trim()
    if (-not $desc -or $desc.ToLower().Contains('qst')) { continue }

    $rawPpto = $valArray14[$r, $colPptoVal]
    $pptoVal = if ($rawPpto -ne $null) { [double]$rawPpto } else { 0.0 }

    $matched = Find-PDVMatch $desc $lider
    if ($matched) {
        $idKey = [string]$matched.Id
        if (-not $pptoTotalMap.ContainsKey($idKey)) { $pptoTotalMap[$idKey] = 0.0 }
        $pptoTotalMap[$idKey] += $pptoVal
    }
}
Write-Host "Stores with budget: $($pptoTotalMap.Keys.Count)"

# 3. Read Sales from Consursos y base actualizada.xlsx (Base gnral sheet)
Write-Host "Reading Consursos y base actualizada.xlsx (Base gnral sheet)..."
$wbAct = $excel.Workbooks.Open('C:\Users\analista.retail\Downloads\Consursos y base actualizada.xlsx')
$shBase = $wbAct.Sheets.Item('Base gnral')
$usedBase = $shBase.UsedRange
$valArrayBase = $usedBase.Value2
$rowsBaseCount = $usedBase.Rows.Count
$wbAct.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

$colNombre = 1; $colSupervisor = 3; $colVenta = 4
$ventaTotalMap = @{}

for ($r = 2; $r -le $rowsBaseCount; $r++) {
    $nombre = ([string]$valArrayBase[$r, $colNombre]).Trim()
    $supervisor = ([string]$valArrayBase[$r, $colSupervisor]).Trim()
    if (-not $nombre -or $nombre.ToLower().Contains('qst')) { continue }

    $rawVenta = $valArrayBase[$r, $colVenta]
    $ventaVal = if ($rawVenta -ne $null) { [double]$rawVenta } else { 0.0 }

    $matched = Find-PDVMatch $nombre $supervisor
    if ($matched) {
        $idKey = [string]$matched.Id
        if (-not $ventaTotalMap.ContainsKey($idKey)) { $ventaTotalMap[$idKey] = 0.0 }
        $ventaTotalMap[$idKey] += $ventaVal
    }
}
Write-Host "Stores with sales: $($ventaTotalMap.Keys.Count)"

# 4. Calculate PDV compliance
$pdvDataMap = @{}
foreach ($p in $seedPdvs) {
    $idKey = [string]$p.Id
    $ppto = if ($pptoTotalMap.ContainsKey($idKey)) { $pptoTotalMap[$idKey] } else { 0.0 }
    $vta = if ($ventaTotalMap.ContainsKey($idKey)) { $ventaTotalMap[$idKey] } else { 0.0 }
    $comp = if ($ppto -gt 0) { [math]::Round(($vta / $ppto) * 100, 1) } else { 0.0 }
    $pdvDataMap[$idKey] = $comp
}

# 5. Calculate Zone compliance
$zoneDataArr = @()
$totalCampaignVenta = 0.0
$totalCampaignPpto = 0.0

foreach ($z in $seedZones) {
    $pdvsInZone = @($seedPdvs | Where-Object { $_.ZoneId -eq $z.Id })
    $zPpto = 0.0
    $zVta = 0.0
    foreach ($p in $pdvsInZone) {
        $idKey = [string]$p.Id
        if ($pptoTotalMap.ContainsKey($idKey)) { $zPpto += $pptoTotalMap[$idKey] }
        if ($ventaTotalMap.ContainsKey($idKey)) { $zVta += $ventaTotalMap[$idKey] }
    }
    $totalCampaignPpto += $zPpto
    $totalCampaignVenta += $zVta
    $comp = if ($zPpto -gt 0) { [math]::Round(($zVta / $zPpto) * 100, 1) } else { 0.0 }
    $zoneDataArr += @{
        zoneId = $z.Id
        compliance = $comp
    }
}

$promedioTotal = if ($totalCampaignPpto -gt 0) { [math]::Round(($totalCampaignVenta / $totalCampaignPpto) * 100, 1) } else { 0.0 }
Write-Host "Promedio Nacional: $promedioTotal%"

# 6. Payload for Supabase
$payloadObj = @{
    id = "corte-3-23-al-27-sept"
    name = "Corte 3 - 23 al 27 de Sept (Avance 23-30)"
    date = "2026-09-28"
    data = $zoneDataArr
    pdv_data = $pdvDataMap
}

$jsonPayload = ConvertTo-Json @($payloadObj) -Depth 10

$outputPath = "C:\Users\analista.retail\Downloads\quest-escuadron\corte3_payload.json"
[System.IO.File]::WriteAllText($outputPath, $jsonPayload, [System.Text.Encoding]::UTF8)

# 7. Upload to Supabase
Write-Host "Syncing Corte 3 to Supabase..."
$headers = @{
    'apikey' = 'sb_publishable_jkaQRTZe82IDDPiXTb0jMg_bj19rK0U'
    'Authorization' = 'Bearer sb_publishable_jkaQRTZe82IDDPiXTb0jMg_bj19rK0U'
    'Content-Type' = 'application/json'
    'Prefer' = 'resolution=merge-duplicates'
}

$response = Invoke-RestMethod -Uri 'https://aqgfocnbsjyhcpqfxrsa.supabase.co/rest/v1/quest_cortes' -Method Post -Headers $headers -Body $jsonPayload
Write-Host "SUCCESS! Corte 3 successfully synced to Supabase!"
