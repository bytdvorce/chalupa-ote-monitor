# Nastavení kódování a modulů
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
if (-not (Get-Module -ListAvailable ImportExcel)) { Install-Module ImportExcel -Force -Scope CurrentUser }

$hdoFile = Join-Path $PSScriptRoot "hdo500.csv"
$outputFile = Join-Path $PSScriptRoot "ceny_final_5min.csv"
$dates = @((Get-Date), (Get-Date).AddDays(1)) # Definice dnů

Write-Host "--- SPUŠTĚNÍ AUTOMATIZACE ---"

if (-not (Test-Path $hdoFile)) { Write-Error "Soubor $hdoFile nebyl nalezen!"; exit 1 }

# --- 1. AKTUALIZACE EUR A PŘEPOČET ---
try {
    $cnbUrl = "https://www.cnb.cz/cs/financni_trhy/devizovy_trh/kurzy_devizoveho_trhu/denni_kurz.txt"
    $cnbContent = Invoke-WebRequest -Uri $cnbUrl -UseBasicParsing
    $eurLine = $cnbContent.Content -split "`n" | Select-String "EUR"
    $eurRateStr = $eurLine.ToString().Split("|")[4]
    $eurRateNum = [double]($eurRateStr -replace ',', '.')
    
    $hdoLinesAll = Get-Content -Path $hdoFile -Encoding UTF8
    
    function Get-ConvertedValue($fromLineIndex, $label, $rate) {
        $parts = $hdoLinesAll[$fromLineIndex] -split ';'
        $czkVal = [double]($parts[1] -replace ',', '.')
        $eurVal = $czkVal / $rate
        return "$label;$("{0:N4}" -f $eurVal)"
    }

    # Ukládáme přepočítané EUR hodnoty na řádky 10-15
    $hdoLinesAll[10] = Get-ConvertedValue 21 "T1" $eurRateNum
    $hdoLinesAll[11] = Get-ConvertedValue 22 "T2" $eurRateNum
    $hdoLinesAll[12] = Get-ConvertedValue 23 "baterie" $eurRateNum
    $hdoLinesAll[14] = Get-ConvertedValue 24 "priplatek" $eurRateNum
    $hdoLinesAll[15] = Get-ConvertedValue 25 "srazka" $eurRateNum
    $hdoLinesAll[17] = "EUR;$eurRateStr"

    $hdoLinesAll | Set-Content -Path $hdoFile -Encoding UTF8
    $rawHdo = $hdoLinesAll # Předání do proměnné pro další funkci
} catch { Write-Host "Chyba ČNB: $($_.Exception.Message)" }

# --- 2. NAČTENÍ KONSTANT (Z přepočítaných EUR řádků) ---
function Get-HdoValue($lineIndex) {
    if ($null -eq $rawHdo -or $rawHdo.Count -le $lineIndex) { return 0 }
    $parts = $rawHdo[$lineIndex] -split ';'
    if ($parts.Count -gt 1) { 
        $val = $parts[1].Trim().Replace('"', '') -replace ',', '.'
        return [double]$val
    }
    return 0
}

$T1 = Get-HdoValue 10   # Používáme EUR verzi T1
$T2 = Get-HdoValue 11   # Používáme EUR verzi T2
$fix = Get-HdoValue 20  # Fixní složka (ujistěte se, že v CSV je v EUR nebo ji také přepočtěte)

# --- 3. NAČTENÍ HDO MAPY ---
$hdoMap = @{}
$hdoConfig = Import-Excel -Path $hdoFile -NoHeader -StartRow 1 -EndRow 7
foreach ($row in $hdoConfig) {
    $dayName = $row.P1.Trim().ToLower()
    $intervals = @()
    if ($row.P2 -and $row.P3) { $intervals += @{ start = [timespan]::Parse($row.P2); end = [timespan]::Parse($row.P3) } }
    if ($row.P4 -and $row.P5) { $intervals += @{ start = [timespan]::Parse($row.P4); end = [timespan]::Parse($row.P5) } }
    if ($row.P6 -and $row.P7) { $intervals += @{ start = [timespan]::Parse($row.P6); end = [timespan]::Parse($row.P7) } }
    $hdoMap[$dayName] = $intervals
}

# --- 4. VÝPOČET ---
$finalRows = @()
foreach ($date in $dates) {
    $yyyy = $date.ToString("yyyy"); $mm = $date.ToString("MM"); $dd = $date.ToString("dd")
    $url = "https://www.ote-cr.cz/pubweb/attachments/01/$yyyy/month$mm/day$dd/DT_15MIN_${dd}_${mm}_${yyyy}_CZ.xlsx"
    $tmpFile = Join-Path $env:TEMP "ote_$($dd)_$($mm).xlsx"

    try {
        Invoke-WebRequest -Uri $url -OutFile $tmpFile -ErrorAction Stop
        $oteData = Import-Excel -Path $tmpFile -NoHeader -StartRow 24 -EndRow 119
        
        for ($i = 0; $i -lt $oteData.Count; $i += 4) {
            $row = $oteData[$i]
            if (-not $row.P3) { continue }
            
            $cenaSpot = [double]($row.P3 -replace ',', '.')
            $currTime = [timespan]::Parse($row.P2.Split("-")[0].Trim())
            $dayName = ([System.Globalization.CultureInfo]::GetCultureInfo("cs-CZ")).DateTimeFormat.GetDayName($date.DayOfWeek).ToLower()
            
            $isLow = $false
            foreach ($interval in $hdoMap[$dayName]) {
                if ($currTime -ge $interval.start -and $currTime -lt $interval.end) { $isLow = $true; break }
            }
            
            $zaklad = if ($isLow) { $fix + $T2 } else { $fix + $T1 }
            $cenaKonecna = ($zaklad + $cenaSpot) * 1.21

            $finalRows += [PSCustomObject]@{
                Datum = $date.ToString("yyyy-MM-dd")
                Cas = "{0:hh\:mm}" -f $currTime
                Cena_Spot = [math]::Round($cenaSpot, 2)
                Tarif = if ($isLow) { "NT" } else { "VT" }
                Cena_Konecna = [math]::Round($cenaKonecna, 2)
            }
        }
    } catch { Write-Warning "Data pro $dd.$mm. nedostupná." }
}

if ($finalRows.Count -gt 0) {
    $finalRows | Export-Csv -Path $outputFile -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Force
}

# --- 5. GITHUB PUSH ---
if ($env:GITHUB_ACTIONS) {
    git config --global user.name "GitHub Action"
    git config --global user.email "action@github.com"
    git add hdo500.csv ceny_final_5min.csv
    git commit -m "Auto-update: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    git push
}
