# Nastavení kódování
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$hdoFile = Join-Path $PSScriptRoot "hdo500.csv"
$outputFile = Join-Path $PSScriptRoot "ceny_final_5min.csv"

Write-Host "--- SPUŠTĚNÍ AUTOMATIZACE (CHALUPA - CZK/kWh) ---"

if (-not (Test-Path $hdoFile)) { Write-Error "Soubor $hdoFile nebyl nalezen!"; exit 1 }

# --- 1. AKTUALIZACE EUR ---
try {
    $cnbUrl = "https://www.cnb.cz/cs/financni_trhy/devizovy-trh/kurzy_devizoveho_trhu/denni_kurz.txt"
    $cnbContent = Invoke-WebRequest -Uri $cnbUrl -UseBasicParsing
    $eurLine = $cnbContent.Content -split "`n" | Select-String "EUR"
    $eurRateStr = $eurLine.ToString().Split("|")[4]
    $eurRateNum = [double]($eurRateStr -replace ',', '.')
    
    $hdoLinesAll = Get-Content -Path $hdoFile -Encoding UTF8
    $hdoLinesAll[17] = "EUR;$eurRateStr" # Uložíme aktuální kurz do CSV
    $hdoLinesAll | Set-Content -Path $hdoFile -Encoding UTF8
    Write-Host "Aktuální kurz ČNB: $eurRateNum CZK/EUR"
} catch { 
    $eurRateNum = 25.0 # Záložní kurz, pokud ČNB nejede
    Write-Host "Chyba ČNB, použit záložní kurz." 
}

$rawHdo = Get-Content -Path $hdoFile -Encoding UTF8

# Pomocná funkce pro načtení hodnot z CSV (předpokládáme, že v CSV jsou hodnoty v CZK)
function Get-HdoValue($lineIndex) {
    if ($null -eq $rawHdo -or $rawHdo.Count -le $lineIndex) { return 0 }
    $parts = $rawHdo[$lineIndex] -split ';'
    if ($parts.Count -gt 1) { 
        $val = $parts[1].Trim().Replace('"', '') -replace ',', '.'
        return [double]$val
    }
    return 0
}

# Načtení poplatků (předpokládám, že T1CZK, T2CZK atd. jsou v CZK za MWh)
$T1_czk_mwh = Get-HdoValue 21
$T2_czk_mwh = Get-HdoValue 22
$Priplatek_czk_mwh = Get-HdoValue 24
$srazka_czk_mwh = Get-HdoValue 25

# --- 2. LOGIKA HDO MAPY ---
$hdoMap = @{}
foreach ($line in $rawHdo[1..7]) {
    $parts = $line -split ';'
    $intervals = @()
    # Rozšířeno na více sloupců (až 14 pro 7 intervalů)
    for ($i = 1; $i -lt $parts.Count; $i += 2) {
        if (-not [string]::IsNullOrWhiteSpace($parts[$i]) -and ($parts[$i] -ne $parts[$i+1])) {
            $end = if ($parts[$i+1] -match "00:00|0:00") { [timespan]"1.00:00:00" } else { [timespan]$parts[$i+1] }
            $intervals += @{ start = [timespan]$parts[$i]; end = $end }
        }
    }
    $hdoMap[$parts[0].ToLower().Trim()] = $intervals
}

$dates = @((Get-Date), (Get-Date).AddDays(1))
$finalRows = @()

foreach ($date in $dates) {
    $yyyy = $date.ToString("yyyy"); $mm = $date.ToString("MM"); $dd = $date.ToString("dd")
    # OTE dává ceny v EUR/MWh
    $url = "https://www.ote-cr.cz/pubweb/attachments/01/$yyyy/month$mm/day$dd/DT_15MIN_${dd}_${mm}_${yyyy}_CZ.xlsx"
    $tmpFile = [System.IO.Path]::GetTempFileName() + ".xlsx"

    try {
        Invoke-WebRequest -Uri $url -OutFile $tmpFile
        $oteData = Import-Excel -Path $tmpFile -NoHeader -StartRow 24 -EndRow 119
        foreach ($row in $oteData) {
            # 1. Cena ze Spotu v EUR/MWh převedená na CZK/kWh
            $cenaSpotEUR_MWh = [double]($row.P3 -replace ',', '.')
            $cenaSpotCZK_kWh = ($cenaSpotEUR_MWh * $eurRateNum) / 1000
            
            for ($m = 0; $m -lt 3; $m++) {
                $currTime = ([timespan]$row.P2.Split("-")[0].Trim()).Add([timespan]::FromMinutes($m * 5))
                $dayName = ([System.Globalization.CultureInfo]::GetCultureInfo("cs-CZ")).DateTimeFormat.GetDayName($date.DayOfWeek).ToLower()
                
                $isLow = $false
                if ($hdoMap.ContainsKey($dayName)) {
                    foreach ($int in $hdoMap[$dayName]) {
                        if ($currTime -ge $int.start -and $currTime -lt $int.end) { $isLow = $true; break }
                    }
                }
                
                # 2. Převod fixních poplatků z CZK/MWh na CZK/kWh
                $distribuce_kWh = if ($isLow) { $T2_czk_mwh / 1000 } else { $T1_czk_mwh / 1000 }
                $priplatek_kWh = $Priplatek_czk_mwh / 1000
                $srazka_kWh = $srazka_czk_mwh / 1000
                
                # Výpočet finální ceny v CZK/kWh
                $cenaKonecna = $cenaSpotCZK_kWh + $distribuce_kWh + $priplatek_kWh
                
                $finalRows += [PSCustomObject]@{
                    Datum = $date.ToString("yyyy-MM-dd")
                    Cas = "{0:hh\:mm}" -f $currTime
                    # Ukládáme jako CZK/kWh s 2 desinnými místy
                    Cena_Konecna = "{0:N2}" -f $cenaKonecna
                    Sell = "{0:N2}" -f ($cenaSpotCZK_kWh - $srazka_kWh)
                }
            }
        }
    } catch {} finally { if (Test-Path $tmpFile) { Remove-Item $tmpFile } }
}

if ($finalRows.Count -gt 0) {
    $finalRows | Export-Csv -Path $outputFile -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Force
    Write-Host "Export hotov v CZK/kWh: $outputFile"
}

# --- 3. AUTO-COMMIT ---
if ($env:GITHUB_ACTIONS) {
    git config --global user.name "GitHub Action"
    git config --global user.name "GitHub Action"
    git config --global user.email "action@github.com"
    git add hdo500.csv ceny_final_5min.csv
    git commit -m "Auto-update CZK: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    git push
}
