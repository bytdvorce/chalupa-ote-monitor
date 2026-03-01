# --- NASTAVENÍ A CESTY ---
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$hdoFile = Join-Path $PSScriptRoot "hdo500.csv"
$outputFile = Join-Path $PSScriptRoot "ceny_final_5min.csv"

Write-Host "--- SPUŠTĚNÍ AUTOMATIZACE (CHALUPA - Kč/kWh) ---"

if (-not (Test-Path $hdoFile)) { Write-Error "Soubor $hdoFile nebyl nalezen!"; exit 1 }

# --- 1. NAČTENÍ DAT A KURZU EUR ---
$hdoLinesAll = Get-Content -Path $hdoFile -Encoding UTF8

try {
    $cnbUrl = "https://www.cnb.cz/cs/financni_trhy/devizovy-trh/kurzy-devizoveho-trhu/kurzy-devizoveho-trhu/denni_kurz.txt"
    $cnbContent = Invoke-WebRequest -Uri $cnbUrl -UseBasicParsing
    $eurLine = $cnbContent.Content -split "`n" | Select-String "EUR"
    $eurRateStr = $eurLine.ToString().Split("|")[4]
    $eurRateNum = [double]($eurRateStr -replace ',', '.')
    
    # Aktualizace kurzu v souboru (hledá řádek začínající "EUR;")
    for ($i=0; $i -lt $hdoLinesAll.Count; $i++) {
        if ($hdoLinesAll[$i].StartsWith("EUR;")) {
            $hdoLinesAll[$i] = "EUR;$eurRateStr"
            break
        }
    }
    $hdoLinesAll | Set-Content -Path $hdoFile -Encoding UTF8
    Write-Host "Aktuální kurz ČNB: $eurRateStr Kč"
} catch { 
    Write-Host "Varování: ČNB nedostupná, hledám kurz v souboru."
    $eurLineInFile = $hdoLinesAll | Where-Object { $_.StartsWith("EUR;") }
    if ($eurLineInFile) {
        $eurRateNum = [double](($eurLineInFile -split ";")[1] -replace ',', '.')
    } else { $eurRateNum = 25.0 } # Poslední záchrana
}

# --- 2. BEZPEČNÉ NAČTENÍ HODNOT (PODLE JMÉNA) ---
function Get-ValueByName($name) {
    $line = $hdoLinesAll | Where-Object { $_.StartsWith("$name;") }
    if ($line) {
        $val = ($line -split ";")[1].Trim() -replace ',', '.'
        if ($val -as [double]) { return [double]$val }
    }
    return 0
}

# Načítáme přímo korunové hodnoty (Kč/MWh)
$T1_mwh = Get-ValueByName "T1CZK"
$T2_mwh = Get-ValueByName "T2CZK"
$prip_mwh = Get-ValueByName "priplatekCZK"
$srazka_mwh = Get-ValueByName "srazkaCZK"

# --- 3. LOGIKA HDO MAPY (PRO VÍCE INTERVALŮ) ---
$hdoMap = @{}
# Procházíme řádky, které začínají názvem dne
$dnyTydne = @("pondělí", "úterý", "středa", "čtvrtek", "pátek", "sobota", "neděle")
foreach ($den in $dnyTydne) {
    $line = $hdoLinesAll | Where-Object { $_.StartsWith($den) }
    if ($line) {
        $parts = $line -split ';'
        $intervals = @()
        # Skáčeme po dvojicích zapnout;vypnout (sloupce 1-2, 3-4 atd.)
        for ($i = 1; $i -lt $parts.Count; $i += 2) {
            if (-not [string]::IsNullOrWhiteSpace($parts[$i]) -and ($parts[$i] -ne $parts[$i+1])) {
                $start = [timespan]$parts[$i]
                $end = if ($parts[$i+1] -match "00:00|0:00") { [timespan]"1.00:00:00" } else { [timespan]$parts[$i+1] }
                $intervals += @{ start = $start; end = $end }
            }
        }
        $hdoMap[$den] = $intervals
    }
}

# --- 4. VÝPOČET CEN (OTE -> CZK/kWh) ---
$dates = @((Get-Date), (Get-Date).AddDays(1))
$finalRows = @()

foreach ($date in $dates) {
    $yyyy = $date.ToString("yyyy"); $mm = $date.ToString("MM"); $dd = $date.ToString("dd")
    $url = "https://www.ote-cr.cz/pubweb/attachments/01/$yyyy/month$mm/day$dd/DT_15MIN_${dd}_${mm}_${yyyy}_CZ.xlsx"
    $tmpFile = [System.IO.Path]::GetTempFileName() + ".xlsx"

    try {
        Invoke-WebRequest -Uri $url -OutFile $tmpFile -ErrorAction Stop
        # Import dat z OTE (P2 = Čas, P3 = Cena EUR/MWh)
        $oteData = Import-Excel -Path $tmpFile -NoHeader -StartRow 24 -EndRow 119
        
        foreach ($row in $oteData) {
            $cenaSpotEUR_MWh = [double]($row.P3 -replace ',', '.')
            # Převod na Kč/kWh
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
                
                # Výpočet poplatků v Kč/kWh
                $dist_kWh = if ($isLow) { $T2_mwh / 1000 } else { $T1_mwh / 1000 }
                $prip_kWh = $prip_mwh / 1000
                $sraz_kWh = $srazka_mwh / 1000
                
                $cenaKonecna = $cenaSpotCZK_kWh + $dist_kWh + $prip_kWh

                $finalRows += [PSCustomObject]@{
                    Datum = $date.ToString("yyyy-MM-dd")
                    Cas   = "{0:hh\:mm}" -f $currTime
                    Cena_Konecna = "{0:N2}" -f $cenaKonecna
                    Sell  = "{0:N2}" -f ($cenaSpotCZK_kWh - $sraz_kWh)
                }
            }
        }
    } catch { Write-Host "Data pro den $dd.$mm. nejsou zatím dostupná." }
    finally { if (Test-Path $tmpFile) { Remove-Item $tmpFile } }
}

# --- 5. EXPORT A GITHUB PUSH ---
if ($finalRows.Count -gt 0) {
    $finalRows | Export-Csv -Path $outputFile -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Force
    Write-Host "Soubor uložen: $outputFile"

    if ($env:GITHUB_ACTIONS) {
        git config --global user.name "GitHub Action"
        git config --global user.email "action@github.com"
        git add hdo500.csv ceny_final_5min.csv
        git commit -m "Auto-update: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
        git push
    }
}
