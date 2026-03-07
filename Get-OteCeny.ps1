# Nastavení kódování
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
if (-not (Get-Module -ListAvailable ImportExcel)) { Install-Module ImportExcel -Force -Scope CurrentUser }

# Cesta k souboru (univerzální pro GitHub i PC)
$hdoFile = Join-Path $PSScriptRoot "hdo500.csv"
$outputFile = Join-Path $PSScriptRoot "ceny_final_5min.csv"

Write-Host "--- SPUŠTĚNÍ AUTOMATIZACE (GitHub Actions) ---"

# Kontrola, zda soubor existuje, než začneme
if (-not (Test-Path $hdoFile)) {
    Write-Error "Soubor $hdoFile nebyl nalezen v repozitáři!"
    exit 1
}
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

    $hdoLinesAll[10] = Get-ConvertedValue 21 "T1" $eurRateNum
    $hdoLinesAll[11] = Get-ConvertedValue 22 "T2" $eurRateNum
    $hdoLinesAll[12] = Get-ConvertedValue 23 "baterie" $eurRateNum
    $hdoLinesAll[14] = Get-ConvertedValue 24 "priplatek" $eurRateNum
    $hdoLinesAll[15] = Get-ConvertedValue 25 "srazka" $eurRateNum
    $hdoLinesAll[17] = "EUR;$eurRateStr"

    $hdoLinesAll | Set-Content -Path $hdoFile -Encoding UTF8
    Write-Host "Kurz EUR aktualizován: $eurRateStr"
    $rawHdo = Get-Content -Path $hdoFile -Encoding UTF8
} catch { Write-Host "Chyba ČNB: $($_.Exception.Message)" }

# --- 2. NAČTENÍ KONSTANT A VÝPOČET ---
function Get-HdoValue($lineIndex) {
    # Znovu se ujistíme, že máme data v $rawHdo
    if ($null -eq $rawHdo -or $rawHdo.Count -le $lineIndex) {
        Write-Host "Varování: Řádek $lineIndex v souboru chybí nebo je soubor prázdný." -ForegroundColor Yellow
        return 0
    }
    
    $parts = $rawHdo[$lineIndex] -split ';'
    if ($parts.Count -gt 1) { 
        # Vyčištění od případných uvozovek a převod na číslo
        $val = $parts[1].Trim().Replace('"', '') -replace ',', '.'
        return [double]$val
    }
    return 0
}

$T1 = Get-HdoValue 21
$T2 = Get-HdoValue 22
$baterie = Get-HdoValue 12
$batProc = Get-HdoValue 13
$Priplatek = Get-HdoValue 14
$srazka = Get-HdoValue 15
$fix = Get-HdoValue 20

# --- 2. NAČTENÍ HDO MAPY (Všech 3 intervalů) ---
$hdoMap = @{}
$hdoConfig = Import-Excel -Path $hdoFile -NoHeader -StartRow 1 -EndRow 7
foreach ($row in $hdoConfig) {
    $dayName = $row.P1.Trim().ToLower()
    $intervals = @()
    
    # Načtení 1. intervalu
    if ($row.P2 -and $row.P3) { $intervals += @{ start = [timespan]::Parse($row.P2); end = [timespan]::Parse($row.P3) } }
    # Načtení 2. intervalu
    if ($row.P4 -and $row.P5) { $intervals += @{ start = [timespan]::Parse($row.P4); end = [timespan]::Parse($row.P5) } }
    # Načtení 3. intervalu (Tento v původním skriptu chyběl!)
    if ($row.P6 -and $row.P7) { $intervals += @{ start = [timespan]::Parse($row.P6); end = [timespan]::Parse($row.P7) } }
    
    $hdoMap[$dayName] = $intervals
}

# --- 3. STAŽENÍ OTE A HODINOVÝ VÝPOČET ---
$finalRows = @()
$dates = @((Get-Date), (Get-Date).AddDays(1))
foreach ($date in $dates) {
    $yyyy = $date.ToString("yyyy"); $mm = $date.ToString("MM"); $dd = $date.ToString("dd")
    $url = "https://www.ote-cr.cz/pubweb/attachments/01/$yyyy/month$mm/day$dd/DT_15MIN_${dd}_${mm}_${yyyy}_CZ.xlsx"
    $tmpFile = Join-Path $env:TEMP "ote_$($dd)_$($mm).xlsx"

    try {
        Invoke-WebRequest -Uri $url -OutFile $tmpFile -ErrorAction Stop
        $oteData = Import-Excel -Path $tmpFile -NoHeader -StartRow 24 -EndRow 119
        
        # Krokování po 4 (15min -> 1h)
        for ($i = 0; $i -lt $oteData.Count; $i += 4) {
            $row = $oteData[$i]
            if (-not $row.P3) { continue }
            
            $cenaSpot = [double]($row.P3 -replace ',', '.')
            $startTimeStr = $row.P2.Split("-")[0].Trim()
            $currTime = [timespan]::Parse($startTimeStr)
            
            $dayName = ([System.Globalization.CultureInfo]::GetCultureInfo("cs-CZ")).DateTimeFormat.GetDayName($date.DayOfWeek).ToLower()
            
            # Kontrola všech intervalů v HDO mapě
            $isLow = $false
            foreach ($interval in $hdoMap[$dayName]) {
                if ($currTime -ge $interval.start -and $currTime -lt $interval.end) {
                    $isLow = $true
                    break
                }
            }
            
            # Výpočet ceny (včetně 21% DPH)
            $zaklad = if ($isLow) { $fix + $T2 } else { $fix + $T1 }
            $cenaKonecna = ($zaklad + $cenaSpot) * 1.21

            $finalRows += [PSCustomObject]@{
                Datum        = $date.ToString("yyyy-MM-dd")
                Cas          = "{0:hh\:mm}" -f $currTime
                Cena_Spot    = [math]::Round($cenaSpot, 2)
                Tarif        = if ($isLow) { "NT" } else { "VT" }
                Cena_Konecna = [math]::Round($cenaKonecna, 2)
            }
        }
    } catch {
        Write-Warning "Data pro $dd.$mm.$yyyy nejsou dostupná."
    }
}

if ($finalRows.Count -gt 0) {
    $finalRows | Export-Csv -Path $outputFile -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Force
    Write-Host "Export hotov: $outputFile"
}

# --- 3. AUTO-COMMIT (Zápis zpět na GitHub) ---
if ($env:GITHUB_ACTIONS) {
    git config --global user.name "GitHub Action"
    git config --global user.email "action@github.com"
    git add hdo500.csv ceny_final_5min.csv
    git commit -m "Auto-update: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    git push

}
