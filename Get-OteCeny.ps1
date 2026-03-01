# Nastavení kódování
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

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

$T1 = Get-HdoValue 10
$T2 = Get-HdoValue 11
$baterie = Get-HdoValue 12
$batProc = Get-HdoValue 13
$Priplatek = Get-HdoValue 14
$srazka = Get-HdoValue 15

$hdoMap = @{}
foreach ($line in $rawHdo[1..7]) {
    $parts = $line -split ';'
    $intervals = @()
    for ($i = 1; $i -lt 4; $i += 2) {
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
    $url = "https://www.ote-cr.cz/pubweb/attachments/01/$yyyy/month$mm/day$dd/DT_15MIN_${dd}_${mm}_${yyyy}_CZ.xlsx"
    $tmpFile = [System.IO.Path]::GetTempFileName() + ".xlsx"

    try {
        Invoke-WebRequest -Uri $url -OutFile $tmpFile
        $oteData = Import-Excel -Path $tmpFile -NoHeader -StartRow 24 -EndRow 119
        foreach ($row in $oteData) {
            $cenaSpot = [double]($row.P3 -replace ',', '.')
            for ($m = 0; $m -lt 3; $m++) {
                $currTime = ([timespan]$row.P2.Split("-")[0].Trim()).Add([timespan]::FromMinutes($m * 5))
                $dayName = ( [System.Globalization.CultureInfo]::GetCultureInfo("cs-CZ")).DateTimeFormat.GetDayName($date.DayOfWeek).ToLower()
                
                $isLow = $false
                if ($hdoMap.ContainsKey($dayName)) {
                    foreach ($int in $hdoMap[$dayName]) {
                        if ($currTime -ge $int.start -and $currTime -lt $int.end) { $isLow = $true; break }
                    }
                }
                
                $cenaKonecnaEUR = if ($isLow) { $cenaSpot + $T2 + $Priplatek } else { $cenaSpot + $T1 + $Priplatek }
                $cenaKonecna = ($cenaKonecnaEUR*$rate)/1000
                $divisor = if ($batProc -gt 0) { $batProc } else { 1 }
                
                $finalRows += [PSCustomObject]@{
                    Datum = $date.ToString("yyyy-MM-dd")
                    Cas = "{0:hh\:mm}" -f $currTime
                    Cena_Spot = $cenaSpot
                    Tarif = if ($isLow) { "NT" } else { "VT" }
                    Cena_Konecna = "{0:N2}" -f $cenaKonecna
#                    Cena_Bat = "{0:N2}" -f ($baterie + ($cenaKonecna / $divisor))
#                    Sell = "{0:N2}" -f ($cenaSpot - $srazka)
                }
            }
        }
    } catch {} finally { if (Test-Path $tmpFile) { Remove-Item $tmpFile } }
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
