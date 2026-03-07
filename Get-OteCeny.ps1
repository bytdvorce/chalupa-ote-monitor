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

$T1 = Get-HdoValue 21
$T2 = Get-HdoValue 22
$baterie = Get-HdoValue 12
$batProc = Get-HdoValue 13
$Priplatek = Get-HdoValue 14
$srazka = Get-HdoValue 15
$fix = Get-HdoValue 20

$hdoMap = @{}
# Procházíme řádky 1 až 7 (pondělí až neděle)
foreach ($line in $rawHdo[1..7]) {
    $parts = $line -split ';'
    if ($parts.Count -lt 7) { continue } # Přeskočit neúplné řádky

    $intervals = @()
    # Procházíme 3 dvojice (zapnout/vypnout): indexy 1+2, 3+4, 5+6
    for ($i = 1; $i -le 5; $i += 2) {
        $startStr = $parts[$i].Trim()
        $endStr = $parts[$i+1].Trim()

        if (-not [string]::IsNullOrWhiteSpace($startStr) -and -not [string]::IsNullOrWhiteSpace($endStr)) {
            $start = [timespan]$startStr
            # KLÍČOVÁ OPRAVA: Pokud je konec 00:00, nastavíme ho na 24:00 (1 den)
            $end = if ($endStr -match "^00:00$|^0:00$") { [timespan]"1.00:00:00" } else { [timespan]$endStr }

            # Přidat pouze pokud je interval platný (start < konec)
            if ($start -lt $end) {
                $intervals += @{ start = $start; end = $end }
            }
        }
    }
    # Uložíme pod malými písmeny (pondělí, úterý...)
    $hdoMap[$parts[0].ToLower().Trim()] = $intervals
}

$dates = @((Get-Date), (Get-Date).AddDays(1))
$finalRows = @()

# --- 3. NAČTENÍ DAT Z OTE A VÝPOČET (OPRAVENO NA HODINOVÉ ÚDAJE) ---
foreach ($date in $dates) {
    $yyyy = $date.ToString("yyyy"); $mm = $date.ToString("MM"); $dd = $date.ToString("dd")
    $url = "https://www.ote-cr.cz/pubweb/attachments/01/$yyyy/month$mm/day$dd/DT_15MIN_${dd}_${mm}_${yyyy}_CZ.xlsx"
    $tmpFile = [System.IO.Path]::GetTempFileName() + ".xlsx"

    try {
        Invoke-WebRequest -Uri $url -OutFile $tmpFile
        # Načítáme data z OTE (předpoklad: řádky 24-119 obsahují 15min intervaly)
        # Pro hodinové údaje budeme brát vždy jen první 15min interval každé hodiny
        $oteData = Import-Excel -Path $tmpFile -NoHeader -StartRow 24 -EndRow 119
        
        # Filtrujeme pouze celé hodiny (indexy 0, 4, 8... nebo časy končící :00)
        for ($i = 0; $i -lt $oteData.Count; $i += 4) {
            $row = $oteData[$i]
            $cenaSpot = [double]($row.P3 -replace ',', '.')
            
            # Získání času začátku intervalu (očekávaný formát z OTE např. "00:00 - 01:00")
            $startTimeStr = $row.P2.Split("-")[0].Trim()
            $currTime = [timespan]$startTimeStr
            
            $dayName = ([System.Globalization.CultureInfo]::GetCultureInfo("cs-CZ")).DateTimeFormat.GetDayName($date.DayOfWeek).ToLower()
            
            # Logika HDO
            $isLow = $false
            if ($hdoMap.ContainsKey($dayName)) {
                foreach ($int in $hdoMap[$dayName]) {
                    if ($currTime -ge $int.start -and $currTime -lt $int.end) { 
                        $isLow = $true
                        break 
                    }
                }
            }
            
            $cenaKonecna = if ($isLow) { $fix + $T2 } else { $fix + $T1 }
            
            $finalRows += [PSCustomObject]@{
                Datum        = $date.ToString("yyyy-MM-dd")
                Cas          = "{0:hh\:mm}" -f $currTime
                Cena_Spot    = $cenaSpot
                Tarif        = if ($isLow) { "NT" } else { "VT" }
                Cena_Konecna = "{0:N2}" -f ($cenaKonecna * 1.21) # Včetně DPH
            }
        }
    } catch {
        Write-Host "Data pro den $($date.ToShortDateString()) nebyla nalezena nebo nastala chyba." -ForegroundColor Gray
    } finally { 
        if (Test-Path $tmpFile) { Remove-Item $tmpFile } 
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
