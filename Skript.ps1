<#
.SYNOPSIS
    Aggregiert Inhalte aus der Spalte "MAC-Adresse" mehrerer HTML-Tabellen aus mehreren ZIP-Dateien und speichert sie in einer CSV-Datei.

.DESCRIPTION
    Dieses Skript durchsucht ein angegebenes Verzeichnis nach HTML-Dateien, extrahiert die Werte aus der angegebenen Spalte (standardmäßig "MAC-Adresse") aller enthaltenen Tabellen und speichert die aggregierten Ergebnisse in einer CSV-Datei.

.PARAMETER htmlFilesDirectory
    Das Verzeichnis, in dem sich die HTML-Dateien befinden.
    Standardwert: "$PSScriptRoot\HTML-Dateien"

.PARAMETER columnNameToSelect
    Der Name der Tabellenspalte, deren Werte extrahiert werden sollen.
    Standardwert: "MAC-Adresse"

.PARAMETER outputCsvPath
    Der Pfad zur Ausgabedatei (CSV), in der die aggregierten MAC-Adressen gespeichert werden.
    Standardwert: "$PSScriptRoot\aggregierte_mac_adressen.csv"

.EXAMPLE
    .\Skript.ps1
    Führt das Skript mit den Standardwerten aus und speichert die aggregierten MAC-Adressen in einer CSV-Datei.

    .\Skript.ps1 -ZipOrdner "Pfad\zu\deinen\ZIPs" -CsvDatei "output.csv" -columnNameToSelect "MAC-Adresse"
    Führt das Skript aus und gibt den Pfad zu den ZIP-Dateien und die gewünschte CSV-Ausgabedatei an.
.LINK
    https://github.com/MichaelSchoenburg/ParseHtmlFileContents

.NOTES
    Autor: Michael Schönburg
    Erstellt: 12.06.2025
#>

#region Variablen
param(
    [Parameter(
        Mandatory = $false,
        HelpMessage = "Dies ist der Pfad zu den ZIP-Dateien, in welchen sich die HTML-Datei 'viewer.html' befindet."
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        if (-not ($_ -is [string] -and $_.Trim().Length -gt 0)) {
            throw "Der Pfad darf nicht leer sein."
        }
        if (-not (Test-Path $_ -PathType Container)) {
            throw "Der angegebene Pfad '$_' existiert nicht oder ist kein Verzeichnis."
        }
        return $true
    })]
    [string]$ZipFilesDirectory = $PSScriptRoot,

    [Parameter(
        Mandatory = $false,
        HelpMessage = "Dies ist der Pfad zur Excel-Datei des Kunden, in welche die Daten eingetragen werden sollen."
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        if (-not ($_ -is [string] -and $_.Trim().Length -gt 0)) {
            throw "Der Dateiname darf nicht leer sein."
        }
        if (-not (Test-Path $_)) {
            throw "Die Datei $_ konnte nicht gefunden werden. Bitte überprüfen Sie den Pfad."
        }
        return $true
    })]
    [string]$PathToExcelFile
    <# 
        $PathToExcelFile = "C:\Users\michael.schoenburg\Git\ParseHtmlFileContents\Airbus REF System Doc Komplett v1.9 - Test.xlsx"
    #>
)

# Initialisieren einer Liste zum Speichern aller extrahierten MAC-Adressen als benutzerdefinierte Objekte.
[System.Collections.ArrayList]$allMacAddresses = @()

# PowerShell-Modul "Selenium" importieren
Write-Host "Prüfe, ob das Selenium-Modul installiert ist..."
if (-not (Get-Module -ListAvailable -Name Selenium)) {
    Write-Host "PowerShell-Modul 'Selenium' wird installiert..."
    Install-Module -Name Selenium -Scope CurrentUser -Force
} else {
    Write-Host "PowerShell-Modul 'Selenium' ist bereits installiert."
}

Write-Host "Prüfe, ob das Selenium-Modul geladen ist..."
if (-not (Get-Module -Name Selenium)) {
    Import-Module Selenium
    Write-Host "PowerShell-Modul 'Selenium' wurde importiert."
} else {
    Write-Host "PowerShell-Modul 'Selenium' ist bereits geladen."
}

#endregion

#region Extrahieren
Write-Host "Starte die Verarbeitung aller HTML-Dateien aus allen ZIP-Dateien im Verzeichnis '$ZipFilesDirectory'..."

$zipFiles  = Get-ChildItem -Path $ZipFilesDirectory -Filter "*.zip" -File -Recurse
$htmlFiles = @()

<# 
    $zip = Get-Item "C:\Users\michael.schoenburg\Git\ParseHtmlFileContents\t-p031ait_TSR20250610103420_7V7V994.zip"
#>
foreach ($zip in $zipFiles) {
    Write-Host "Erstelle ein temporäres Verzeichnis für die Extraktion der ZIP- und HTML-Dateien..."
    $baseTempDir = Join-Path -Path $env:TEMP -ChildPath "_ParseHtmlFileContentsSkript"
    if (-not (Test-Path $baseTempDir)) {
        New-Item -ItemType Directory -Path $baseTempDir | Out-Null
    }
    $tempDir = Join-Path -Path $baseTempDir -ChildPath ([System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    Write-Host "Extrahiere ZIP-Datei: $($zip.FullName)"
    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zip.FullName, $tempDir)
        $SubZipFiles = Get-ChildItem -Path $tempDir -Filter "*.zip" -File
        
        foreach ($subZip in $SubZipFiles) {
            Write-Host "Durchsuche ZIP-Datei nach HTML-Datei: $($subZip.FullName)"
            $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($subZip.FullName)
            foreach ($entry in $zipArchive.Entries) {
                if ($entry.FullName -match '\.html?$') {
                    Write-Host "  - HTML-Datei gefunden: $($entry.FullName) ($($entry.Length) Bytes)"

                    $destPath = Join-Path $tempDir $entry.Name
                    Write-Host "    - Extrahiere $($entry.FullName) nach: $destPath"

                    $entryStream = $entry.Open()
                    $fileStream  = [System.IO.File]::OpenWrite($destPath)
                    $entryStream.CopyTo($fileStream)
                    $fileStream.Close()
                    $entryStream.Close()

                    $htmlFiles += $destPath
                }
            }
            $zipArchive.Dispose()
        }
    } catch {
        Write-Warning "  Fehler beim Lesen von '$($zip.FullName)': $($_.Exception.Message)"
    }
}

if (-not $htmlFiles -or $htmlFiles.Count -eq 0) {
    Write-Error "Es wurden keine HTML-Dateien in den ZIP-Archiven gefunden. Skript wird beendet."
    exit
} else {
    Write-Host "Insgesamt $($htmlFiles.Count) HTML-Dateien gefunden und extrahiert."
    Write-Host "Extrahierte HTML-Dateien:"
    foreach ($htmlFile in $htmlFiles) {
        Write-Host " - $htmlFile"
    }
}
#endregion

#region Selenium
foreach ($file in $htmlFiles) {
    Write-Host "Verarbeite Datei: $($file)"
    $filePath = $file
    <# 
        $filePath = "C:\Users\michael.schoenburg\Git\ParseHtmlFileContents\export.html"
        $filePath = "C:\Users\michael.schoenburg\Git\ParseHtmlFileContents\viewer.html"
    #>
    $fileName = [System.IO.Path]::GetFileName($filePath)

    try {
        # Create a new instance of the Chrome driver
        $driver = Start-SeEdge

        # Navigate to a website
        $driver.Navigate().GoToUrl($filePath)

        # Modell auslesen
        Write-Host "Lese Modellinformationen aus..."
        $Modell = $driver.FindElementByXPath("/html/body/main/div/div/router-view/div/div/div/div[1]/system-summary/article/div[3]/div/table/tbody/tr[1]/td/span").Text
        Write-Host "Modell: $Modell" -ForegroundColor Green

        # Servername auslesen
        Write-Host "Lese Servername aus..."
        $navElement17 = $driver.FindElementByXPath("/html/body/aside/nav/ul/li[16]/div/a")
        $navElement17.Click()

        $DnsIdracName = $driver.FindElementByXPath("//div[@class='key' and normalize-space(text())='DNS iDRAC Name']")
        $ServernameWithIdrac = $DnsIdracName.FindElementByXPath("following-sibling::*[1]").Text.Trim()
        $Servername = $ServernameWithIdrac -replace '-idrac$', ''

        Write-Host "Servername: $Servername" -ForegroundColor Green

        # iDrac-MAC-Adresse auslesen
        Write-Host "Lese iDrac-MAC-Adresse aus..."

        $navElement = $driver.FindElementByXPath("/html/body/aside/nav/ul/li[4]")
        $navElement.Click()

        $IdracMacAddress = $driver.FindElementByXPath("/html/body/main/div/div/router-view/div/div/div/div[2]/article[2]/div[3]/div/table/tbody/tr[5]/td").Text

        Write-Host "iDrac-MAC-Adresse: $IdracMacAddress" -ForegroundColor Green

        # Tabelle mit MAC-Adressen auslesen
        Write-Host "Lese alle MAC-Adressen aus..."

        # Klicke auf den Navigationspunkt "Ethernet"
        $navElement11 = $driver.FindElementByXPath("/html/body/aside/nav/ul/li[11]/div/a")
        $navElement11.Click()

        # Tabelle per XPath finden
        $table = $driver.FindElementByXPath("/html/body/main/div/div/router-view/div/div/div/div[1]/article/div[3]/div/table")

        # Alle Zeilen (tr) holen
        $rows = $table.FindElements([OpenQA.Selenium.By]::TagName("tr"))

        # Die Header-Zeile extrahieren
        $headers = $rows[0].FindElements([OpenQA.Selenium.By]::TagName("th")) | ForEach-Object { $_.Text.Trim() }

        # Die Datenzeilen extrahieren
        $dataRows = $rows | Select-Object -Skip 1

        # Jede Datenzeile in ein PSCustomObject umwandeln
        $Nics = foreach ($row in $dataRows) {
            $cells = $row.FindElements([OpenQA.Selenium.By]::TagName("td"))
            if ($cells.Count -eq $headers.Count) {
                $obj = [PSCustomObject]@{}
                for ($i = 0; $i -lt $headers.Count; $i++) {
                    $obj | Add-Member -NotePropertyName $headers[$i] -NotePropertyValue $cells[$i].Text.Trim()
                }
                $obj
            }
        }

        # Ausgabe prüfen
        $Nics | Format-Table -AutoSize

        #endregion

        #region Excel

        # Ausgabe in Excel eintragen
        Write-Host "Trage Daten in Excel ein..."

        # Excel-Objekt erstellen und Datei öffnen
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $workbook = $excel.Workbooks.Open($PathToExcelFile)
        $worksheet = $workbook.Worksheets.Item(10)

        # Optional: Überschriften in Zeile 1 schreiben
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $worksheet.Cells.Item(1, $i + 1).Value2 = $headers[$i]
        }

        $lastRow = $worksheet.UsedRange.Rows.Count
        for ($i = 3; $i -le $lastRow; $i++) {
            $cellValue = $worksheet.Cells.Item($i, 3).Value2
            if ($cellValue -eq $Servername) {
                Write-Host "Servername gefunden in Zeile $i" -ForegroundColor Green
                $ServerRow = $i
                break
            }
        }

        # MAC Address > MAC
        $worksheet.Cells.Item($ServerRow, 6).Value2 = $IdracMacAddress

        # MAC Address > User / PW
        $worksheet.Cells.Item($ServerRow, 7).Value2 = "root/calvin"

        # Firmware > Port 1
        $FirmwarePort1MacAddress = $Nics.Where({ $_.Location -eq "Embedded 1, Port 1" })."MAC Address"
        $worksheet.Cells.Item($ServerRow, 8).Value2 = $FirmwarePort1MacAddress

        # Firmware > Port 2
        $FirmwarePort2MacAddress = $Nics.Where({ $_.Location -eq "Embedded 2, Port 1" })."MAC Address"
        $worksheet.Cells.Item($ServerRow, 9).Value2 = $FirmwarePort2MacAddress

        # Netzwerkkarten
        $NicsSelected = $Nics.Where({ ($_.Location -like "*Slot*") -or ($_.Location -like "*Integrated*") })
        $j = 0
        foreach ($nic in $NicsSelected) {
            $worksheet.Cells.Item($ServerRow, 10 + $j).Value2 = $nic.Location
            $j += 1
            $worksheet.Cells.Item($ServerRow, 10 + $j).Value2 = $nic."MAC Address"
            $j += 1
        }

        # Speichern der Excel-Datei
        Write-Host "Speichere Excel-Datei..."
        $workbook.Save()
    } catch {
        Write-Error "  Fehler beim Verarbeiten der Datei '$($fileName)': $($_.Exception.Message)"
        # Fährt fort mit der nächsten Datei, auch wenn ein Fehler auftritt
    } finally {
        # Excel schließen
        Write-Host "Schließe Excel..."
        $workbook.Close()
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

        # Browser schließen (optional)
        Write-Host "Schließe Browser..."
        $driver.Quit()
    }
}
#endregion

#region Export zur CSV-Datei
if ($allMacAddresses.Count -gt 0) {
    Write-Host "Alle MAC-Adressen wurden gesammelt. Exportiere nach '$outputCsvPath'..."
    $allMacAddresses | Export-Csv -Path $outputCsvPath -NoTypeInformation -Encoding UTF8

    Write-Host "Erfolgreich aggregierte MAC-Adressen in '$outputCsvPath' gespeichert."
    Write-Host "Anzahl der gefundenen Einträge: $($allMacAddresses.Count)"
} else {
    Write-Warning "Keine MAC-Adressen in den verarbeiteten Dateien gefunden. Es wurde keine CSV-Datei erstellt."
}
#endregion

#region Aufräumen
Write-Host "Bereinige temporäre Dateien und Ordner..."

# Lösche die temporär angelegten Ordner inkl. Dateien darin
try {
    Remove-Item -Path $baseTempDir -Recurse -Force -ErrorAction Stop
    Write-Host "Temporärer Ordner gelöscht: $baseTempDir"
} catch {
    Write-Warning "Konnte temporären Ordner '$baseTempDir' nicht löschen: $($_.Exception.Message)"
}

#endregion

Write-Host "Skriptausführung abgeschlossen."
