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

# Assembly laden, um ZIP-Dateien zu bearbeiten
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Initialisieren einer Liste zum Speichern aller extrahierten MAC-Adressen als benutzerdefinierte Objekte.
[System.Collections.ArrayList]$allMacAddresses = @()

# PowerShell-Modul "Selenium" importieren
Write-Verbose "Prüfe, ob das Selenium-Modul installiert ist..."
if (-not (Get-Module -ListAvailable -Name Selenium)) {
    Write-Verbose "PowerShell-Modul 'Selenium' wird installiert..."
    Install-Module -Name Selenium -Scope CurrentUser -Force
} else {
    Write-Verbose "PowerShell-Modul 'Selenium' ist bereits installiert."
}

Write-Verbose "Prüfe, ob das Selenium-Modul geladen ist..."
if (-not (Get-Module -Name Selenium)) {
    Import-Module Selenium
    Write-Verbose "PowerShell-Modul 'Selenium' wurde importiert."
} else {
    Write-Verbose "PowerShell-Modul 'Selenium' ist bereits geladen."
}

#endregion

#region Extrahieren
Write-Verbose "Starte die Verarbeitung aller ZIP-Dateien im Verzeichnis '$ZipFilesDirectory'..."

$zipFiles  = Get-ChildItem -Path $ZipFilesDirectory -Filter "*.zip" -File -Recurse
$htmlFiles = @()

<# 
    $zip = Get-Item "C:\Users\michael.schoenburg\Git\ParseHtmlFileContents\t-p031ait_TSR20250610103420_7V7V994.zip"
#>
foreach ($zip in $zipFiles) {
    Write-Verbose "Verarbeite ZIP-Datei: $($zip.FullName)"
    Write-Verbose "Erstelle ein temporäres Verzeichnis für die Extraktion der ZIP- und HTML-Dateien..."
    $baseTempDir = Join-Path -Path $env:TEMP -ChildPath "_ParseHtmlFileContentsSkript"
    if (-not (Test-Path $baseTempDir)) {
        New-Item -ItemType Directory -Path $baseTempDir | Out-Null
    }
    $tempDir = Join-Path -Path $baseTempDir -ChildPath ([System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tempDir | Out-Null
    Write-Verbose "Temporäres Verzeichnis erstellt: $tempDir"

    Write-Verbose "Extrahiere ZIP-Datei: $($zip.FullName)"
    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zip.FullName, $tempDir)

        Write-Verbose "Suche ZIP-Dateien im temporären Verzeichnis '$tempDir'..."
        $SubZipFiles = Get-ChildItem -Path $tempDir -Filter "*.zip" -File
        Write-Verbose "Gefundene ZIP-Dateien: $($SubZipFiles.FullName)"

        foreach ($subZip in $SubZipFiles) {
            Write-Verbose "Durchsuche ZIP-Datei nach HTML-Datei: $($subZip.FullName)"
            $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($subZip.FullName)
            foreach ($entry in $zipArchive.Entries) {
                if ($entry.FullName -match '\.html?$') {
                    Write-Verbose "  - HTML-Datei gefunden: $($entry.FullName) ($($entry.Length) Bytes)"

                    $destPath = Join-Path $tempDir $entry.Name
                    Write-Verbose "    - Extrahiere $($entry.FullName) nach: $destPath"

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
    Write-Verbose "Insgesamt $($htmlFiles.Count) HTML-Dateien gefunden und extrahiert."
    Write-Verbose "Extrahierte HTML-Dateien:"
    foreach ($htmlFile in $htmlFiles) {
        Write-Verbose " - $htmlFile"
    }
}
#endregion

#region Selenium
foreach ($file in $htmlFiles) {
    Write-Verbose "Verarbeite Datei: $($file)"
    $filePath = $file
    <# 
        $filePath = "C:\Users\michael.schoenburg\Git\ParseHtmlFileContents\viewer.html"
    #>
    $fileName = [System.IO.Path]::GetFileName($filePath)

    try {
        # Create a new instance of the Edge driver
        $driver = Start-SeEdge -Quiet -Maximized

        # Navigate to a website
        $driver.Navigate().GoToUrl($filePath)

        # Modell auslesen
        # Write-Verbose "Lese Modellinformationen aus..."
        # $Modell = $driver.FindElementByXPath("/html/body/main/div/div/router-view/div/div/div/div[1]/system-summary/article/div[3]/div/table/tbody/tr[1]/td/span").Text
        # Write-Verbose "Modell: $Modell" -ForegroundColor Green

        # Servername auslesen
        Write-Verbose "Lese Servername aus..."

        try {
            $driver.Navigate().GoToUrl("$($filePath)#/hardware/system-setup")
            $DnsIdracName = $driver.FindElementByXPath("//div[@class='key' and normalize-space(text())='DNS iDRAC Name']")
            $ServernameWithIdrac = $DnsIdracName.FindElementByXPath("following-sibling::*[1]").Text.Trim()
            $Servername = $ServernameWithIdrac -replace '-idrac$', ''
            if ([string]::IsNullOrWhiteSpace($Servername)) {
                throw "Servername konnte nicht ausgelesen werden oder ist leer."
            } else {
                Write-Debug "Servername: $Servername"
            }
        } catch {
            Write-Error "Fehler beim Auslesen des Servernamens für Datei '$fileName': $($_.Exception.Message). Das Skript wird trotzdem weiter ausgeführt. Es kommt zu keinen Folgefehlern." -ErrorAction Continue
            $Servername = "Fehler"
        }

        # iDrac-MAC-Adresse auslesen
        Write-Verbose "Lese iDrac-MAC-Adresse aus..."

        try {
            $driver.Navigate().GoToUrl("$($filePath)#/hardware/systemboard")
            Write-Verbose "Navigiere zu Systemboard-Seite: $($filePath)#/hardware/systemboard"
            $macAddressHeader = $driver.FindElementByXPath("//th[normalize-space(text())='MAC Address']")
            Write-Verbose "Header 'MAC Address' gefunden."
            $IdracMacAddress = $macAddressHeader.FindElementByXPath("following-sibling::td[1]").Text.Trim()
            Write-Debug "iDrac-MAC-Adresse extrahiert: '$IdracMacAddress'"
            if ([string]::IsNullOrWhiteSpace($IdracMacAddress)) {
                throw "Die iDrac-MAC-Adresse konnte nicht ausgelesen werden oder ist leer."
            }
        } catch {
            Write-Error "Fehler beim Auslesen der iDrac-MAC-Adresse für Datei '$fileName': $($_.Exception.Message)" -ErrorAction Continue
            $IdracMacAddress = "Fehler"
        }

        # Tabelle mit MAC-Adressen auslesen
        Write-Verbose "Lese alle MAC-Adressen aus..."

        # Klicke auf den Navigationspunkt "Ethernet"
        Write-Verbose "Navigiere zur Ethernet-Seite..."
        $driver.Navigate().GoToUrl("$($filePath)#/hardware/ethernet")

        # Warte, bis die Seite vollständig geladen ist
        Write-Verbose "Warte auf das Laden der Ethernet-Seite..."
        $wait = New-Object OpenQA.Selenium.Support.UI.WebDriverWait($driver, [System.TimeSpan]::FromSeconds(6))
        try {
            $null = $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists([OpenQA.Selenium.By]::XPath("//td[normalize-space(text())='AutoNegotiation']")))
            $null = $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists([OpenQA.Selenium.By]::XPath("//h2[@class='ui left header' and contains(text(),'Part Information')]")))
            $null = $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists([OpenQA.Selenium.By]::XPath("//tbody/tr/td[normalize-space(text())='AutoNegotiation']")))
            $null = $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists([OpenQA.Selenium.By]::XPath("//th[normalize-space(text())='Firmware']")))
        } catch {
            throw "Timeout: Das erforderliche Element auf der Ethernet-Seite konnte nicht gefunden werden. Die Webseite wurde wohl nicht vollständig geladen. Skript wird abgebrochen."
        }

        # Stelle sicher, dass die Spalte 'Firmware' sichtbar ist
        Write-Verbose "Überprüfe, ob die Spalte 'Firmware' sichtbar ist..."
        $firmwareHeader = $driver.FindElementByXPath("//th[normalize-space(text())='Firmware']")
        if (-not $firmwareHeader.Displayed) {
            throw "Fehler: Die Spalte 'Firmware' ist nicht sichtbar! Tipp: maximiere die Webseite im Browser, um alle Spalten anzuzeigen. Das Skript wird abgebrochen."
        }

        # Die erste Tabelle auf der Seite finden
        Write-Verbose "Suche die erste Tabelle auf der Seite..."
        $table = $driver.FindElementByXPath("(//table)[1]")
        Write-Debug "Tabelle gefunden: $($table.Text)"

        # Alle Zeilen (tr) holen
        Write-Verbose "Extrahiere Zeilen aus der Tabelle..."
        $rows = $table.FindElements([OpenQA.Selenium.By]::TagName("tr"))
        Write-Debug "Zeilen gefunden: $($rows.Text)"

        # Die Header-Zeile extrahieren
        Write-Verbose "Extrahiere Header-Zeile..."
        Write-Debug "Versuche, die Header-Zeile zu extrahieren..."
        $headerRow = $rows[0]
        if ($null -eq $headerRow) {
            Write-Debug "Header-Zeile (rows[0]) ist $null!"
        } else {
            Write-Debug "Header-Zeile gefunden: $($headerRow.Text)"
        }

        $headerElements = $headerRow.FindElements([OpenQA.Selenium.By]::TagName("th"))
        if ($null -eq $headerElements -or $headerElements.Count -eq 0) {
            Write-Debug "Keine <th>-Elemente in der Header-Zeile gefunden!"
        } else {
            Write-Debug "Gefundene <th>-Elemente: $($headerElements.Count)"
            foreach ($el in $headerElements) {
            Write-Debug "Header-Element: $($el.Text)"
            }
        }

        $headers = @()
        foreach ($el in $headerElements) {
            $headerText = $el.Text
            if ($null -eq $headerText -or $headerText.Trim() -eq "") {
                throw "Fehler: Eine Header-Zeile konnte nicht korrekt ausgelesen werden (leer oder null). Das Skript wird abgebrochen."
            } else {
                $headerText = $headerText.Trim()
            }
            Write-Debug "Header-Text extrahiert: '$headerText'"
            $headers += $headerText
        }
        Write-Debug "Finale Header-Liste: $($headers -join ', ')"

        # Die Datenzeilen extrahieren
        Write-Verbose "Extrahiere Datenzeilen..."
        $dataRows = $rows | Select-Object -Skip 1
        Write-Debug "Datenzeilen gefunden: $($dataRows.Text)"

        # Jede Datenzeile in ein PSCustomObject umwandeln
        Write-Verbose "Wandle Datenzeilen in PSCustomObject..."

        $Nics = foreach ($row in $dataRows) {
            $cells = $row.FindElements([OpenQA.Selenium.By]::TagName("td"))
            Write-Debug "Zellen gefunden: $($cells.Text)"
            if ($cells.Count -eq $headers.Count) {
                $obj = [PSCustomObject]@{}
                for ($i = 0; $i -lt $headers.Count; $i++) {
                    $obj | Add-Member -NotePropertyName $headers[$i] -NotePropertyValue $cells[$i].Text.Trim()
                }
                $obj
            }
        }

        # Ausgabe prüfen
        Write-Verbose "Extrahierte Daten aus der Tabelle:"
        $Nics | Format-Table -AutoSize | Out-String | Write-Verbose

        # Extrahiere die Port-Geschwindigkeit aus der Spalte "Model" per Regex und schreibe sie in eine neue Spalte "PortSpeed"
        Write-Verbose "Extrahiere Port-Geschwindigkeit aus der Spalte 'Model' in die neue Spalte 'PortSpeed'..."
        foreach ($nic in $Nics) {
            $model = $nic.Model
            if ($model -match ' \dx(100G) ') {
                $nic | Add-Member -NotePropertyName "PortSpeed" -NotePropertyValue '100 G QSFP' -Force
            } elseif ($model -match ' \dx(10G|25G) ') {
                $nic | Add-Member -NotePropertyName "PortSpeed" -NotePropertyValue '10/25 GbE SFP' -Force
            } else {
                $nic | Add-Member -NotePropertyName "PortSpeed" -NotePropertyValue $null -Force
            }
        }

        # Ausgabe prüfen
        Write-Verbose "Gemappte Tabelle:"
        $Nics | Format-Table -AutoSize | Out-String | Write-Verbose

        #endregion

        #region Excel

        # Ausgabe in Excel eintragen
        Write-Verbose "Trage Daten in Excel ein..."

        # Excel-Objekt erstellen und Datei öffnen
        Write-Verbose "Erstelle Excel-Objekt..."
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        # Hole den absoluten Pfad und Dateinamen aus $PathToExcelFile
        Write-Verbose "Löse Pfad zur Excel-Datei auf: $PathToExcelFile"
        $excelFullPath = Resolve-Path -Path $PathToExcelFile | Select-Object -ExpandProperty Path
        Write-Debug "Excel-Dateipfad aufgelöst: $excelFullPath"
        $excelFileName = [System.IO.Path]::GetFileName($excelFullPath)
        Write-Debug "Excel-Dateiname extrahiert: $excelFileName"
        Write-Verbose "Öffne Excel-Datei: $excelFullPath"
        $workbook = $excel.Workbooks.Open($excelFullPath)
        Write-Verbose "Arbeitsblatt 10 auswählen..."
        $worksheet = $workbook.Worksheets.Item(10)

        $lastRow = $worksheet.UsedRange.Rows.Count
        $ServerRow = $null
        for ($i = 3; $i -le $lastRow; $i++) {
            $cellValue = $worksheet.Cells.Item($i, 3).Value2
            if ($cellValue -eq $Servername) {
                Write-Verbose "Servername gefunden in Zeile $i"
                $ServerRow = $i
                break
            }
        }
        if (-not $ServerRow) {
            # Suche die nächste komplett leere Zeile ab Zeile 3
            for ($i = 3; $i -le ($lastRow + 10); $i++) {
                $rowIsEmpty = $true
                for ($col = 1; $col -le $worksheet.UsedRange.Columns.Count; $col++) {
                    if ($worksheet.Cells.Item($i, $col).Value2) {
                        $rowIsEmpty = $false
                        break
                    }
                }
                if ($rowIsEmpty) {
                    $ServerRow = $i
                    $worksheet.Cells.Item($ServerRow, 3).Value2 = $Servername
                    Write-Verbose "Servername '$Servername' in neue Zeile $ServerRow eingetragen."
                    break
                }
            }
            if (-not $ServerRow) {
                throw "Keine freie Zeile gefunden, um den Servernamen '$Servername' einzutragen."
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

        # Netzwerkkarten 100 G QSFP
        $Nics100G = $Nics.Where({ ($_.Location -like "*Slot*") -or ($_.Location -like "*Integrated*") -and ($_.PortSpeed -eq "100 G QSFP") })
        $j = 0
        foreach ($nic in $Nics100G) {
            $worksheet.Cells.Item($ServerRow, 10 + $j).Value2 = $nic.Location
            $j += 1
            $worksheet.Cells.Item($ServerRow, 10 + $j).Value2 = $nic."MAC Address"
            $j += 1
        }

        # Netzwerkkarten 10/25 GbE SFP
        $Nics1025G = $Nics.Where({ ($_.Location -like "*Slot*") -or ($_.Location -like "*Integrated*") -and ($_.PortSpeed -eq "10/25 GbE SFP") })
        $n = 0
        foreach ($nic in $Nics1025G) {
            $worksheet.Cells.Item($ServerRow, 26 + $n).Value2 = $nic.Location
            $n += 1
            $worksheet.Cells.Item($ServerRow, 26 + $n).Value2 = $nic."MAC Address"
            $n += 1
        }

        # Speichern der Excel-Datei
        Write-Verbose "Speichere Excel-Datei..."
        $workbook.Save()
    } catch {
        Write-Error "  Fehler beim Verarbeiten der Datei '$($fileName)': $($_.Exception.Message)"
        # Fährt fort mit der nächsten Datei, auch wenn ein Fehler auftritt
    } finally {
        # Excel schließen
        Write-Verbose "Schließe Excel..."
        $workbook.Close()
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

        # Browser schließen (optional)
        Write-Verbose "Schließe Browser..."
        $driver.Quit()
    }
}
#endregion

#region Export zur CSV-Datei
if ($allMacAddresses.Count -gt 0) {
    Write-Verbose "Alle MAC-Adressen wurden gesammelt. Exportiere nach '$outputCsvPath'..."
    $allMacAddresses | Export-Csv -Path $outputCsvPath -NoTypeInformation -Encoding UTF8

    Write-Verbose "Erfolgreich aggregierte MAC-Adressen in '$outputCsvPath' gespeichert."
    Write-Verbose "Anzahl der gefundenen Einträge: $($allMacAddresses.Count)"
} else {
    Write-Warning "Keine MAC-Adressen in den verarbeiteten Dateien gefunden. Es wurde keine CSV-Datei erstellt."
}
#endregion

#region Aufräumen
Write-Verbose "Bereinige temporäre Dateien und Ordner..."

# Lösche die temporär angelegten Ordner inkl. Dateien darin
try {
    Remove-Item -Path $baseTempDir -Recurse -Force -ErrorAction Stop
    Write-Verbose "Temporärer Ordner gelöscht: $baseTempDir"
} catch {
    Write-Warning "Konnte temporären Ordner '$baseTempDir' nicht löschen: $($_.Exception.Message)"
}

#endregion

Write-Verbose "Skriptausführung abgeschlossen."
