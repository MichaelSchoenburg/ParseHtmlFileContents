<#
.SYNOPSIS
    Aggregiert Inhalte aus der Spalte "MAC-Adresse" mehrerer HTML-Tabellen aus mehreren ZIP-Dateien und speichert sie in einer CSV-Datei.

.DESCRIPTION
    Dieses Skript durchsucht ein angegebenes Verzeichnis nach HTML-Dateien, extrahiert die Werte aus der angegebenen Spalte (standardmaeßig "MAC-Adresse") aller enthaltenen Tabellen und speichert die aggregierten Ergebnisse in einer CSV-Datei.

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
    Fuehrt das Skript mit den Standardwerten aus und speichert die aggregierten MAC-Adressen in einer CSV-Datei.

    .\Skript.ps1 -ZipOrdner "Pfad\zu\deinen\ZIPs" -CsvDatei "output.csv" -columnNameToSelect "MAC-Adresse"
    Fuehrt das Skript aus und gibt den Pfad zu den ZIP-Dateien und die gewuenschte CSV-Ausgabedatei an.
.LINK
    https://github.com/MichaelSchoenburg/ParseHtmlFileContents

.NOTES
    Autor: Michael Schoenburg
    Erstellt: 12.06.2025
#>

#region Parameter

param(
    # Pfad zum Export der Support-Daten
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
    [string]
    $ZipFilesDirectory = $PSScriptRoot,

    # Pfad zur Excel-Datei
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
            throw "Die Datei $_ konnte nicht gefunden werden. Bitte ueberpruefen Sie den Pfad."
        }
        return $true
    })]
    [string]
    $PathToExcelFile,

    # Pfad zum Chrome Web Driver
    [Parameter(
        Mandatory = $false,
        HelpMessage = "Dies ist der Pfad zum Ordner, in welchem der Chrome Web Driver (muss 'chromedriver.exe' heißen) liegt."
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
    [string]
    $PathToWebDriverDirectory,
    
    # Pfad zur Chrome.exe
    [Parameter(
        Mandatory = $false,
        HelpMessage = "Dies ist der Pfad zur Chrome.exe, falls eine andere Version genutzt werden soll, als die installierte Chrome-Standard-Version."
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        if (-not ($_ -is [string] -and $_.Trim().Length -gt 0)) {
            throw "Der Dateiname darf nicht leer sein."
        }
        if (-not (Test-Path $_)) {
            throw "Die Datei $_ konnte nicht gefunden werden. Bitte ueberpruefen Sie den Pfad."
        }
        return $true
    })]
    [string]
    $PathToChromeBinary,

    # Pfad fuer die Log-Datei
    [Parameter(
        Mandatory = $false,
        HelpMessage = "Dies ist der Pfad zur Log-Datei."
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
    [string]
    $PathToLogDirectory = $PSScriptRoot,

    # Silent-Mode
    [Parameter(
        HelpMessage = "Im Silent-Mode gibt das Skript keinen Output/Text in der Konsole aus, wodurch es schneller laeuft."
    )]
    [switch]
    $Silent = $false
)

#endregion

#region Funktionen$LogFilePath = "$($PSScriptRoot)\ParseHtmlFileContents-$(Get-Date -Format "dd.MM.yyyy-HH.mm.ss").log"

function Write-ConsoleLog {
    <#
    .SYNOPSIS
    Protokolliert ein Ereignis in der Konsole.
    
    .DESCRIPTION
    Schreibt Text in die Konsole mit dem aktuellen Datum (US-Format) davor.
    
    .PARAMETER Text
    Ereignis/Text, der in die Konsole ausgegeben werden soll.
    
    .EXAMPLE
    Write-ConsoleLog -Text 'Subscript XYZ aufgerufen.'
    
    Lange Form
    .EXAMPLE
    Log 'Subscript XYZ aufgerufen.'
    
    Kurze Form
    #>
    [alias('Log')]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
        Position = 0)]
        [string]
        $Text
    )

    if (-not $Silent) {
        Write-Host "$( Get-Date -Format 'MM/dd/yyyy HH:mm:ss' ) - $( $Text )"
    } else {
        Add-Content -Path $LogFilePath -Value "$( Get-Date -Format 'MM/dd/yyyy HH:mm:ss' ) - $( $Text )"
    }
}

function Wait-ForElement {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Driver,
        [Parameter(Mandatory = $true)]
        [string]$Text,
        [Parameter(Mandatory = $true)]
        [string]$XPath,
        [int]$TimeoutSeconds = 4
    )

    $elementFound = $false
    for ($i = 0; $i -lt $TimeoutSeconds; $i++) {
        try {
            $element = $driver.FindElementByXPath($XPath)
            log "i = $($i)"
            log "element: $($element)"
            log "element.Displayed = $($element.Displayed)"
            log "element.Text = $($element.Text)"
            if ($element.Displayed -and $element.Text -eq $Text) {
                $elementFound = $true
                log "Element '$Text' gefunden und sichtbar."
                break
            }
            log "found = $($elementFound)"
            Start-Invoke-Sleep -Seconds 1
        } catch {
            # nix
        }
    }
    if (-not $elementFound) {
        throw "Timeout: Das Element mit dem Text '$Text' und XPath '$XPath' wurde nicht gefunden oder ist nicht sichtbar."
    } elseif ($elementFound) {
        log "Element '$Text' ist sichtbar."
    }
}

function Invoke-Sleep {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [int]$Seconds
    )
    for ($i = $Seconds; $i -gt 0; $i--) {
        Log "Schlafe fuer Sekunden = $i"
        Start-Sleep -Seconds 1
    }
}

#endregion

#region Variablen

$baseTempDir = Join-Path -Path $env:TEMP -ChildPath "_ParseHtmlFileContentsSkript"
$LogFilePath = "$($PathToLogDirectory)\ParseHtmlFileContents-$(Get-Date -Format "dd.MM.yyyy-HH.mm.ss").log"

#endregion

#region Transkript

if ($Silent) {
    $null = New-Item -Path $LogFilePath -ItemType File
} else {
    Start-Transcript -Path $LogFilePath
}

#endregion

#region Initialisierung

# Assembly laden, um ZIP-Dateien zu bearbeiten
Log "Lade Assembly 'System.IO.Compression.FileSystem'..."
Add-Type -AssemblyName System.IO.Compression.FileSystem

# PowerShell-Modul "Selenium" importieren
Log "Pruefe, ob das Selenium-Modul installiert ist..."
if (-not (Get-Module -ListAvailable -Name Selenium)) {
    Log "PowerShell-Modul 'Selenium' wird installiert..."
    Install-Module -Name Selenium -Scope CurrentUser
} else {
    Log "PowerShell-Modul 'Selenium' ist bereits installiert."
}

Log "Pruefe, ob das Selenium-Modul geladen ist..."
if (-not (Get-Module -Name Selenium)) {
    Import-Module Selenium
    Log "PowerShell-Modul 'Selenium' wurde importiert."
} else {
    Log "PowerShell-Modul 'Selenium' ist bereits geladen."
}

#endregion

if (-not $Silent) {
    Log "Die Variable 'Silent' steht auf '$($Silent)'."
}

#region Extrahieren
try {
    Log "Pruefe, ob ein temporaeres Verzeichnis fuer die Extraktion der ZIP- und HTML-Dateien existiert..."
    if (-not (Test-Path $baseTempDir)) {
        Log "  Dies wurde nicht gefunden. Erstelle ein temporaeres Verzeichnis fuer die Extraktion der ZIP- und HTML-Dateien..."
        New-Item -ItemType Directory -Path $baseTempDir | Out-Null
    } else {
        Log "  Wurde gefunden. Existiert bereits."
    }

    Log "Starte die Verarbeitung aller ZIP-Dateien im Verzeichnis '$($ZipFilesDirectory)'..."

    $zipFiles = Get-ChildItem -Path $ZipFilesDirectory -Filter "*.zip" -File -Recurse
    Log "Gefundene ZIP-Dateien: $($zipFiles.Count)"
    
    $htmlFiles = New-Object System.Collections.Generic.List[object]
    $n = 0

    foreach ($zip in $zipFiles) {
        $n++
        Log "Verarbeite ZIP-Datei #$($n) von $($zipFiles.Count): $($zip.FullName)"
        
        $zipBaseName = [System.IO.Path]::GetFileNameWithoutExtension($zip.FullName)
        $tempDir = Join-Path -Path $baseTempDir -ChildPath $zipBaseName
        Log "Erstelle temporaeres Verzeichnis '$($tempDir)' fuer diese ZIP-Datei..."
        New-Item -ItemType Directory -Path $tempDir | Out-Null

        try {
            Log "Extrahiere ZIP-Datei $($zip.FullName) nach $($tempDir)..."
            [System.IO.Compression.ZipFile]::ExtractToDirectory($zip.FullName, $tempDir)

            Log "  Suche ZIP-Dateien soeben extrahierten Verzeichnis '$($tempDir)'..."
            $SubZipFiles = Get-ChildItem -Path $tempDir -Filter "*.zip" -File
            Log "    Anzahl der gefundene ZIP-Dateien: $($SubZipFiles.Count)"
            Log "    Pfade der gefundenen ZIP-Dateien:"
            foreach ($subZipFileFullName in $SubZipFiles.FullName) {
                Log "    - $($subZipFileFullName)"
            }

            foreach ($subZip in $SubZipFiles) {
                Log "      Durchsuche ZIP-Datei $($subZip.FullName) nach der HTML-Datei..."
                $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($subZip.FullName)
                foreach ($entry in $zipArchive.Entries) {
                    if ($entry.FullName -match '\.html?$') {
                        Log "        HTML-Datei gefunden: $($entry.FullName) ($($entry.Length) Bytes)"

                        $destPath = Join-Path $tempDir $entry.Name
                        Log "          Extrahiere $($entry.FullName) nach: $($destPath)"

                        $entryStream = $entry.Open()
                        $fileStream  = [System.IO.File]::OpenWrite($destPath)
                        $entryStream.CopyTo($fileStream)
                        $fileStream.Close()
                        $entryStream.Close()

                        $htmlFiles.Add($destPath)
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
        Log "Insgesamt $($htmlFiles.Count) HTML-Dateien gefunden und extrahiert."
        Log "Extrahierte HTML-Dateien:"
        foreach ($htmlFile in $htmlFiles) {
            Log " - $htmlFile"
        }
    }
    #endregion

    #region Initialisierung

    Log "Initialisiere Excel..."
    # Excel-Objekt erstellen und Datei oeffnen
    Log "Erstelle Excel-Objekt..."
    $excel = New-Object -ComObject Excel.Application

    if ($Silent) {
        $excel.Visible = $false
    } else {
        $excel.Visible = $true
    }

    # Hole den absoluten Pfad und Dateinamen aus $PathToExcelFile
    Log "Loese Pfad zur Excel-Datei auf: $PathToExcelFile"
    $excelFullPath = Resolve-Path -Path $PathToExcelFile | Select-Object -ExpandProperty Path
    Log "Excel-Dateipfad aufgeloest: $excelFullPath"
    $excelFileName = [System.IO.Path]::GetFileName($excelFullPath)
    Log "Excel-Dateiname extrahiert: $excelFileName"
    Log "oeffne Excel-Datei: $excelFullPath"
    $workbook = $excel.Workbooks.Open($excelFullPath)

    # Create a new instance of the Chrome driver
    Log "Initialisiere Webbrowser..."

    # Baue Splatting-Hashtable für Start-SeChrome
    $chromeSplat = @{
        Quiet = $true
    }
    if ($Silent) {
        $chromeSplat.Headless = $true
        $chromeSplat.Arguments = "--window-size=1920,1080"
    } else {
        $chromeSplat.Arguments = @('start-maximized')
    }
    if ($PathToWebDriverDirectory) {
        $chromeSplat.WebDriverDirectory = $PathToWebDriverDirectory
    }
    if ($PathToChromeBinary) {
        $chromeSplat.BinaryPath = $PathToChromeBinary
    }

    $driver = Start-SeChrome @chromeSplat

    # Timeout fuer FindElements auf 3 Sekunden setzen
    $Seconds = 2
    $driver.Manage().Timeouts().ImplicitWait = [System.TimeSpan]::FromSeconds($Seconds)
    $driver.Manage().Timeouts().PageLoad = [System.TimeSpan]::FromSeconds($Seconds)
    $driver.Manage().Timeouts().AsynchronousJavaScript = [System.TimeSpan]::FromSeconds($Seconds)
    $wait = New-Object OpenQA.Selenium.Support.UI.WebDriverWait($driver, [System.TimeSpan]::FromSeconds($Seconds))

    #endregion

    #region Selenium

    foreach ($file in $htmlFiles) {
        Log " "
        Log "################################################"
        Log "Verarbeite Datei: $($file)"
        Log "################################################"
        Log " "
        $filePath = $file
        $fileName = [System.IO.Path]::GetFileName($filePath)

        try {
            #endregion

            #region Selenium: Model

            Log "------------------------------------------------"
            Log "Lese Modell aus..."
            $driver.Navigate().GoToUrl($filePath)
            # Warte darauf, dass die Webseite fertig geladen ist
            Wait-ForElement -Driver $driver -Text "Model" -XPath "//th[normalize-space(text())='Model' and @scope='row']"
            # Finde das <th>-Element mit Text "Model" und scope="row"
            $modelTh = $driver.FindElementByXPath("//th[normalize-space(text())='Model' and @scope='row']")
            # Finde das zugehoerige <td>-Element (direktes Geschwisterelement)
            $modelTd = $modelTh.FindElementByXPath("following-sibling::td[1]")
            # Finde das <span>-Element im <td> und lese den Text aus
            $ServerModel = $modelTd.FindElementByXPath(".//span").Text.Trim()
            Log "--> Modell = $ServerModel"

            #endregion

            #region Selenium: Tag

            Log "------------------------------------------------"
            Log "Lese Tag aus..."
            # Finde das <th>-Element mit Text "Tag" und scope="row"
            $tagTh = $driver.FindElementByXPath("//th[normalize-space(text())='Tag' and @scope='row']")
            # Finde das zugehoerige <td>-Element (direktes Geschwisterelement)
            $tagTd = $tagTh.FindElementByXPath("following-sibling::td[1]")
            # Finde das <span>-Element im <td> und lese den Text aus
            $Tag = $tagTd.FindElementByXPath(".//span").Text.Trim()
            Log "--> Tag = $Tag"

            #endregion

            #region Selenium: Servername

            Log "------------------------------------------------"
            Log "Lese Servername aus..."
            try {
                $driver.Navigate().GoToUrl("$($filePath)#/hardware/system-setup")
                Wait-ForElement -Driver $driver -Text "DNS iDRAC Name" -XPath "//div[@class='key' and normalize-space(text())='DNS iDRAC Name']"
                $DnsIdracName = $driver.FindElementByXPath("//div[@class='key' and normalize-space(text())='DNS iDRAC Name']")
                $ServernameWithIdrac = $DnsIdracName.FindElementByXPath("following-sibling::*[1]").Text.Trim()
                $Servername = $ServernameWithIdrac -replace '-idrac$', ''
                [string]$Servername = $Servername
                if ([string]::IsNullOrWhiteSpace($Servername)) {
                    throw "Servername konnte nicht ausgelesen werden oder ist leer."
                } else {
                    Log "--> Servername: $Servername"
                }
            } catch {
                Write-Error "Fehler beim Auslesen des Servernamens fuer Datei '$fileName': $($_.Exception.Message). Das Skript wird trotzdem weiter ausgefuehrt. Es kommt zu keinen Folgefehlern." -ErrorAction Continue
                $Servername = "Fehler"
            }

            #endregion

            #region Selenium: iDrac-MAC-Adresse
            # iDrac-MAC-Adresse auslesen
            Log "------------------------------------------------"
            Log "Lese iDrac-MAC-Adresse aus..."

            try {
                $driver.Navigate().GoToUrl("$($filePath)#/hardware/systemboard")
                Log "Navigiere zu Systemboard-Seite: $($filePath)#/hardware/systemboard"
                if (-not $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible([OpenQA.Selenium.By]::XPath("//th[normalize-space(text())='MAC Address']")))) {
                    throw "Timeout: Die Spalte 'MAC Address' konnte nicht gefunden werden. Die Webseite wurde wohl nicht vollstaendig geladen. Skript wird fortgesetzt."
                }
                $macAddressHeader = $driver.FindElementByXPath("//th[normalize-space(text())='MAC Address']")
                Log "Header 'MAC Address' gefunden."
                $IdracMacAddress = $macAddressHeader.FindElementByXPath("following-sibling::td[1]").Text.Trim()
                Log "--> iDrac-MAC-Adresse extrahiert: '$IdracMacAddress'"
                if ([string]::IsNullOrWhiteSpace($IdracMacAddress)) {
                throw "Die iDrac-MAC-Adresse konnte nicht ausgelesen werden oder ist leer."
                }
            } catch {
                Write-Error "Fehler beim Auslesen der iDrac-MAC-Adresse fuer Datei '$fileName': $($_.Exception.Message)" -ErrorAction Continue
                $IdracMacAddress = "Fehler"
            }

            #endregion

            #region Selenium: MAC-Adressen

            # Tabelle mit MAC-Adressen auslesen
            Log "------------------------------------------------"
            Log "Lese alle MAC-Adressen aus..."

            # Klicke auf den Navigationspunkt "Ethernet"
            Log "Navigiere zur Ethernet-Seite..."
            $driver.Navigate().GoToUrl("$($filePath)#/hardware/ethernet")

            # Warte auf Ethernet-ueberschrift
            Wait-ForElement -Driver $driver -Text "Ethernet" -XPath "//h2[contains(@class,'ui left header') and contains(normalize-space(text()),'Ethernet')]"

            # Warte auf Firmware-Tabellenspalte
            Wait-ForElement -Driver $driver -Text "Firmware" -XPath "//th[normalize-space(text())='Firmware']"

            # Die erste Tabelle auf der Seite finden
            Log "Suche die erste Tabelle auf der Seite..."
            $table = $driver.FindElementByXPath("(//table)[1]")
            Log "Tabelle gefunden: $($table.Text)"

            # Alle Zeilen (tr) holen
            Log "Extrahiere Zeilen aus der Tabelle..."
            $rows = $table.FindElements([OpenQA.Selenium.By]::TagName("tr"))
            Log "Zeilen gefunden: $($rows.Text)"

            # Die Header-Zeile extrahieren
            Log "Extrahiere Header-Zeile..."
            Log "Versuche, die Header-Zeile zu extrahieren..."
            $headerRow = $rows[0]
            if ($null -eq $headerRow) {
                Log "Header-Zeile (rows[0]) ist $null!"
            } else {
                Log "Header-Zeile gefunden: $($headerRow.Text)"
            }

            $headerElements = $headerRow.FindElements([OpenQA.Selenium.By]::TagName("th"))
            if ($null -eq $headerElements -or $headerElements.Count -eq 0) {
                Log "Keine <th>-Elemente in der Header-Zeile gefunden!"
            } else {
                Log "Gefundene <th>-Elemente: $($headerElements.Count)"
                foreach ($el in $headerElements) {
                Log "Header-Element: $($el.Text)"
                }
            }

            $headers = New-Object System.Collections.Generic.List[object]
            foreach ($el in $headerElements) {
                $headerText = $el.Text
                if ($null -eq $headerText -or $headerText.Trim() -eq "") {
                    throw "Fehler: Eine Header-Zeile konnte nicht korrekt ausgelesen werden (leer oder null). Das Skript wird abgebrochen."
                } else {
                    $headerText = $headerText.Trim()
                }
                Log "Header-Text extrahiert: '$headerText'"
                $headers.Add($headerText)
            }
            Log "Finale Header-Liste: $($headers -join ', ')"

            # Die Datenzeilen extrahieren
            Log "Extrahiere Datenzeilen..."
            $dataRows = $rows | Select-Object -Skip 1
            Log "Datenzeilen gefunden: $($dataRows.Text)"

            # Jede Datenzeile in ein PSCustomObject umwandeln
            Log "Wandle Datenzeilen in PSCustomObject..."

            $Nics = foreach ($row in $dataRows) {
                $cells = $row.FindElements([OpenQA.Selenium.By]::TagName("td"))
                Log "Zellen gefunden: $($cells.Text)"
                if ($cells.Count -eq $headers.Count) {
                    $obj = [PSCustomObject]@{}
                    for ($i = 0; $i -lt $headers.Count; $i++) {
                        $obj | Add-Member -NotePropertyName $headers[$i] -NotePropertyValue $cells[$i].Text.Trim()
                    }
                    $obj
                }
            }

            # Ausgabe pruefen
            Log "Extrahierte Daten aus der Tabelle:"
            if (-not $Silent) {
                $Nics | Format-Table -AutoSize
            }

            # Extrahiere die Port-Geschwindigkeit aus der Spalte "Model" per Regex und schreibe sie in eine neue Spalte "PortSpeed"
            Log "Extrahiere Port-Geschwindigkeit aus der Spalte 'Model' in die neue Spalte 'PortSpeed'..."
            foreach ($nic in $Nics) {
                $model = $nic.Model
                if ($model -match ' \dx(100G) ') {
                    $nic | Add-Member -NotePropertyName "PortSpeed" -NotePropertyValue '100 G QSFP'
                } elseif ($model -match ' \dx(10G|25G) ') {
                    $nic | Add-Member -NotePropertyName "PortSpeed" -NotePropertyValue '10/25 GbE SFP'
                } else {
                    $nic | Add-Member -NotePropertyName "PortSpeed" -NotePropertyValue $null
                }
            }

            # Ausgabe der NICs mit PortSpeed pruefen
            Log "NICs mit PortSpeed:"
            if (-not $Silent) {
                $Nics | Format-Table -AutoSize
            }

            #endregion

            #region Selenium: InfiniBand

            # Pruefe, ob der Menuepunkt "InfiniBand" existiert und navigiere ggf. dorthin
            Log "------------------------------------------------"
            Log "Pruefe, ob der Menuepunkt 'InfiniBand' existiert..."
            try {
                $infinibandMenu = $driver.FindElements([OpenQA.Selenium.By]::XPath("//a[contains(@href,'#/hardware/infiniband') and contains(normalize-space(text()),'InfiniBand')]"))
                if ($infinibandMenu.Count -gt 0) {
                    Log "'InfiniBand'-Menuepunkt gefunden. Navigiere zur InfiniBand-Seite..."
                    $driver.Navigate().GoToUrl("$($filePath)#/hardware/infiniband")

                    # Warte auf InfiniBand
                    Log "Warte auf das Laden der ueberschrift auf der InfiniBand-Seite..."
                    $timeout = 10
                    $infinibandHeaderFound = $false
                    for ($i = 0; $i -lt $timeout; $i++) {
                        try {
                            $infinibandHeader = $driver.FindElementByXPath("//h2[contains(@class,'ui left header') and contains(normalize-space(text()),'InfiniBand')]")
                            if ($infinibandHeader.Displayed -and $infinibandHeader.Text -like "*InfiniBand*") {
                                $infinibandHeaderFound = $true
                                Log "InfiniBand-Header gefunden und sichtbar."
                                break
                            }
                        } catch {
                            Log "InfiniBand-Header noch nicht gefunden: $($_.Exception.Message)"
                            Start-Invoke-Sleep -Seconds 1
                        }
                    }
                    if (-not $infinibandHeaderFound) {
                        throw "Timeout: Das <h2>-Element mit dem Text 'InfiniBand' wurde nicht gefunden oder ist nicht sichtbar."
                    } elseif ($infinibandHeaderFound) {
                        Log "InfiniBand-Header ist sichtbar."
                    }

                    # Warte auf Firmware-Tabellenspalte
                    Log "Warte auf die Firmware-Spalte auf der InfiniBand-Seite..."
                    $timeout = 10
                    $found = $false
                    for ($i = 0; $i -lt $timeout; $i++) {
                        try {
                            $firmwareHeader = $driver.FindElementByXPath("//th[normalize-space(text())='Firmware']")
                            Log "i = $($i)"
                            Log "firmwareHeader: $($firmwareHeader)"
                            Log "firmwareHeader.Displayed = $($firmwareHeader.Displayed)"
                            Log "firmwareHeader.Text = $($firmwareHeader.Text)"
                            if ($firmwareHeader.Displayed -eq $true -and $firmwareHeader.Text -eq "Firmware") {
                                $found = $true
                                Log "found = $($found)"
                                break
                            }
                            Log "found = $($found)"
                            Start-Invoke-Sleep -Seconds 1
                        } catch {
                            # nix
                        }
                    }
                    if (-not $found) { 
                        throw "Element nicht gefunden Seite wurde nicht vollstaendig geladen." 
                    } elseif ($found) { 
                        Log "Found Firmware-Spalte."
                    }


                    # Warte darauf, dass das gewuenschte <th>-Element mit dem Text 'Firmware' angezeigt wird
            
                    # Tabelle auslesen und in PSCustomObject(s) konvertieren
                    Log "Lese die erste gefundene Tabelle aus und konvertiere sie in PSCustomObject..."

                    # Die erste Tabelle auf der Seite finden
                    try {
                        $tableElement = $driver.FindElementByXPath("(//table)[1]")
                    } catch {
                        Write-Error "Keine Tabelle auf der Seite gefunden: $($_.Exception.Message)"
                        throw "Abbruch: Tabelle konnte nicht gefunden werden."
                    }

                    # Header extrahieren
                    try {
                        $headerElements = $tableElement.FindElements([OpenQA.Selenium.By]::XPath(".//thead//th"))
                        if (-not $headerElements -or $headerElements.Count -eq 0) {
                        throw "Keine Header (<th>) in der Tabelle gefunden."
                        }
                        $headers = New-Object System.Collections.Generic.List[object]
                        foreach ($el in $headerElements) {
                        $headerText = $el.Text
                        if ($null -eq $headerText -or $headerText.Trim() -eq "") {
                            Write-Warning "Leerer Header-Text gefunden, wird uebersprungen."
                            continue
                        }
                        $headers.Add($headerText.Trim())
                        }
                        if ($headers.Count -eq 0) {
                        throw "Alle Header-Zellen sind leer. Tabelle ungueltig."
                        }
                    } catch {
                        Write-Error "Fehler beim Extrahieren der Header-Zeile: $($_.Exception.Message)"
                        throw "Abbruch: Header konnten nicht extrahiert werden."
                    }

                    # Datenzeilen extrahieren
                    try {
                        $rowElements = $tableElement.FindElements([OpenQA.Selenium.By]::XPath(".//tbody/tr"))
                        if (-not $rowElements -or $rowElements.Count -eq 0) {
                        throw "Keine Datenzeilen (<tr>) in der Tabelle gefunden."
                        }
                        $TableObjects = New-Object System.Collections.Generic.List[object]
                        foreach ($row in $rowElements) {
                        $cellElements = $row.FindElements([OpenQA.Selenium.By]::TagName("td"))
                        if ($cellElements.Count -ne $headers.Count) {
                            Write-Warning "Zeile uebersprungen: Anzahl der Zellen ($($cellElements.Count)) stimmt nicht mit Header-Anzahl ($($headers.Count)) ueberein."
                            continue
                        }
                        $obj = [PSCustomObject]@{}
                        for ($i = 0; $i -lt $headers.Count; $i++) {
                            $cellValue = $cellElements[$i].Text
                            if ($null -eq $cellValue) { $cellValue = "" }
                            $obj | Add-Member -NotePropertyName $headers[$i] -NotePropertyValue $cellValue.Trim()
                            $obj | Add-Member -NotePropertyName "PortSpeed" -NotePropertyValue 'InfiniBand'
                        }
                        $TableObjects.Add($obj)
                        }
                        if ($TableObjects.Count -eq 0) {
                        throw "Keine gueltigen Datenzeilen extrahiert."
                        }
                    } catch {
                        Write-Error "Fehler beim Extrahieren der Datenzeilen: $($_.Exception.Message)"
                        throw "Abbruch: Datenzeilen konnten nicht extrahiert werden."
                    }

                    # Die Tabelle steht jetzt als Array von PSCustomObject in $TableObjects zur Verfuegung
                    Log "Tabelle erfolgreich als PSCustomObject(s) extrahiert. Anzahl Zeilen: $($TableObjects.Count)"

                    # Fuege die InfiniBand-NICs zu den Ethernet-NICs hinzu
                    $Nics.Add($TableObjects)

                    # Ausgabe pruefen
                    Log "Tabelle inkl. PortSpeed und InfiniBand:"
                    if (-not $Silent) {
                        $Nics | Format-Table -AutoSize
                    }
                } else {
                    Log "'InfiniBand'-Menuepunkt nicht gefunden. Fahre fort..."
                }
            } catch {
                Log "Fehler beim Pruefen/Navigieren zum 'InfiniBand'-Menuepunkt: $($_.Exception.Message)"
            }

            #region Selenium: Physical Disk

            Log "------------------------------------------------"
            Log "Pruefe, ob der Menuepunkt 'Physical Disks' existiert..."
            try {
                # Pruefe, ob der Menuepunkt "Physical Disks" existiert, sonst ueberspringen
                $physicalDisksMenu = $driver.FindElements([OpenQA.Selenium.By]::XPath("//a[contains(@href,'#/hardware/physical-disk') and contains(normalize-space(text()),'Physical Disks')]"))
                if (-not $physicalDisksMenu -or $physicalDisksMenu.Count -eq 0) {
                    Log "'Physical Disks'-Menuepunkt nicht gefunden, ueberspringe diesen Abschnitt."
                    $PhysicalDisks = $null
                } else {
                    Log "'Physical Disks'-Menuepunkt gefunden."
                        
                    # Fuehre den Klick auf den Sidebar-Button nur aus, wenn $physicalDisksMenu nicht klickbar/verfuegbar/sichtbar ist
                    if (-not $physicalDisksMenu -or $physicalDisksMenu.Count -eq 0 -or -not $physicalDisksMenu[0].Displayed) {
                        Log "physicalDisksMenu wird NICHT angezeigt, darum wird die Sidebar nun angeklickt, um sie auszufahren."
                        Log "Suche das Sidebar-Button-Element mit Klasse 'sidebar icon'..."
                        try {
                            $sidebarButton = $driver.FindElementByXPath("//button[contains(@class,'header') and contains(@class,'item') and .//i[contains(@class,'sidebar') and contains(@class,'icon')]]")
                            if ($sidebarButton -and $sidebarButton.Displayed) {
                                Log "Sidebar-Button gefunden. Klicke darauf..."
                                $sidebarButton.Click()
                                Log "Sidebar-Button wurde erfolgreich geklickt."
                            } else {
                                Log "Sidebar-Button nicht sichtbar oder nicht gefunden."
                            }
                        } catch {
                            Log "Fehler beim Finden/Klicken des Sidebar-Buttons: $($_.Exception.Message)"
                        }
                    } else {
                        Log "physicalDisksMenu wird angezeigt, darum wird die Sidebar nicht angeklickt."
                    }

                    Log "Klicke auf das Element 'Physical Disks' im Menue"
                    if ($physicalDisksMenu.Count -gt 0) {
                        $physicalDisksMenu[0].Click()
                        Log "Habe auf den Menuepunkt 'Physical Disks' geklickt."
                    } else {
                        Log "Habe NICHT auf den Menuepunkt 'Physical Disks' geklickt, da $physicalDisksMenu.Count kleiner, als Null ist."
                    }

                    # Warte darauf, dass die ueberschrift "Physical Disks" und die Spalte "SAS Address" angezeigt werden
                    Log "Warte darauf, dass die ueberschrift 'Physical Disks' angezeigt wird..."

                    $XPath = "//h2[contains(@class,'ui left header') and contains(normalize-space(text()),'Physical Disks')]"
                    $Text = "Physical Disks"
                    $TimeoutSeconds = 6
                    $elementFound = $false
                    for ($i = 0; $i -lt $TimeoutSeconds; $i++) {
                        try {
                            Log "i = $($i)"
                            $element = $driver.FindElementByXPath($XPath)
                            Log "element: $($element)"
                            Log "element.Displayed = $($element.Displayed)"
                            Log "element.Text = $($element.Text)"
                            if ($element.Displayed -and $element.Text -eq $Text) {
                                $elementFound = $true
                                Log "Element '$Text' gefunden und sichtbar."
                                break
                            }
                            Log "found = $($elementFound)"
                            Start-Invoke-Sleep -Seconds 1
                        } catch {
                            Log "Element ist aktuell noch nicht da. Warte noch..."
                        }
                    }

                        
                    if (-not $elementFound) {
                        throw "Timeout: Das Element mit dem Text '$Text' und XPath '$XPath' wurde nicht gefunden oder ist nicht sichtbar."
                    } elseif ($elementFound) {
                        Log "Element '$Text' ist sichtbar."
                    }

                    # Log "Warte darauf, dass die Überschrift 'Physical Disks' angezeigt wird..."
                    Wait-ForElement -Driver $driver -Text "Physical Disks" -XPath "//h2[contains(@class,'header') and contains(normalize-space(text()),'Physical Disks')]"
                    # Log "Warte darauf, dass die Spalte 'SAS Address' angezeigt wird..."
                    Wait-ForElement -Driver $driver -Text "SAS Address" -XPath "//th[normalize-space(text())='SAS Address']"

                    Log "Lese Storage-Controller-Tabelle (Physical Disk) aus..."

                    # Suche das <h2>-Element mit Text "Physical Disks"
                    $physicalDisksHeader = $driver.FindElementByXPath("//h2[contains(@class,'header') and contains(normalize-space(text()),'Physical Disks')]")
                    if (-not $physicalDisksHeader) {
                        throw "ueberschrift 'Physical Disks' nicht gefunden."
                    }
                    Log "ueberschrift 'Physical Disks' gefunden."

                    # Gehe von <h2> nach oben zum naechsten <article> und suche darin das <table>
                    $articleElement = $physicalDisksHeader.FindElementByXPath("ancestor::article[1]")
                    if (-not $articleElement) {
                        throw "<article>-Element fuer 'Physical Disks' nicht gefunden."
                    }
                    Log "<article>-Element gefunden."

                    # Suche das <table> innerhalb des <article>
                    $tableElement = $articleElement.FindElement([OpenQA.Selenium.By]::XPath(".//table"))
                    if (-not $tableElement) {
                        throw "<table>-Element fuer 'Physical Disks' nicht gefunden."
                    }
                    Log "<table>-Element fuer 'Physical Disks' gefunden."

                    # Hole alle <tbody>-Elemente der Tabelle
                    $tbodyElements = $tableElement.FindElements([OpenQA.Selenium.By]::TagName("tbody"))

                    $PhysicalDisks = New-Object System.Collections.Generic.List[object]

                    foreach ($tbody in $tbodyElements) {
                        # Versuche den Storage Controller Namen aus dem ersten <tr> mit <div class='ui ribbon label'> zu extrahieren
                        try {
                            $controllerRow = $tbody.FindElement([OpenQA.Selenium.By]::XPath("./tr[td/div[contains(@class,'ribbon label')]]"))
                            $controllerName = $controllerRow.FindElement([OpenQA.Selenium.By]::XPath(".//div[contains(@class,'ribbon label')]")).Text.Trim()
                        } catch {
                            $controllerName = "Unbekannt"
                            Log "Controller-Name konnte nicht extrahiert werden: $($_.Exception.Message)"
                        }

                        # Hole alle Datenzeilen (tr), die KEIN <div class='ui ribbon label'> enthalten
                        try {
                            $dataRows = $tbody.FindElements([OpenQA.Selenium.By]::XPath("./tr[not(td/div[contains(@class,'ribbon label')])]"))
                        } catch {
                            Log "Fehler beim Extrahieren der Datenzeilen: $($_.Exception.Message)"
                            continue
                        }

                        foreach ($row in $dataRows) {
                            try {
                                $cells = $row.FindElements([OpenQA.Selenium.By]::TagName("td"))
                                if ($cells.Count -lt 10) { 
                                    Log "Zeile uebersprungen: Zu wenig Zellen ($($cells.Count))."
                                    continue 
                                }

                                $obj = [PSCustomObject]@{
                                    "Server Name"      = $Servername
                                    "Server Model"     = $ServerModel
                                    "Server Serial"    = $Tag
                                    "Storage Controller" = $controllerName
                                    "Slot"              = $cells[1].Text.Trim()
                                    "Geraet"             = $cells[5].Text.Trim()
                                    "Serial"            = $cells[7].Text.Trim()
                                    "SAS Address"       = $cells[8].Text.Trim()
                                }
                                $PhysicalDisks.Add($obj)
                            } catch {
                                Log "Fehler beim Verarbeiten einer Datenzeile: $($_.Exception.Message)"
                                continue
                            }
                        }
                    }

                    # Ausgabe pruefen
                    Log "PhysicalDisks extrahiert:"
                    if (-not $Silent) {
                        $PhysicalDisks | Format-Table -AutoSize
                    }
                }
            } catch {
                Write-Error "Fehler beim Auslesen der Physical Disks: $($_.Exception.Message)"
                # Skript laeuft weiter, keinen terminierenden Error werfen
            }

            #endregion

            #region Excel

            # Ausgabe in Excel eintragen
            Log "------------------------------------------------"
            Log "Arbeitsblatt 10 auswaehlen..."
            $worksheet = $workbook.Worksheets.Item(10)
            $lastRow = $worksheet.UsedRange.Rows.Count

            Log "Trage Daten in Excel (Arbeitsblatt 'Geraete Interface MAC') ein..."
            $ServerRow = $null
            for ($i = 3; $i -le $lastRow; $i++) {
                $cellValue = $worksheet.Cells.Item($i, 3).Value2
                if ($cellValue -eq $Servername) {
                    Log "Servername '$Servername' gefunden in Zeile $i"
                    $ServerRow = $i
                    break
                }
            }
            if (-not $ServerRow) {
                # Suche die naechste komplett leere Zeile ab Zeile 3
                for ($i = 3; $i -le ($lastRow + 1000); $i++) {
                    $rowIsEmpty = $true
                    for ($col = 1; $col -le $worksheet.UsedRange.Columns.Count; $col++) {
                        if ($worksheet.Cells.Item($i, $col).Value2) {
                            $rowIsEmpty = $false
                            break
                        }
                    }
                    if ($rowIsEmpty) {
                        $ServerRow = $i
                        Log "Finde freie Zeile und trage Servername '$Servername' ein..."
                        $worksheet.Cells.Item($ServerRow, 3).Value2 = $Servername
                        Log "Servername '$Servername' in neue Zeile $ServerRow eingetragen."
                        break
                    }
                }
                if (-not $ServerRow) {
                    throw "Keine freie Zeile gefunden, um den Servernamen '$Servername' einzutragen."
                }
            }

            # Geraet
            Log "Trage Model '$ServerModel' in Zeile $ServerRow, Spalte 4 mit der ueberschrift 'Geraet' ein..."
            $worksheet.Cells.Item($ServerRow, 4).Value2 = $ServerModel

            # MAC Address > MAC
            Log "Trage iDrac-MAC-Adresse in Zeile $ServerRow, Spalte 6 ein..."
            $worksheet.Cells.Item($ServerRow, 6).Value2 = $IdracMacAddress

            # MAC Address > User / PW
            Log "Trage Benutzer und -Passwort in Zeile $ServerRow, Spalte 7 ein..."
            $worksheet.Cells.Item($ServerRow, 7).Value2 = "root/calvin"

            # Firmware > Port 1
            Log "Finde die MAC-Adresse von Embedded 1 Port 1..."
            $FirmwarePort1MacAddress = $Nics.Where({ $_.Location -eq "Embedded 1, Port 1" })."MAC Address"
            Log "Trage MAC-Adresse von Embedded 1 Port 1 in Zeile $ServerRow, Spalte 8 ein..."
            $worksheet.Cells.Item($ServerRow, 8).Value2 = $FirmwarePort1MacAddress

            # Firmware > Port 2
            Log "Finde die MAC-Adresse von Embedded 2 Port 1..."
            $FirmwarePort2MacAddress = $Nics.Where({ $_.Location -eq "Embedded 2, Port 1" })."MAC Address"
            Log "Trage MAC-Adresse von Embedded 2 Port 1 in Zeile $ServerRow, Spalte 9 ein..."
            $worksheet.Cells.Item($ServerRow, 9).Value2 = $FirmwarePort2MacAddress

            # Netzwerkkarten 100 G QSFP
            Log "Bereite 100 G QSFP-Netzwerkkarten vor..."
            $Nics100G = $Nics.Where({ ($_.Location -like "*Slot*") -or ($_.Location -like "*Integrated*") -and ($_.PortSpeed -eq "100 G QSFP") })
            $j = 0
            foreach ($nic in $Nics100G) {
                Log "Trage Name von $($nic.Location) in Zeile $($ServerRow), Spalte $($j) ein..."
                $worksheet.Cells.Item($ServerRow, 10 + $j).Value2 = $nic.Location
                $j += 1
                Log "Trage MAC-Adresse von $($nic.Location) in Zeile $($ServerRow), Spalte $($j) ein..."
                $worksheet.Cells.Item($ServerRow, 10 + $j).Value2 = $nic."MAC Address"
                $j += 1
            }

            # Netzwerkkarten 10/25 GbE SFP
            $Nics1025G = $Nics.Where({ ($_.Location -like "*Slot*") -or ($_.Location -like "*Integrated*") -and ($_.PortSpeed -eq "10/25 GbE SFP") })
            $n = 0
            foreach ($nic in $Nics1025G) {
                Log "Trage Name von $($nic.Location) in Zeile $($ServerRow), Spalte $($n) ein..."
                $worksheet.Cells.Item($ServerRow, 26 + $n).Value2 = $nic.Location
                $n += 1
                Log "Trage MAC-Adresse von $($nic.Location) in Zeile $($ServerRow), Spalte $($n) ein..."
                $worksheet.Cells.Item($ServerRow, 26 + $n).Value2 = $nic."MAC Address"
                $n += 1
            }

            # Netzwerkkarten InfiniBand
            Log "Bereite InfiniBand-Netzwerkkarten vor..."
            $NicsInfiniBand = $Nics.Where({ $_.PortSpeed -eq "InfiniBand" })
            $m = 0
            foreach ($nic in $NicsInfiniBand) {
                Log "Trage Name von $($nic.Location) in Zeile $($ServerRow), Spalte $($m) ein..."
                $worksheet.Cells.Item($ServerRow, 42 + $m).Value2 = $nic.Location
                $m += 1
                Log "Trage MAC-Adresse von $($nic.Location) in Zeile $($ServerRow), Spalte $($m) ein..."
                $worksheet.Cells.Item($ServerRow, 42 + $m).Value2 = $nic."MAC Address"
                $m += 1
            }

            # Physical Disks / HW&Disk Serial Nr.
            Log "------------------------------------------------"
            Log "Trage Daten in Excel (Arbeitsblatt 'HW&Disk Serial Nr.') ein..."
            $worksheet = $workbook.Worksheets.Item(9)
            $lastRow = $worksheet.UsedRange.Rows.Count
            # Setze Hintergrundfarbe der Zellen (Spalten 3 bis 7 in der neuen Zeile) auf Grau
            for ($col = 1; $col -le 10; $col++) {
                $cell = $worksheet.Cells.Item($lastRow + 1, $col)
                $cell.Interior.ColorIndex = 15  # 15 = Grau (Excel Standard)
            }
            # Server-Eckdaten
            $worksheet.Cells.Item($lastRow + 2,3).Value2 = $Servername
            $worksheet.Cells.Item($lastRow + 2,3).Font.Bold = $true
            $worksheet.Cells.Item($lastRow + 2,7).Value2 = $Tag
            $worksheet.Cells.Item($lastRow + 2,7).Font.Bold = $true

            if ($null -ne $PhysicalDisks -and $PhysicalDisks.Count -gt 0) {
                # Disks
                $i = 0
                foreach ($pd in $PhysicalDisks) {
                    $i++

                    # Spalte: Geraet
                    $worksheet.Cells.Item($lastRow + 2 + $i,4).Value2 = $pd."Geraet"
                    
                    # Spalte: Disk-Fach/Slot
                    $DiskType = $pd."Storage Controller".Split('')[0]
                    if ($DiskType -eq "Backplane") {
                        $DiskType = "Slot"
                    }

                    # Fuehrende Null hinzufuegen
                    if ([int]$pd.Slot -lt 10) {
                        $Slot = "{0:D2}" -f [int]$pd.Slot
                    } else {
                        $Slot = $pd.Slot
                    }

                    $worksheet.Cells.Item($lastRow + 2 + $i,5).Value2 = "$($DiskType) $($Slot)"

                    # Spalte: Disk Type
                    $worksheet.Cells.Item($lastRow + 2 + $i,6).Value2 = $pd."Geraet"

                    # Spalte: Seriennummer
                    $worksheet.Cells.Item($lastRow + 2 + $i,7).Value2 = $pd.Serial

                    # Spalte: SAS Address
                    $worksheet.Cells.Item($lastRow + 2 + $i,8).Value2 = $pd."SAS Address"
                }
            }
        } catch {
            Write-Error "  Fehler beim Verarbeiten der Datei '$($fileName)': $($_.Exception.Message)"
            break
        }
    }
    #endregion

} finally {
    #region Aufraeumen

    # Browser schließen (optional)
    if ($driver) {
        Log "Schließe Browser..."
        $driver.Quit()
    }
    
    # Excel schließen
    if ($workbook) {
        # Speichern der Excel-Datei
        Log "Speichere Excel-Datei..."
        $workbook.Save()
        
        Log "Schließe Workbook..."
        $workbook.Close()
    }
    if ($excel) {
        Log "Schließe Excel..."
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    # Loesche die temporaer angelegten Ordner inkl. Dateien darin
    Log "Bereinige temporaere Dateien und Ordner..."
    try {
        Remove-Item -Path $baseTempDir -Recurse -ErrorAction Stop
        Log "Temporaerer Ordner geloescht: $baseTempDir"
    } catch {
        Write-Warning "Konnte temporaeren Ordner '$baseTempDir' nicht loeschen: $($_.Exception.Message)"
    }

    if (-not $Silent) {
        Stop-Transcript
    }

    #endregion

    Log "Skriptausfuehrung abgeschlossen."
}
