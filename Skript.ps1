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

#region Parameter und Vorbereitung
param(
    [Parameter(Mandatory = $false)]
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
    [string]$htmlFilesDirectory = (Join-Path $PSScriptRoot "HTML-Dateien"),

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$columnNameToSelect = "MAC-Adresse",

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        if (-not ($_ -is [string] -and $_.Trim().Length -gt 0)) {
            throw "Der Dateiname darf nicht leer sein."
        }
        if ($_ -notmatch '^[^\\/:*?"<>|]+?\.[a-zA-Z0-9]+$') {
            throw "Der Ausgabepfad muss einen gültigen Dateinamen mit Dateiendung enthalten (z.B. 'datei.csv')."
        }
        return $true
    })]
    [string]$outputCsvPath = (Join-Path $PSScriptRoot "aggregierte_mac_adressen.csv")
)

# Überprüfen, ob das angegebene Verzeichnis existiert.
if (-not (Test-Path $htmlFilesDirectory -PathType Container)) {
    Write-Error "Das Verzeichnis '$htmlFilesDirectory' wurde nicht gefunden. Bitte stellen Sie sicher, dass der Pfad korrekt ist."
    exit
}

# Initialisieren einer Liste zum Speichern aller extrahierten MAC-Adressen als benutzerdefinierte Objekte.
[System.Collections.ArrayList]$allMacAddresses = @()
#endregion

#region ZIP-Dateien durchsuchen und HTML extrahieren
Write-Host "Starte die Verarbeitung der HTML-Dateien im Verzeichnis '$htmlFilesDirectory'..."

$zipFiles  = Get-ChildItem -Path $htmlFilesDirectory -Filter "*.zip" -File -Recurse
$htmlFiles = @()

foreach ($zip in $zipFiles) {
    Write-Host "Durchsuche ZIP-Datei: $($zip.FullName)"
    try {
        $zipPath    = $zip.FullName
        $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
        foreach ($entry in $zipArchive.Entries) {
            if ($entry.FullName -match '\.html?$') {
                Write-Host "  - HTML-Datei gefunden: $($entry.FullName) ($($entry.Length) Bytes)"

                # Erzeuge ein eindeutiges temporäres Verzeichnis für die Extraktion
                $baseTempDir = Join-Path -Path $env:TEMP -ChildPath "_ParseHtmlFileContentsSkript"
                if (-not (Test-Path $baseTempDir)) {
                    New-Item -ItemType Directory -Path $baseTempDir | Out-Null
                }
                $tempDir = Join-Path -Path $baseTempDir -ChildPath ([System.IO.Path]::GetRandomFileName())
                New-Item -ItemType Directory -Path $tempDir | Out-Null
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
    catch {
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

#region HTML-Dateien verarbeiten und MAC-Adressen extrahieren
foreach ($file in $htmlFiles) {
    Write-Host "Verarbeite Datei: $($file)"
    $filePath = $file
    $fileName = [System.IO.Path]::GetFileName($filePath)

    try {
        # Schritt 1: Das HTML-Dokument laden
        $htmlContent = Get-Content -Path $filePath -Raw

        # Erstellen eines HTML-Dokumentobjekts
        $htmlDoc     = New-Object -ComObject "HTMLFile"

        # HTML-Inhalt korrekt laden (funktioniert nur in PS 5.1 oder höher)
        $htmlDoc.IHTMLDocument2_write($htmlContent)

        # Schritt 2: Die Tabelle im Dokument finden
        # Wir suchen die erste Tabelle im Dokument.
        $table = $htmlDoc.getElementsByTagName("table") | Select-Object -First 1

        if ($null -eq $table) {
            Write-Warning "  Warnung: Keine Tabelle in Datei '$($fileName)' gefunden. Überspringe diese Datei."
            continue
        }

        # Schritt 3: Den Index der gewünschten Spalte bestimmen
        $columnIndex = -1 # Standardwert, falls Spalte nicht gefunden wird

        # Iteriere durch die Header-Zellen, um den Index zu finden
        $headerCells = $table.getElementsByTagName("th")
        for ($i = 0; $i -lt $headerCells.length; $i++) {
            if ($headerCells.item($i).innerText.Trim() -eq $columnNameToSelect) {
                $columnIndex = $i
                break
            }
        }

        if ($columnIndex -eq -1) {
            Write-Warning "  Warnung: Spalte '$columnNameToSelect' nicht in Datei '$($fileName)' gefunden. Überspringe diese Datei."
            continue
        }

        Write-Host "  Suche nach '$columnNameToSelect' (Index $columnIndex)..."

        # Schritt 4: Inhalte aus der ausgewählten Spalte auslesen
        # Nur die Zeilen des <tbody> betrachten, falls vorhanden, sonst alle tr-Elemente nach der Kopfzeile.
        $dataRows = @()
        $tbody    = $table.getElementsByTagName("tbody") | Select-Object -First 1
        if ($null -ne $tbody) {
            $dataRows = $tbody.getElementsByTagName("tr")
        } else {
            # Wenn kein tbody gefunden wird, nehmen wir an, die erste tr ist die Kopfzeile
            $allRows = $table.getElementsByTagName("tr")
            if ($allRows.Count -gt 1) {
                $dataRows = $allRows | Select-Object -Skip 1
            }
        }

        if ($dataRows.Count -eq 0) {
            Write-Host "  Keine Datenzeilen in der Tabelle von '$($fileName)' gefunden."
            continue
        }

        foreach ($row in $dataRows) {
            $cells = $row.getElementsByTagName("td")
            if ($columnIndex -lt $cells.length) {
                $cellContent = $cells.item($columnIndex).innerText.Trim()
                if (-not [string]::IsNullOrWhiteSpace($cellContent)) {
                    # Füge die extrahierte MAC-Adresse als benutzerdefiniertes Objekt zur Liste hinzu
                    $allMacAddresses.Add([PSCustomObject]@{ 'MAC-Adresse' = $cellContent })
                    Write-Host "    - Gefunden: $cellContent"
                }
            }
        }
    } catch {
        Write-Error "  Fehler beim Verarbeiten der Datei '$($fileName)': $($_.Exception.Message)"
        # Fährt fort mit der nächsten Datei, auch wenn ein Fehler auftritt
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
