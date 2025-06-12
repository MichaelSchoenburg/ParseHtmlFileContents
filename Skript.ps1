# Dieses Skript demonstriert, wie man mit PowerShell Inhalte aus der Spalte "MAC-Adresse"
# mehrerer HTML-Tabellen ausliest, aggregiert und in einer CSV-Datei speichert.

# WICHTIGER HINWEIS:
# Die Verwendung von ComObject "HTMLFile" funktioniert nur auf Windows-Systemen,
# da sie die Internet Explorer-Engine nutzt. Für Cross-Plattform-Lösungen oder
# modernere Web-Parsing-Anforderungen müssten Sie möglicherweise ein
# .NET-basiertes HTML-Parsing-Framework (z.B. HTML Agility Pack) verwenden
# oder die HTML-Daten mit regulären Ausdrücken bearbeiten (was komplexer ist).

# --- KONFIGURATION ---
# Definieren Sie das Verzeichnis, in dem sich Ihre HTML-Dateien befinden.
# Sie können hier einen absoluten Pfad angeben oder einen relativen Pfad zum Skriptverzeichnis.
$htmlFilesDirectory = Join-Path $PSScriptRoot "HTML-Dateien"

# Der Name der Spalte, deren Inhalt ausgelesen werden soll.
$columnNameToSelect = "MAC-Adresse"

# Der Pfad zur Ausgabedatei (CSV), in der die aggregierten MAC-Adressen gespeichert werden.
$outputCsvPath = $PSScriptRoot "aggregierte_mac_adressen.csv"

# --- VORBEREITUNG ---
# Überprüfen, ob das angegebene Verzeichnis existiert.
if (-not (Test-Path $htmlFilesDirectory -PathType Container)) {
    Write-Error "Das Verzeichnis '$htmlFilesDirectory' wurde nicht gefunden. Bitte stellen Sie sicher, dass der Pfad korrekt ist."
    exit
}

# Initialisieren einer Liste zum Speichern aller extrahierten MAC-Adressen als benutzerdefinierte Objekte.
# Dies ist wichtig, damit die CSV-Datei eine korrekte Spaltenüberschrift hat.
[System.Collections.ArrayList]$allMacAddresses = @()

# --- VERARBEITUNG DER HTML-DATEIEN ---
Write-Host "Starte die Verarbeitung der HTML-Dateien im Verzeichnis '$htmlFilesDirectory'..."

# Durchsuchen aller .html-Dateien im angegebenen Verzeichnis.
$htmlFiles = Get-ChildItem -Path $htmlFilesDirectory -Filter "*.html" -File

if ($htmlFiles.Count -eq 0) {
    Write-Warning "Keine HTML-Dateien (*.html) im Verzeichnis '$htmlFilesDirectory' gefunden. Es gibt nichts zu verarbeiten."
    exit
}

foreach ($file in $htmlFiles) {
    Write-Host "`nVerarbeite Datei: $($file.Name)"
    $filePath = $file.FullName

    try {
        # Schritt 1: Das HTML-Dokument laden
        $htmlContent = Get-Content -Path $filePath -Raw

        # Erstellen eines HTML-Dokumentobjekts
        $htmlDoc = New-Object -ComObject "HTMLFile"
        # HTML-Inhalt korrekt laden (funktioniert nur in PS 5.1 oder höher)
        $htmlDoc.IHTMLDocument2_write($htmlContent)

        # Schritt 2: Die Tabelle im Dokument finden
        # Wir suchen die erste Tabelle im Dokument.
        $table = $htmlDoc.getElementsByTagName("table") | Select-Object -First 1

        if ($table -eq $null) {
            Write-Warning "  Warnung: Keine Tabelle in Datei '$($file.Name)' gefunden. Überspringe diese Datei."
            continue # Springe zur nächsten Datei
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
            Write-Warning "  Warnung: Spalte '$columnNameToSelect' nicht in Datei '$($file.Name)' gefunden. Überspringe diese Datei."
            continue # Springe zur nächsten Datei
        }

        Write-Host "  Suche nach '$columnNameToSelect' (Index $columnIndex)..."

        # Schritt 4: Inhalte aus der ausgewählten Spalte auslesen
        # Nur die Zeilen des <tbody> betrachten, falls vorhanden, sonst alle tr-Elemente nach der Kopfzeile.
        $dataRows = @()
        $tbody = $table.getElementsByTagName("tbody") | Select-Object -First 1
        if ($tbody -ne $null) {
            $dataRows = $tbody.getElementsByTagName("tr")
        } else {
            # Wenn kein tbody gefunden wird, nehmen wir an, die erste tr ist die Kopfzeile
            $allRows = $table.getElementsByTagName("tr")
            if ($allRows.Count -gt 1) {
                $dataRows = $allRows | Select-Object -Skip 1
            }
        }

        if ($dataRows.Count -eq 0) {
            Write-Host "  Keine Datenzeilen in der Tabelle von '$($file.Name)' gefunden."
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
    }
    catch {
        Write-Error "  Fehler beim Verarbeiten der Datei '$($file.Name)': $($_.Exception.Message)"
        # Fährt fort mit der nächsten Datei, auch wenn ein Fehler auftritt
    }
}

# --- EXPORT ZUR CSV-DATEI ---
if ($allMacAddresses.Count -gt 0) {
    Write-Host "`nAlle MAC-Adressen wurden gesammelt. Exportiere nach '$outputCsvPath'..."
    $allMacAddresses | Export-Csv -Path $outputCsvPath -NoTypeInformation -Encoding UTF8

    Write-Host "`nErfolgreich aggregierte MAC-Adressen in '$outputCsvPath' gespeichert."
    Write-Host "Anzahl der gefundenen Einträge: $($allMacAddresses.Count)"
} else {
    Write-Warning "`nKeine MAC-Adressen in den verarbeiteten Dateien gefunden. Es wurde keine CSV-Datei erstellt."
}

Write-Host "`nSkriptausführung abgeschlossen."
