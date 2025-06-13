# ParseHtmlFileContents

## Übersicht

Dieses PowerShell-Skript aggregiert die Inhalte aus der Spalte **"MAC-Adresse"** mehrerer HTML-Tabellen, die sich in verschiedenen ZIP-Dateien befinden, und speichert alle gefundenen MAC-Adressen in einer CSV-Datei.

## Nutzung

1. **Repository klonen**  
   ```powershell
   git clone https://github.com/deinbenutzername/ParseHtmlFileContents.git
   ```

   **Oder:**  
   - Das Release (`Skript.ps1`) herunterladen  
   - Oder `Skript.ps1` direkt aus dem Repository herunterladen oder per Copy & Paste speichern

2. **Skript ausführen**  
   ```powershell
   .\Skript.ps1 -ZipOrdner "Pfad\zu\deinen\ZIPs" -CsvDatei "output.csv" -columnNameToSelect "MAC-Adresse"
   ```
   - `-ZipOrdner`: Ordnerpfad, in dem sich die ZIP-Dateien befinden
   - `-CsvDatei`: Pfad zur Ausgabedatei (CSV)
   - `-columnNameToSelect`: Der Name der Tabellenspalten, deren Werte extrahiert werden sollen (Spaltenname muss in allen Tabellen gleich sein)

## Anforderungen

- **PowerShell 5.1** (PowerShell 7 wird nicht unterstützt)
- **Windows Betriebssystem**

## Beispiel

Nach der Ausführung enthält die CSV-Datei alle MAC-Adressen aus den HTML-Tabellen der ZIP-Dateien, jeweils in einer eigenen Zeile.
