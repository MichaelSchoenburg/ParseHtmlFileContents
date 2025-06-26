# ParseHtmlFileContents

## Übersicht

Dieses PowerShell-Skript automatisiert das Extrahieren und Aggregieren von MAC-Adressen sowie weiterer Netzwerkdaten aus HTML-Dateien, die sich in ZIP-Archiven befinden. Die HTML-Dateien werden entpackt, mit Selenium im Edge-Browser analysiert und die gewünschten Daten anschließend in eine bestehende Excel-Datei eingetragen. Zusätzlich können alle MAC-Adressen in eine CSV-Datei exportiert werden.

---

## Was macht das Skript?

- **Durchsucht ZIP-Archive** nach HTML-Dateien (z.B. `viewer.html`)
- **Extrahiert Netzwerkdaten** (z.B. MAC-Adressen, Port-Informationen) aus Tabellen in den HTML-Dateien
- **Schreibt die Daten** in eine vorhandene Excel-Datei an die passende oder nächste freie Stelle
- **Automatisiert den Browser** (Edge) mit Selenium, um auch dynamisch generierte Inhalte auszulesen

---

## Nutzung

1. **Skript herunterladen**  
   Lade die Datei `Skript.ps1` aus diesem Repository herunter und speichere sie auf deinem Rechner.

2. **Skript ausführen**  
   Öffne eine PowerShell-Konsole und führe das Skript mit den gewünschten Parametern aus.

   Beispielaufruf:
   ```powershell
   .\Skript.ps1 -ZipFilesDirectory "C:\Pfad\zu\deinen\ZIPs" `
                -PathToExcelFile "C:\Pfad\zur\Exceldatei.xlsx" `
                -Verbose
   ```

---

## Parameter

| Parameter            | Beschreibung                                                                                                  | Pflicht? | Standardwert                      |
|----------------------|---------------------------------------------------------------------------------------------------------------|----------|-----------------------------------|
| `-ZipFilesDirectory` | Pfad zum Ordner mit den ZIP-Dateien, in denen sich die HTML-Dateien befinden.                                 | Nein     | Skriptverzeichnis                 |
| `-PathToExcelFile`   | Pfad zur bestehenden Excel-Datei, in die die Daten eingetragen werden sollen.                                 | Ja       | -                                 |

**Hinweise:**
- Die Excel-Datei muss bereits existieren.
- Die ZIP-Dateien können auch verschachtelt sein (ZIP-in-ZIP wird unterstützt).

---

## Beispielaufruf

```powershell
.\Skript.ps1 -ZipFilesDirectory "C:\Daten\Zips" -PathToExcelFile "C:\Daten\Netzwerkdaten.xlsx" -Verbose
```

---

## Ausführliche Ausgaben

- **`-Verbose`**: Zeigt detaillierte Statusmeldungen während der Ausführung an (z.B. welche Datei gerade verarbeitet wird).
- **`-Debug`**: Zeigt noch detailliertere Informationen, z.B. den Inhalt von Variablen und Zwischenschritten.

Beispiel für Debug-Ausgabe:
```powershell
.\Skript.ps1 -ZipFilesDirectory "C:\Daten\Zips" -PathToExcelFile "C:\Daten\Netzwerkdaten.xlsx" -Verbose -Debug
```

---

## Anforderungen

- **PowerShell 5.1**
- **Windows Betriebssystem**
- **Microsoft Edge Browser**
- **Microsoft Edge WebDriver** (`msedgedriver.exe`)  
  - Muss im System-PATH liegen oder der Pfad muss beim Start von Selenium angegeben werden
- **Selenium PowerShell-Modul**  
  - Wird beim ersten Start automatisch installiert, falls nicht vorhanden

---

## Fehler & Support

- Bei Änderungswünschen oder Fehlern bitte ein [Issue](https://github.com/MichaelSchoenburg/ParseHtmlFileContents/issues) im Repository eröffnen.
- Für Fragen oder Hilfestellung gerne ebenfalls ein Issue erstellen.

---

## Lizenz

Dieses Projekt steht unter der MIT-Lizenz. Siehe [LICENSE](LICENSE) für Details.

---

## Autor

Michael Schönburg  
Stand:
