# ParseHtmlFileContents
[![CodeFactor](https://www.codefactor.io/repository/github/michaelschoenburg/parsehtmlfilecontents/badge/main)](https://www.codefactor.io/repository/github/michaelschoenburg/parsehtmlfilecontents/overview/main) ![PSScriptAnalyzer](https://github.com/michaelschoenburg/ParseHtmlFileContents/actions/workflows/psscriptanalyzer.yml/badge.svg) ![GitHub last commit (branch)](https://img.shields.io/github/last-commit/michaelschoenburg/parsehtmlfilecontents/main?display_timestamp=author)

## Inhaltsverzeichnis

- [Übersicht](#übersicht)
- [Was macht das Skript?](#was-macht-das-skript)
- [Nutzung](#nutzung)
- [Parameter](#parameter)
- [Erwartete Datenstruktur](#erwartete-datenstruktur)
- [Beispielaufruf](#beispielaufruf)
- [Ausführliche Ausgaben](#ausführliche-ausgaben)
- [Anforderungen](#anforderungen)
- [Fehler & Support](#fehler--support)
- [Lizenz](#lizenz)
- [Autor](#autor)

---

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

   > **Hinweis:**  
   > Falls beim Ausführen des Skripts eine Fehlermeldung bezüglich der Ausführungsrichtlinie (`Execution Policy`) erscheint, kann das Skript mit folgendem Befehl gestartet werden, um die Richtlinie temporär zu umgehen:
   > 
   > ```powershell
   > powershell.exe -ExecutionPolicy Bypass -File .\Skript.ps1 -ZipFilesDirectory "C:\Pfad\zu\deinen\ZIPs" -PathToExcelFile "C:\Pfad\zur\Exceldatei.xlsx"
   > ```

---


| Parameter            | Beschreibung                                                                                                  | Pflicht? | Standardwert                      |
|----------------------|---------------------------------------------------------------------------------------------------------------|----------|-----------------------------------|
| `-ZipFilesDirectory` | Pfad zum Ordner mit den ZIP-Dateien, in denen sich die HTML-Dateien befinden.                                 | Nein     | Skriptverzeichnis                 |
| `-PathToExcelFile`   | Pfad zur bestehenden Excel-Datei, in die die Daten eingetragen werden sollen.                                 | Ja       | -                                 |
| `-Verbose`           | Gibt während der Ausführung detaillierte Statusmeldungen aus (z.B. welche Datei gerade verarbeitet wird).     | Nein     | Ausgeschaltet                     |
| `-Debug`             | Zeigt noch detailliertere Informationen und Zwischenschritte für die Fehlersuche an.                          | Nein     | Ausgeschaltet                     |


**Hinweise:**
- Die Excel-Datei muss bereits existieren.
- Es wird erwartet, dass die ZIP-Dateien eine Ebene verschachtelt sind (ZIP-in-ZIP). Siehe auch die folgende Erklärung ("Erwartete Datenstruktur").

---

## Erwartete Datenstruktur

Im angegebenen Ordner muss sich, in beliebiger Tiefe (rekursiv), mindestens eine ZIP-Datei befinden. Innerhalb dieser ZIP-Datei muss sich wiederum, ebenfalls rekursiv, eine weitere ZIP-Datei befinden. In dieser letzten ZIP-Datei muss sich schließlich, wiederum rekursiv, eine HTML-Datei befinden. Diese HTML-Datei wird für die weitere Verarbeitung genutzt.

Die Verarbeitungslogik sucht also rekursiv nach ZIP-Dateien, entpackt diese und sucht weiter, bis eine HTML-Datei gefunden wird.

### Veranschaulichung der Datenstruktur

```text
Wurzelordner
└── (beliebige Unterordner)
   └── ErsteEbene.zip
      └── (beliebige Unterordner in ZIP)
         └── ZweiteEbene.zip
            └── (beliebige Unterordner in ZIP)
               └── Datei.html
```

- **Wurzelordner**: Startpunkt der Suche
- **ZIP-in-ZIP**: ZIP-Dateien können beliebig verschachtelt sein
- **HTML-Datei**: Ziel der Suche, wird extrahiert und verarbeitet

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
