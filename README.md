# Verbrauchs-Prognose

Web-App zur Prognose von Verbrauchsdaten basierend auf historischen Excel-Daten.

## Features

- Excel-Import (.xlsx/.xls)
- Automatische Erkennung von Produkten und Untertypen
- Holt-Winters Triple Exponential Smoothing Prognose
- 5-Jahres-Prognose (monatlich, quartalsweise, jährlich)
- Interaktive Charts mit Chart.js
- Anpassbarer Skalierungsfaktor
- Dark/Light Mode
- Excel-Export

## Nutzung

1. Öffne `index.html` im Browser
2. Lade eine Excel-Datei mit Verbrauchsdaten hoch

### Excel-Format

| Produkt | Untertyp | 01/2024 | 02/2024 | ... |
|---------|----------|---------|---------|-----|
| Strom   | Küche    | 150     | 145     | ... |

- Spalten: Produkt, Untertyp, Monatsspalten (MM/YYYY)

## Technologien

- Vanilla JavaScript
- Chart.js – Diagramme
- SheetJS – Excel-Import/Export

## Algorithmus

Holt-Winters Triple Exponential Smoothing mit Fallback auf lineare Regression bei zu wenig Daten.

## Projektstruktur

```
verbrauch_app_html/
├── index.html
├── app.js
├── style.css
└── README.md
```
