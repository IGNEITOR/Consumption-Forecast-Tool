// ================================
// Translations
// ================================
const translations = {
    de: {
        upload: {
            description: "Excel/CSV/LibreOffice/OpenOffice (ODS)-Datei mit Produkten, Untertypen und monatlichen Verbrauchsdaten hochladen",
            selectBtn: "📁 Datei auswählen",
            noFile: "Keine Datei gewählt"
        },
        app: {
            changeFile: "📁 Datei wechseln",
            product: "Produkt:",
            first6Months: "Erste 6 Monate",
            last6Months: "Letzte 6 Monate",
            month: "Monat",
            consumption: "Verbrauch",
            factor: "Faktor für",
            calculateForecast: "Prognose berechnen",
            noSubproducts: "Keine Unterprodukte gefunden",
            noData: "Keine Daten vorhanden"
        },
        forecast: {
            title: "Prognose",
            tablesTitle: "Prognose-Tabellen",
            year1Monthly: "Jahr 1 (monatlich)",
            year2Quarterly: "Jahr 2 (quartalsweise)",
            year3Quarterly: "Jahr 3 (quartalsweise)",
            year4Yearly: "Jahr 4-5 (jährlich)",
            changeTitle: "📈 Veränderung zum Vorjahr",
            year: "Jahr",
            change: "Veränderung %",
            value: "Prognose"
        },
        export: {
            title: "💾 Export",
            downloadBtn: "📥 Excel herunterladen",
            noResults: "Noch keine Ergebnisse berechnet",
            headers: {
                product: "Produkt",
                subtype: "Untertyp",
                factor: "Faktor",
                year1Monthly: "Jahr 1 (monatlich)",
                year2Quarter: "Jahr 2 (Quartal)",
                year3Quarter: "Jahr 3 (Quartal)",
                year4: "Jahr 4",
                year5: "Jahr 5",
                change1: "Veränderung Jahr 1 %",
                change2: "Veränderung Jahr 2 %",
                change3: "Veränderung Jahr 3 %",
                change4: "Veränderung Jahr 4 %",
                change5: "Veränderung Jahr 5 %"
            }
        },
        errors: {
            minColumns: "Excel muss mindestens 3 Spalten haben",
            noDateCols: "Keine gültigen Monatsspalten gefunden (erwartet: MM/YYYY)",
            noData: "Keine gültigen Daten gefunden"
        },
        infoBar: {
            records: "Datensätze",
            period: "Zeitraum"
        }
    },
    en: {
        upload: {
            description: "Upload an Excel/CSV/LibreOffice/OpenOffice (ODS) file with products, subtypes and monthly consumption data",
            selectBtn: "📁 Select file",
            noFile: "No file selected"
        },
        app: {
            changeFile: "📁 Change File",
            product: "Product:",
            first6Months: "First 6 Months",
            last6Months: "Last 6 Months",
            month: "Month",
            consumption: "Consumption",
            factor: "Factor for",
            calculateForecast: "Calculate Forecast",
            noSubproducts: "No subproducts found",
            noData: "No data available"
        },
        forecast: {
            title: "Forecast",
            tablesTitle: "Forecast tables",
            year1Monthly: "Year 1 (monthly)",
            year2Quarterly: "Year 2 (quarterly)",
            year3Quarterly: "Year 3 (quarterly)",
            year4Yearly: "Year 4-5 (yearly)",
            changeTitle: "📈 Change from Previous Year",
            year: "Year",
            change: "Change %",
            value: "Forecast"
        },
        export: {
            title: "💾 Export",
            downloadBtn: "📥 Download Excel",
            noResults: "No results calculated yet",
            headers: {
                product: "Product",
                subtype: "Subtype",
                factor: "Factor",
                year1Monthly: "Year 1 (monthly)",
                year2Quarter: "Year 2 (Quarter)",
                year3Quarter: "Year 3 (Quarter)",
                year4: "Year 4",
                year5: "Year 5",
                change1: "Change Year 1 %",
                change2: "Change Year 2 %",
                change3: "Change Year 3 %",
                change4: "Change Year 4 %",
                change5: "Change Year 5 %"
            }
        },
        errors: {
            minColumns: "Excel must have at least 3 columns",
            noDateCols: "No valid month columns found (expected: MM/YYYY)",
            noData: "No valid data found"
        },
        infoBar: {
            records: "records",
            period: "Period"
        }
    }
};
