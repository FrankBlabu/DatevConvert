#
# convert.R - Convert InBehandlung table output into DATEV format
#
# Frank Blankenburg, Jun. 2016
#

library ("XLConnect")

datev <- data.frame (
    "Umsatz (ohne Soll/Haben-Kz)" = double (0),       # 0
    "Soll/Haben-Kennzeichen" = character (0),         # 1
    "WKZ Umsatz" = double (0),                        # 2
    "Kurs" = double (0),                              # 3
    "Basis-Umsatz" = double (0),                      # 4
    "WKZ Basis-Umsatz" = double (0),                  # 5
    "Konto" = character (0),                          # 6
    "Gegenkonto (ohne BU-Schl�ssel)" = character (0), # 7
    "BU-Schl�ssel" = character (0),                   # 8
    "Belegdatum" = character (0),                     # 9
    "Belegfeld 1" = character (0),                    # 10
    "Belegfeld 2" = character (0),                    # 11
    "Skonto" = double (0),                            # 12
    "Buchungstext" = character (0),                   # 13
    "Postensperre" = character(0),                    # 14
    "Diverse Adressnummer" = character (0),           # 15
    "Gesch�ftspartnerbank" = character (0),           # 16
    "Sachverhalt" = character (0),                    # 17
    "Zinssperre" = character (0),                     # 18
    "Beleglink" = character (0),                      # 19
    "Beleginfo - Art 1" = character (0),              # 20
    "Beleginfo - Inhalt 1" = character (0),           # 21
    "Beleginfo - Art 2" = character (0),              # 22
    "Beleginfo - Inhalt 2" = character (0),           # 23
    "Beleginfo - Art 3" = character (0),              # 24
    "Beleginfo - Inhalt 3" = character (0),           # 25
    "Beleginfo - Art 4" = character (0),              # 26
    "Beleginfo - Inhalt 4" = character (0),           # 27
    "Beleginfo - Art 5" = character (0),              # 28
    "Beleginfo - Inhalt 5" = character (0),           # 29
    "Beleginfo - Art 6" = character (0),              # 30
    "Beleginfo - Inhalt 6" = character (0),           # 31
    "Beleginfo - Art 7" = character (0),              # 32
    "Beleginfo - Inhalt 7" = character (0),           # 33
    "Beleginfo - Art 8" = character (0),              # 34
    "Beleginfo - Inhalt 8" = character (0),           # 35
    "KOST1 - Kostenstelle" = character (0),           # 36
    "KOST2 - Kostenstelle" = character (0),           # 37
    "Kost-Menge" = double (0),                        # 38
    "EU-Land u. UStID" = character (0),               # 39
    "EU-Steuersatz" = double (0),                     # 40
    "Abw. Versteuerungsart" = character (0),          # 41
    "Sachverhalt L+L" = character (0),                # 42
    "Funktionserg�nzung L+L" = character (0),         # 43
    "BU 49 Hauptfunktionstyp" = character (0),        # 44
    "BU 49 Hauptfunktionsnummer" = character (0),     # 45
    "BU 49 Funktionserg�nzung" = character (0),       # 46
    "Zusatzinformation - Art 1" = character (0),      # 47
    "Zusatzinformation- Inhalt 1" = character (0),    # 48
    "Zusatzinformation - Art 2" = character (0),      # 49
    "Zusatzinformation- Inhalt 2" = character (0),    # 50
    "Zusatzinformation - Art 3" = character (0),      # 51
    "Zusatzinformation- Inhalt 3" = character (0),    # 52
    "Zusatzinformation - Art 4" = character (0),      # 53
    "Zusatzinformation- Inhalt 4" = character (0),    # 54
    "Zusatzinformation - Art 5" = character (0),      # 55
    "Zusatzinformation- Inhalt 5" = character (0),    # 56
    "Zusatzinformation - Art 6" = character (0),      # 57
    "Zusatzinformation- Inhalt 6" = character (0),    # 58
    "Zusatzinformation - Art 7" = character (0),      # 59
    "Zusatzinformation- Inhalt 7" = character (0),    # 60
    "Zusatzinformation - Art 8" = character (0),      # 61
    "Zusatzinformation- Inhalt 8" = character (0),    # 62
    "Zusatzinformation - Art 9" = character (0),      # 63
    "Zusatzinformation- Inhalt 9" = character (0),    # 64
    "Zusatzinformation - Art 10" = character (0),     # 65
    "Zusatzinformation- Inhalt 10" = character (0),   # 66
    "Zusatzinformation - Art 11" = character (0),     # 67
    "Zusatzinformation- Inhalt 11" = character (0),   # 68
    "Zusatzinformation - Art 12" = character (0),     # 69
    "Zusatzinformation- Inhalt 12" = character (0),   # 70
    "Zusatzinformation - Art 13" = character (0),     # 71
    "Zusatzinformation- Inhalt 13" = character (0),   # 72
    "Zusatzinformation - Art 14" = character (0),     # 73
    "Zusatzinformation- Inhalt 14" = character (0),   # 74
    "Zusatzinformation - Art 15" = character (0),     # 75
    "Zusatzinformation- Inhalt 15" = character (0),   # 76
    "Zusatzinformation - Art 16" = character (0),     # 77
    "Zusatzinformation- Inhalt 16" = character (0),   # 78
    "Zusatzinformation - Art 17" = character (0),     # 79
    "Zusatzinformation- Inhalt 17" = character (0),   # 80
    "Zusatzinformation - Art 18" = character (0),     # 81
    "Zusatzinformation- Inhalt 18" = character (0),   # 82
    "Zusatzinformation - Art 19" = character (0),     # 83
    "Zusatzinformation- Inhalt 19" = character (0),   # 84
    "Zusatzinformation - Art 20" = character (0),     # 85
    "Zusatzinformation- Inhalt 20" = character (0),   # 86
    "St�ck" = integer (0),                            # 87
    "Gewicht" = double (0),                           # 88
    "Zahlweise" = character (0),                      # 89
    "Forderungsart" = character (0),                  # 90
    "Veranlagungsjahr" = character (0),               # 91
    "Zugeordnete Falligkeit" = character (0),         # 92
    "Skontotyp" = character (0),                      # 93
    "Auftragsnummer" = character (0),                 # 94
    "Buchungstyp" = character (0),                    # 95
    "USt-Schl�ssel (Anzahlungen)" = double (0),       # 96
    "EU-Land (Anzahlungen)" = character (0),          # 97
    "Sachverhalt L+L (Anzahlungen)" = double (0),     # 98
    "EU-Steuersatz (Anzahlungen)" = double (0),       # 99
    "Erl�skonto (Anzahlungen)" = double (0),          # 100
    "Herkunft-Kz" = double (0),                       # 101
    "Buchungs GUID" = double (0),                     # 102
    "KOST-Datum" = as.POSIXct (character (0)),        # 103
    "SEPA-Mandatsreferenz" = double (0),              # 104
    "Skontosperre" = double (0),                      # 105
    "Gesellschaftername" = double (0),                # 106
    "Beteiligtennummer" = double (0),                 # 107
    "Identifikationsnummer" = double (0),             # 108
    "Zeichnernummer" = double (0),                    # 109
    "Postensperre bis" = double (0),                  # 110
    "Bezeichnung SoBil-Sachverhalt" = double (0),     # 111
    "Kennzeichen SoBil-Buchung" = double (0),         # 112
    "Festschreibung" = double (0),                    # 113
    "Leistungsdatum" = as.POSIXct (character (0)),    # 114
    "Datum Zuord. Steuerperiode" = as.POSIXct (character (0)), # 115
    stringsAsFactors=FALSE, check.names=FALSE)


convertDate <- function (date) {
	d = as.POSIXlt (date)
	return (sprintf ("%02d%02d", d$mon, d$year - 100))
}

fillCommonFields <- function (sheet, title) {
	for (i in 1:nrow (sheet)) {
		line = sheet[i,]
	
		row = nrow(datev) + 1

		datev[row,]$'Beleginfo - Art 1'            <<- "Art"
		datev[row,]$'Beleginfo - Inhalt 1'         <<- title

		datev[row,]$'Umsatz (ohne Soll/Haben-Kz)'  <<- abs (line$Gesamtpreis.brutto)
		if (line$Gesamtpreis.brutto > 0) {
			datev[row,]$'Soll/Haben-Kennzeichen' <<- "H"
		}
-		else {
			datev[row,]$'Soll/Haben-Kennzeichen' <<- "S"
		}

		datev[row,]$'Belegdatum'                   <<- convertDate (line$Rechnungsdatum)
		datev[row,]$'Buchungstext'                 <<- line$Position
		datev[row,]$'Beleginfo - Art 2'            <<- "Rechnungsnummer"
		datev[row,]$'Beleginfo - Inhalt 2'         <<- line$Rechnungsnummer
		datev[row,]$'Zahlweise'                    <<- line$Zahlungsweise
		datev[row,]$'EU-Steuersatz'                <<- line$Steuersatz
	}
}

#
# Import sheets
#

sheet.leistungen <- readWorksheetFromFile ("e:/test/export.xlsx", sheet=1)
fillCommonFields (sheet.leistungen, "Leistungen")

sheet.medikamente.angewendet <- readWorksheetFromFile ("e:/test/export.xlsx", sheet=2)
fillCommonFields (sheet.medikamente.angewendet, "Medikamente (angewendet)")

sheet.medikamente.abgegeben <- readWorksheetFromFile ("e:/test/export.xlsx", sheet=3)
fillCommonFields (sheet.medikamente.abgegeben, "Medikamente (abgegeben)")

sheet.produkte <- readWorksheetFromFile ("e:/test/export.xlsx", sheet=4)
fillCommonFields (sheet.produkte, "Produkte")

#
# Write everything out
#
handle <- file ('e:/test/output.csv', encoding="UTF-8")
write.table (datev, file=handle, row.names=FALSE, na="", sep=",")

#
# Print some control values
#
soll  <- sum (datev[datev$'Soll/Haben-Kennzeichen' == 'S']$'Umsatz (ohne Soll/Haben-Kz)')
haben <- sum (datev[datev$'Soll/Haben-Kennzeichen' == 'H']$'Umsatz (ohne Soll/Haben-Kz)')

print ("Summary")
print ("-------------------------------")
print (paste ("Soll  :", soll, sep=" "))
print (paste ("Haben :", haben, sep=" "))


