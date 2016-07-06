#
# convert.R - Convert InBehandlung table output into DATEV format
#
# Frank Blankenburg, Jun. 2016
#

library ("XLConnect")

#
# Configuration
#

#input_file  <- "c:/Users/Frank/Documents/Projects/DatevConvert/buchhaltung-export-2016-05.xlsx"
input_file  <- "e:/test/convert/datevconvert/export-2016-05.xlsx"

#output_file <- "c:/Users/Frank/Documents/Projects/DatevConvert/datev-2016-05.csv"
output_file <- "e:/test/convert/datevconvert/datev-2016-05.csv"


#
# Setup
#
datev <- data.frame (
    "Umsatz (ohne Soll/Haben-Kz)" = numeric (0),      # 0
    "Soll/Haben-Kennzeichen" = character (0),         # 1
    "WKZ Umsatz" = double (0),                        # 2
    "Kurs" = double (0),                              # 3
    "Basis-Umsatz" = double (0),                      # 4
    "WKZ Basis-Umsatz" = double (0),                  # 5
    "Konto" = character (0),                          # 6
    "Gegenkonto (ohne BU-Schlüssel)" = character (0), # 7
    "BU-Schlüssel" = character (0),                   # 8
    "Belegdatum" = character (0),                     # 9
    "Belegfeld 1" = character (0),                    # 10
    "Belegfeld 2" = character (0),                    # 11
    "Skonto" = double (0),                            # 12
    "Buchungstext" = character (0),                   # 13
    "Postensperre" = character(0),                    # 14
    "Diverse Adressnummer" = character (0),           # 15
    "Geschäftspartnerbank" = character (0),           # 16
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
    "Funktionsergänzung L+L" = character (0),         # 43
    "BU 49 Hauptfunktionstyp" = character (0),        # 44
    "BU 49 Hauptfunktionsnummer" = character (0),     # 45
    "BU 49 Funktionsergänzung" = character (0),       # 46
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
    "Stück" = integer (0),                            # 87
    "Gewicht" = double (0),                           # 88
    "Zahlweise" = character (0),                      # 89
    "Forderungsart" = character (0),                  # 90
    "Veranlagungsjahr" = character (0),               # 91
    "Zugeordnete Falligkeit" = character (0),         # 92
    "Skontotyp" = character (0),                      # 93
    "Auftragsnummer" = character (0),                 # 94
    "Buchungstyp" = character (0),                    # 95
    "USt-Schlüssel (Anzahlungen)" = double (0),       # 96
    "EU-Land (Anzahlungen)" = character (0),          # 97
    "Sachverhalt L+L (Anzahlungen)" = double (0),     # 98
    "EU-Steuersatz (Anzahlungen)" = double (0),       # 99
    "Erlöskonto (Anzahlungen)" = double (0),          # 100
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

#
# Convert POSIXct date into DATEV string like '0416' for 04-2016
#
convertDate <- function (date) {
	d = as.POSIXlt (date)
	return (sprintf ("%02d%02d", d$mday, d$mon + 1))
}

#
# Skip leading blanks
#
trim <- function (text) {
	sub ("^\\s+", "", text)
}

#
# Fill DATEV frame with the content of one single exported sheet
#
# @param sheet     Imported worksheet
# @param title     Sheet title
# @param account7  Account number used for 7% tax entries
# @param account19 Account number used for 19% tax entries
#
fillCommonFields <- function (sheet, title, account7, account19) {
	for (i in 1:nrow (sheet)) {
		line = sheet[i,]
	
		row = nrow(datev) + 1

		turnover <- abs (line$Gesamtpreis.brutto)

		#
		# Skip lines with 0€ turnover. The import does report an error otherwise, because
		# this cannot be in the world of finance software.
		#
		if (round (turnover, 2) != 0) {

			datev[row,]$'Belegfeld 1'                    <<- line$Rechnungsnummer
			datev[row,]$'Beleginfo - Art 1'              <<- "Art"
			datev[row,]$'Beleginfo - Inhalt 1'           <<- title

			datev[row,]$'Umsatz (ohne Soll/Haben-Kz)'    <<- turnover
			if (line$Gesamtpreis.brutto >= 0) {
				datev[row,]$'Soll/Haben-Kennzeichen'   <<- "H"
			}
			else {
				datev[row,]$'Soll/Haben-Kennzeichen'   <<- "S"
			}

			datev[row,]$'Belegdatum'                     <<- convertDate (line$Rechnungsdatum)
			datev[row,]$'Buchungstext'                   <<- line$Position
			datev[row,]$'Beleginfo - Art 2'              <<- "Rechnungsnummer"
			datev[row,]$'Beleginfo - Inhalt 2'           <<- line$Rechnungsnummer
			datev[row,]$'Beleginfo - Art 3'              <<- "Rechnungsdatum"
			datev[row,]$'Beleginfo - Inhalt 3'           <<- format (line$Rechnungsdatum, "%d.%m.%Y")
			datev[row,]$'Zahlweise'                      <<- line$Zahlungsweise
			datev[row,]$'EU-Steuersatz'                  <<- line$Steuersatz
			datev[row,]$'Gegenkonto (ohne BU-Schlüssel)' <<- 1000

			if (line$Steuersatz == 7.0) {
				datev[row,]$'BU-Schlüssel' <<- 2
				datev[row,]$'Konto'        <<- account7
			}
			else if (line$Steuersatz == 19.0) {
				datev[row,]$'BU-Schlüssel' <<- 3
				datev[row,]$'Konto'        <<- account19
			}

			datev[row,]$'USt-Schlüssel (Anzahlungen)'  <<- 0

			if (!is.na (customer.ids[line$Rechnungsnummer])) {
				datev[row,]$'Beleginfo - Art 4'    <<- "Kundennummer"
				datev[row,]$'Beleginfo - Inhalt 4' <<- customer.ids[line$Rechnungsnummer]
			}
		}
		else {
			print (paste ("WARNING: 0€ turnover at", line$Rechnungsnummer, sep=" "))
		}
	}
}

#
# Import sheet 'Zahlungen' and extract a customer ids per bill number
#
sheet.zahlungen <- readWorksheetFromFile (input_file, sheet=5)

customer.ids <- c ()

for (i in 1:nrow (sheet.zahlungen)) {
	line = sheet.zahlungen[i,]

	if (!is.na (line$Rechnungsnummer) & !is.na (line$Kundennummer)) {
		customer.ids[line$Rechnungsnummer] <- line$Kundennummer
	}
}

#
# Import sheets with relevant data and add them to the DATEV frame
#
sheet.leistungen <- readWorksheetFromFile (input_file, sheet=1)
fillCommonFields (sheet.leistungen, title="Leistungen", account7=8004, account19=8004)

sheet.medikamente.angewendet <- readWorksheetFromFile (input_file, sheet=2)
fillCommonFields (sheet.medikamente.angewendet, title="Medikamente (angewendet)", account7=8011, account19=8014)

sheet.medikamente.abgegeben <- readWorksheetFromFile (input_file, sheet=3)
fillCommonFields (sheet.medikamente.abgegeben, title="Medikamente (abgegeben)", account7=8021, account19=8024)

sheet.produkte <- readWorksheetFromFile (input_file, sheet=4)
fillCommonFields (sheet.produkte, title="Produkte", account7=8031, account19=8034)

#
# Write everything into the output file
# 
# The turnover is formatted into a rounded string before.
#
output <- datev

output$'Umsatz (ohne Soll/Haben-Kz)' <- trim (format (round (output$'Umsatz (ohne Soll/Haben-Kz)', 2), nsmall=2, decimal.mark=","))

handle <- file (output_file, encoding="latin1")
write.table (output, file=handle, row.names=FALSE, quote=FALSE, na="", sep=";", dec=",", qmethod=c("escape", "double"))

#
# Print some statistics
#
soll <- 0
haben <- 0

soll  <- sum (datev[datev$'Soll/Haben-Kennzeichen' == 'S',]$'Umsatz (ohne Soll/Haben-Kz)')
haben <- sum (datev[datev$'Soll/Haben-Kennzeichen' == 'H',]$'Umsatz (ohne Soll/Haben-Kz)')

print ("Summary")
print ("-------------------------------")
print (paste ("Soll  :", soll, sep=" "))
print (paste ("Haben :", haben, sep=" "))
