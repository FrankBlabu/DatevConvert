#
# convert.R - Convert InBehandlung table output into DATEV format
#
# The MIT License (MIT)
#
# Copyright (c) 2016 Frank Blankenburg
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software
# and associated documentation files (the "Software"), to deal in the Software without restriction,
# including without limitation the rights to use, copy, modify, merge, publish, distribute,
# sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all copies or substantial
# portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT
# NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
# IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
# WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
# SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#

library ("XLConnect")

#
# Configuration
#

#input_file  <- "c:/Users/Frank/Documents/Projects/DatevConvert/buchhaltung-export-2016-06.xlsx"
input_file  <- "e:/test/convert/datevconvert/buchhaltung-export-2016-06.xlsx"

#output_file <- "c:/Users/Frank/Documents/Projects/DatevConvert/datev-2016-06.csv"
output_file <- "e:/test/convert/datevconvert/datev-2016-06.csv"

account.cash     <- 1001
account.bank     <- 1360
account.card     <- 1361
account.transfer <- 1362

customer.ids <- c ()
remarks <- c ()
responsible <- c ()
payment.ids <- c ()
payment.dates <- c()

#
# Intermediate frame for listing all entries
#
data <- data.frame (
	"bill.id"          = character (0), # Id of the bill
	"payment.id"       = numeric (0),   # Id of the payment itself
	"payment.date"     = character (0), # Date of payment
	"item.kind"        = character (0), # Kind of item applied/sold
	"item.date"        = character (0), # Date the item was applied/sold
	"item.description" = character (0), # Description of the item
	"item.tax"         = numeric (0),   # Tax of the item
	"customer.id"      = character (0), # Id of the customer
	"amount"           = double (0),    # Amount of money
	"remarks"          = character (0), # Payment remarks
	"responsible"      = character (0), # Name of the responsible person
	"account"          = numeric (0),   # Account where the money goes to / came from
	stringsAsFactors=FALSE, check.names=FALSE)
	
#
# Datev frame
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
# Add counter entry
#
# @param line         Input line
# @param row          DATEV frame row number to add
# @param account.from Where the money comes from
# @param account.to   Where the mones goes to
#
addCounterEntry <- function (line, row, account.from, account.to) {
	if (datev[row,]$'Soll/Haben-Kennzeichen' == 'H') {
		datev[row,]$'Soll/Haben-Kennzeichen' <<- 'S'
	}
	else {
		datev[row,]$'Soll/Haben-Kennzeichen' <<- 'H'
	}

	datev[row,]$'Konto'                          <<- account.from
	datev[row,]$'Gegenkonto (ohne BU-Schlüssel)' <<- account.to
}

#
# Fill data structure with the content of one single exported sheet
#
# @param sheet      Imported worksheet
# @param title      Sheet title
# @param account.7  Account number used for 7% tax entries
# @param account.19 Account number used for 19% tax entries
#
addTurnover <- function (sheet, title, account.7, account.19) {

	for (i in 1:nrow (sheet)) {

		line <- sheet[i,]
		bill.id <- line$Rechnungsnummer

		row <- nrow (data) + 1

		#
		# Skip lines with 0€ turnover. The import does report an error otherwise, because
		# this cannot be in the world of finance software.
		#
		if ( !is.na (line$Gesamtpreis.brutto) &&
                 !is.na (line$Rechnungsnummer) &&
                  round (abs (line$Gesamtpreis.brutto), 2) != 0 ) {

			data[row,]$bill.id          <<- bill.id
			data[row,]$payment.id       <<- NA
			data[row,]$payment.date     <<- NA
			data[row,]$item.kind        <<- title
			data[row,]$item.date        <<- line$Rechnungsdatum
			data[row,]$item.description <<- line$Position
			data[row,]$item.tax         <<- line$Steuersatz
			data[row,]$customer.id      <<- NA
			data[row,]$amount           <<- round (line$Gesamtpreis.brutto, 2)
			data[row,]$remarks          <<- NA
			data[row,]$responsible      <<- NA
			data[row,]$account          <<- NA

			if (!is.na (customer.ids[bill.id]))
				data[row,]$customer.id <<- customer.ids[bill.id]

			if (!is.na (remarks[bill.id]))
				data[row,]$remarks <<- remarks[bill.id]

			if (!is.na (responsible[bill.id]))
				data[row,]$responsible <<- responsible[bill.id]

			if (!is.na (payment.ids[bill.id]))
				data[row,]$payment.id <<- payment.ids[bill.id]

			if (!is.na (payment.dates[bill.id]))
				data[row,]$payment.date <<- payment.dates[bill.id]

			if (line$Steuersatz == 7)
				data[row,]$account <<- account.7
			else if (line$Steuersatz == 19)
				data[row,]$account <<- account.19
			else
				stop (paste ("Unknown tax level '", line$Steuersatz, "' at ", bill.id, sep=""))
		}
		else
			print (paste ("WARNING: 0€ turnover or empty fields at", line$Rechnungsnummer, sep=" "))
	}
}

#
# Add entry for payment from cash account
#
# @param sheet Imported worksheet
# @param title Sheet title
#
addPayment <- function (sheet, title) {

	s <- sheet[is.na (sheet$Rechnungsnummer),]

	for (i in 1:nrow (s)) {
		line = s[i,]

		row = nrow (data) + 1

		#
		# Skip lines with 0€ turnover. The import does report an error otherwise, because
		# this cannot be in the world of finance software.
		#
		if ( !is.na (line$Betrag) &&
                  round (abs (line$Betrag), 2) != 0 ) {

			data[row,]$bill.id          <<- NA
			data[row,]$payment.id       <<- line$Nummer
			data[row,]$payment.date     <<- line$Datum
			data[row,]$item.kind        <<- title
			data[row,]$item.date        <<- line$Datum
			data[row,]$item.description <<- line$Bemerkungen
			data[row,]$item.tax         <<- line$Steuersatz
			data[row,]$customer.id      <<- NA
			data[row,]$amount           <<- round (line$Betrag, 2)
			data[row,]$remarks          <<- NA
			data[row,]$responsible      <<- line$Benutzername

			if (line$Bemerkungen == "Geld auf Bank")
				data[row,]$account <<- account.bank
			else
				data[row,]$account <<- 0
		}
	}
}

#
# Withdraw the amount of money gathered via EC card in daily doses
#
addDailyCardTransfers <- function () {

	ec <- datev[datev$Zahlweise == 'EC Karte',]

	for (i in 1:nrow (ec)) {
		if (!is.na (ec$Soll[i]) && ec$Soll[i] == "S")
			ec$Umsatz[i] <- ec$Umsatz[i] * -1.0
	}

	sums <- tapply (ec$Umsatz, factor (ec$Belegdatum), sum)

	for (i in 1:length (sums)) {
		date <- names (sums)[i]
		sum <- sums[i]

		if (!is.na (sum) && round (sum, 2) > 0) {

			row = nrow(datev) + 1	

			datev[row,]$'Belegdatum'                     <<- date
			datev[row,]$'Umsatz (ohne Soll/Haben-Kz)'    <<- sum
			datev[row,]$'Soll/Haben-Kennzeichen'         <<- "H"
			datev[row,]$'Buchungstext'                   <<- "EC-Übertrag"
			datev[row,]$'Konto'                          <<- account.cash
			datev[row,]$'Gegenkonto (ohne BU-Schlüssel)' <<- account.bank
		}
	}
}

#
# Import sheet 'Zahlungen' and extract a customer id per bill number
#
sheet.zahlungen <- readWorksheetFromFile (input_file, sheet=5)

payments.all <- sheet.zahlungen[!is.na (sheet.zahlungen$Rechnungsnummer),]

bill.ids = levels (factor (payments.all$Rechnungsnummer))

reduce_vector <- function (v) {
	v <- v[!is.na (v)]
	v <- unique (v)
	return (paste (v, collapse=", "))
}

for (bill.id in bill.ids) {

	payment <- payments.all[payments.all$Rechnungsnummer == bill.id,]

	customer.ids[bill.id]  <- reduce_vector (payment$Kundennummer)
	remarks[bill.id]       <- reduce_vector (payment$Bemerkungen)
	responsible[bill.id]   <- reduce_vector (payment$Benutzername)
	payment.ids[bill.id]   <- reduce_vector (payment$Nummer)
	payment.dates[bill.id] <- reduce_vector (payment$Datum)
}

#
# Import sheets with relevant data and add them to the DATEV frame
#
sheet.leistungen <- readWorksheetFromFile (input_file, sheet=1)
#addTurnover (sheet.leistungen, title="Leistungen", account.7=8004, account.19=8004)

sheet.medikamente.angewendet <- readWorksheetFromFile (input_file, sheet=2)
#addTurnover (sheet.medikamente.angewendet, title="Medikamente (angewendet)", account.7=8011, account.19=8014)

sheet.medikamente.abgegeben <- readWorksheetFromFile (input_file, sheet=3)
#addTurnover (sheet.medikamente.abgegeben, title="Medikamente (abgegeben)", account.7=8021, account.19=8024)

sheet.produkte <- readWorksheetFromFile (input_file, sheet=4)
#addTurnover (sheet.produkte, title="Produkte", account.7=8031, account.19=8034)

sheet.cash <- readWorksheetFromFile (input_file, sheet=6)
addPayment (sheet.cash, title="Ausgabe")

#addDailyCardTransfers ()

#
# Write everything into the output file
#
# The turnover is formatted into a rounded string before.
#
#output <- datev

#output$'Umsatz (ohne Soll/Haben-Kz)' <- trim (format (round (output$'Umsatz (ohne Soll/Haben-Kz)', 2), nsmall=2, decimal.mark=","))

#handle <- file (output_file, encoding="latin1")
#write.table (output, file=handle, row.names=FALSE, quote=FALSE, na="", sep=";", dec=",", qmethod=c("escape", "double"))



#data <- data[order (data$bill.id),]


#
# Checks: 
#
# * Zero sum not allowed
#
