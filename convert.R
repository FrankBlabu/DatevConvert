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

input_file  <- "c:/Users/Frank/Documents/Projects/DatevConvert/buchhaltung-export-2016-08.xlsx"
#input_file  <- "e:/test/convert/datevconvert/buchhaltung-export-2016-07.xlsx"

output_file <- "c:/Users/Frank/Documents/Projects/DatevConvert/datev-2016-08.csv"
#output_file <- "e:/test/convert/datevconvert/datev-2016-07.csv"

account.main     <- 1001
account.bank     <- 1360
account.card     <- 1361
account.transfer <- 1362

customer.ids  <- c ()
remarks       <- c ()
responsible   <- c ()
payment.ids   <- c ()
payment.dates <- c ()
bill.dates    <- c ()

#
# Intermediate frame for processing all entries
#
data <- data.frame (
	"bill.id"          = character (0), # Id of the bill
	"bill.date"        = character (0), # Date of the bill
	"payment.id"       = numeric (0),   # Id of the payment itself
	"payment.date"     = character (0), # Date of payment
	"payment.kind"     = character (0), # Way of payment (cash, card, ...)
	"payment.type"     = character (0), # Type of payment (transfer, ...)
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
# Datev frame for the final output format
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
    "BU-Schlüssel" = character (0),                # 8
    "Belegdatum" = character (0),                     # 9
    "Belegfeld 1" = character (0),                    # 10
    "Belegfeld 2" = character (0),                    # 11
    "Skonto" = double (0),                            # 12
    "Buchungstext" = character (0),                   # 13
    "Postensperre" = character(0),                    # 14
    "Diverse Adressnummer" = character (0),           # 15
    "Geschäftspartnerbank" = character (0),        # 16
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
    "Funktionsergänzung L+L" = character (0),      # 43
    "BU 49 Hauptfunktionstyp" = character (0),        # 44
    "BU 49 Hauptfunktionsnummer" = character (0),     # 45
    "BU 49 Funktionsergänzung" = character (0)     # 46
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
    "Stück" = integer (0),                         # 87
    "Gewicht" = double (0),                           # 88
    "Zahlweise" = character (0),                      # 89
    "Forderungsart" = character (0),                  # 90
    "Veranlagungsjahr" = character (0),               # 91
    "Zugeordnete Falligkeit" = character (0),         # 92
    "Skontotyp" = character (0),                      # 93
    "Auftragsnummer" = character (0),                 # 94
    "Buchungstyp" = character (0),                    # 95
    "USt-Schlüssel (Anzahlungen)" = double (0),    # 96
    "EU-Land (Anzahlungen)" = character (0),          # 97
    "Sachverhalt L+L (Anzahlungen)" = double (0),     # 98
    "EU-Steuersatz (Anzahlungen)" = double (0),       # 99
    "Erlöskonto (Anzahlungen)" = double (0),       # 100
    "Herkunft-Kz" = double (0),                       # 101
    "Buchungs GUID" = double (0),                     # 102
    "KOST-Datum" = as.Date (character ()),            # 103
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
    "Leistungsdatum" = character (),                  # 114
    "Datum Zuord. Steuerperiode" = character (),      # 115
    stringsAsFactors=FALSE, check.names=FALSE)



#
# Convert data (as.Date () or as.POSIXct ()) into string
#
convert_date <- function (date) {
	d <- as.POSIXlt (date)
	return (sprintf ("%04d-%02d-%02d", 1900 + d$year, d$mon + 1, d$mday))
}

#
# Skip leading blanks
#
trim <- function (text) {
	sub ("^\\s+", "", text)
}

#
# Reduce vector into comma separated representation
#
reduce_vector <- function (v) {
	v <- v[!is.na (v)]
	v <- unique (v)
	return (paste (v, collapse=", "))
}

#
# Fill data structure with the content of one single exported sheet
#
# @param sheet      Imported worksheet
# @param title      Sheet title
# @param account.7  Account number used for 7% tax entries
# @param account.19 Account number used for 19% tax entries
#
add_turnover <- function (sheet, title, account.7, account.19) {

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
			data[row,]$bill.date        <<- NA
			data[row,]$payment.id       <<- NA
			data[row,]$payment.date     <<- NA
			data[row,]$payment.kind     <<- line$Zahlungsweise
			data[row,]$payment.type     <<- "Umsatz"
			data[row,]$item.kind        <<- title
			data[row,]$item.date        <<- convert_date (line$Rechnungsdatum)
			data[row,]$item.description <<- line$Position
			data[row,]$item.tax         <<- line$Steuersatz
			data[row,]$customer.id      <<- NA
			data[row,]$amount           <<- round (line$Gesamtpreis.brutto, 2)
			data[row,]$remarks          <<- NA
			data[row,]$responsible      <<- NA
			data[row,]$account          <<- NA

			#
			# Insert information we ripped from the sheet 'Zahlungen'
			#
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

			if (!is.na (bill.dates[bill.id]))
				data[row,]$bill.date <<- bill.dates[bill.id]

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
add_payment <- function (sheet, title) {

	s <- sheet[is.na (sheet$Rechnungsnummer),]

	for (i in 1:nrow (s)) {
		line <- s[i,]

		row <- nrow (data) + 1

		#
		# Skip lines with 0€ turnover. The import does report an error otherwise, because
		# this cannot be in the world of finance software.
		#
		if ( !is.na (line$Betrag) &&
                  round (abs (line$Betrag), 2) != 0 ) {

			data[row,]$bill.id          <<- NA
			data[row,]$bill.date        <<- NA
			data[row,]$payment.id       <<- line$Nummer
			data[row,]$payment.date     <<- line$Datum
			data[row,]$payment.kind     <<- "Bar"
			data[row,]$payment.type     <<- "Barausgabe"
			data[row,]$item.kind        <<- title
			data[row,]$item.date        <<- line$Datum
			data[row,]$item.description <<- line$Bemerkungen
			data[row,]$item.tax         <<- line$Steuersatz
			data[row,]$customer.id      <<- NA
			data[row,]$amount           <<- round (line$Betrag, 2)
			data[row,]$remarks          <<- NA
			data[row,]$responsible      <<- line$Benutzername

			#
			# Special treatment for deposits from the cash into the bank account
			#
			if (startsWith (line$Bemerkungen, "Geld auf Bank")) {
				data[row,]$account      <<- account.bank
				data[row,]$payment.type <<- "Einzahlung"
			}

			#
			# All other cash payments do not have an account number
			#
			else {
				data[row,]$account <<- 0
			}
		}
	}
}

#
# Withdraw the amount of money gathered via EC card per transaction
#
# @param sheet Imported worksheet
# @param title Sheet title
#
add_card_transfers <- function (sheet, title) {
	
	s <- sheet[!is.na (sheet$Nummer),]

	for (i in 1:nrow (sheet)) {
		line <- sheet[i,]

		#
		# Process all EC card payments which are not canceled out by a matching
		# payment number ending with 'X'. For each payment, an extra transfer to 
		# the EC card account is added.
		#
		if ( !is.na (line$Nummer) && 
                 !endsWith (line$Nummer, "X") &&
                 nrow (s[s$Nummer == paste (line$Nummer, "X", sep=""),]) == 0 &&
                 line$Zahlungsweise == "EC Karte" ) {

			row <- nrow (data) + 1

			data[row,]$bill.id          <<- line$Rechnungsnummer
			data[row,]$bill.date        <<- NA
			data[row,]$payment.id       <<- line$Nummer
			data[row,]$payment.date     <<- convert_date (line$Datum)
			data[row,]$payment.kind     <<- NA
			data[row,]$payment.type     <<- "Umbuchung"
			data[row,]$item.kind        <<- title
			data[row,]$item.date        <<- convert_date (line$Datum)
			data[row,]$item.description <<- line$Bemerkungen
			data[row,]$item.tax         <<- 0
			data[row,]$customer.id      <<- line$Kundennummer
			data[row,]$amount           <<- round (-line$Betrag, 2)
			data[row,]$remarks          <<- "Übertrag EC-Karten-Zahlung"
			data[row,]$responsible      <<- line$Benutzername
			data[row,]$account          <<- account.card

			if (!is.na (bill.dates[bill.id]))
				data[row,]$bill.date <<- bill.dates[bill.id]

		}
	}
}

#
# Generate DATEV representation from intermediate format
#
generate_datev <- function () {

	for (i in 1:nrow (data)) {
		line <- data[i,]
		row <- nrow (datev) + 1

		#
		# The transferred sum us is always unsigned
		#
		datev[row,]$'Umsatz (ohne Soll/Haben-Kz)' <<- abs (line$amount)

		#
		# 'H' means that the transfer is from 'Konto' to 'Gegenkonto' while
		# 'S' is the other way round
		#
		if (line$amount < 0)
			datev[row,]$'Soll/Haben-Kennzeichen' <<- "S"
		else
			datev[row,]$'Soll/Haben-Kennzeichen' <<- "H"

		if (line$item.tax == 19)
			datev[row,]$'BU-Schlüssel' <<- 3
		else if (line$item.tax == 7)
			datev[row,]$'BU-Schlüssel' <<- 2

		datev[row,]$'Konto' <<- line$account

		#
		# All payments are booked as turnarounds to the cash account. The only
		# exception are transfers which are going into the transfer account directly.
		#
		if (!is.na (line$payment.kind) && line$payment.kind == "Überweisung")
			datev[row,]$'Gegenkonto (ohne BU-Schlüssel)' <<- account.transfer
		else
			datev[row,]$'Gegenkonto (ohne BU-Schlüssel)' <<- account.main

		datev[row,]$'Belegdatum'         <<- format (as.Date (line$payment.date), "%d%m%Y")
		datev[row,]$'Buchungstext'       <<- line$item.description
		datev[row,]$'EU-Steuersatz'      <<- line$item.tax
		datev[row,]$'Zahlweise'          <<- line$payment.kind
		datev[row,]$'Buchungstyp'        <<- line$payment.type
		datev[row,]$'Leistungsdatum'     <<- format (as.Date (line$item.date), "%d%m%Y")
		datev[row,]$'Gesellschaftername' <<- line$responsible
		datev[row,]$'Sachverhalt'        <<- line$item.kind

		if (!is.na (line$bill.id))
			datev[row,]$'Buchungstext' <<- paste ("Rechnung ", line$bill.id)

		datev[row,]$'Beleginfo - Art 1'    <<- "Rechnungsnummer"
		datev[row,]$'Beleginfo - Inhalt 1' <<- line$bill.id
		datev[row,]$'Beleginfo - Art 2'    <<- "Rechnungsdatum"
		datev[row,]$'Beleginfo - Inhalt 2' <<- line$bill.date
		datev[row,]$'Beleginfo - Art 3'    <<- "Vorgangsnummer"
		datev[row,]$'Beleginfo - Inhalt 3' <<- line$payment.id
		datev[row,]$'Beleginfo - Art 4'    <<- "Typ"
		datev[row,]$'Beleginfo - Inhalt 4' <<- line$item.kind
		datev[row,]$'Beleginfo - Art 5'    <<- "Kundennummer"
		datev[row,]$'Beleginfo - Inhalt 5' <<- line$customer.id
		datev[row,]$'Beleginfo - Art 6'    <<- "Bemerkungen"
		datev[row,]$'Beleginfo - Inhalt 6' <<- line$remarks
		
	}
}

#----------------------------------------------------------------------------------
# MAIN
#----------------------------------------------------------------------------------


#
# Import sheet 'Zahlungen' and extract customer data per bill number
#
# Because the dates are eventually converted into integers, the column types
# have to be specified explicitly.
#
sheet.zahlungen <- readWorksheetFromFile (input_file, sheet="Zahlungen", forceConversion=TRUE,
	colTypes = c (XLC$DATA_TYPE.STRING,   # Rechnungsnummer
                    XLC$DATA_TYPE.STRING,   # Nummer
                    XLC$DATA_TYPE.DATETIME, # Datum
                    XLC$DATA_TYPE.NUMERIC,  # Betrag
                    XLC$DATA_TYPE.STRING,   # Zahlungsweise
                    XLC$DATA_TYPE.STRING,   # Bemerkungen
                    XLC$DATA_TYPE.STRING,   # Kundennummer
                    XLC$DATA_TYPE.STRING    # Benutzername
))

#
# Iterate over all known bill ids and extract useful information
#
payments.all <- sheet.zahlungen[!is.na (sheet.zahlungen$Rechnungsnummer),]
payment.bill.ids <- levels (factor (payments.all$Rechnungsnummer))

for (bill.id in payment.bill.ids) {

	payment <- payments.all[payments.all$Rechnungsnummer == bill.id,]

	customer.ids[bill.id]  <- reduce_vector (payment$Kundennummer)
	remarks[bill.id]       <- reduce_vector (payment$Bemerkungen)
	responsible[bill.id]   <- reduce_vector (payment$Benutzername)
	payment.ids[bill.id]   <- reduce_vector (payment$Nummer)

	dates <- convert_date (payment$Datum)
	dates <- dates[!is.na (dates)]
	dates <- unique (dates)
	
	if (length (dates) > 0)
		payment.dates[bill.id] <- dates[1]
}

#
# Import sheet 'Rechnungen' to extract the correct bill date
# 
sheet.rechnungen <- readWorksheetFromFile (input_file, sheet="Rechnungen", forceConversion=TRUE)                    
invoices.all <- sheet.rechnungen[!is.na (sheet.rechnungen$Rechnungsnummer),]
invoice.bill.ids <- levels (factor (invoices.all$Rechnungsnummer))

for (bill.id in invoice.bill.ids) {
	invoice <- invoices.all[invoices.all$Rechnungsnummer == bill.id,]
	bill.dates[bill.id] <- reduce_vector (convert_date (invoice$Rechnungsdatum))
}


#
# Import sheets with relevant turnaround and cash payment data and add them to the 
# internal data representation
#
sheet.leistungen <- readWorksheetFromFile (input_file, sheet="Leistungen")
add_turnover (sheet.leistungen, title="Leistungen", account.7=8004, account.19=8004)

sheet.medikamente.angewendet <- readWorksheetFromFile (input_file, sheet="Medikamente angewendet")
add_turnover (sheet.medikamente.angewendet, title="Medikamente (angewendet)", account.7=8011, account.19=8014)

sheet.medikamente.abgegeben <- readWorksheetFromFile (input_file, sheet="Medikamente abgegeben")
add_turnover (sheet.medikamente.abgegeben, title="Medikamente (abgegeben)", account.7=8021, account.19=8024)

sheet.produkte <- readWorksheetFromFile (input_file, sheet="Produkte")
add_turnover (sheet.produkte, title="Produkte", account.7=8031, account.19=8034)

sheet.cash <- readWorksheetFromFile (input_file, sheet="Zahlungen MwSt")
add_payment (sheet.cash, title="Ausgabe")

#
# Add counter transfers for EC card payments
#
add_card_transfers (sheet.zahlungen, "Umbuchung EC-Karten-Zahlung")

#
# Sort whole table
#
#data <- data[order (data$bill.id),]


generate_datev ()

#
# Write everything into the output file
#
output <- datev

handle <- file (output_file, encoding="latin1")
write.table (output, file=handle, row.names=FALSE, quote=FALSE, na="", sep=";", dec=",", qmethod=c("escape", "double"))
