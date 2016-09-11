#
# datevexport.py - Test for a valid DATEV conversion and export
#
# Syntax: datevexport.py <backup zip file> <month (MM)> <year (YYYY)> <output file>
#

import copy
import csv
import io
import sys
import time
import zipfile

#---------------------------------------------------------------------
# Configuration
#---------------------------------------------------------------------

#
# Account number used for DATEV transfers
#
class Accounts:
    Null     = 0000
    Main     = 1001
    Bank     = 1360
    EC       = 1361
    Transfer = 1362

#
# DATEV table column headers
#
datev_headers = [
    "Umsatz (ohne Soll/Haben-Kz)",    #1
    "Soll/Haben-Kennzeichen",         #2
    "WKZ Umsatz",                     #3
    "Kurs",                           #4
    "Basis-Umsatz",                   #5
    "WKZ Basis-Umsatz",               #6
    "Konto",                          #7
    "Gegenkonto (ohne BU-Schlüssel)", #8
    "BU-Schlüssel",                   #9
    "Belegdatum",                     #10
    "Belegfeld 1",                    #11
    "Belegfeld 2",                    #12
    "Skonto",                         #13
    "Buchungstext",                   #14
    "Postensperre",                   #15
    "Diverse Adressnummer",           #16
    "Geschäftspartnerbank",           #17
    "Sachverhalt",                    #18
    "Zinssperre",                     #19
    "Beleglink",                      #20
    "Beleginfo - Art 1",              #21
    "Beleginfo - Inhalt 1",           #22
    "Beleginfo - Art 2",              #23
    "Beleginfo - Inhalt 2",           #24
    "Beleginfo - Art 3",              #25
    "Beleginfo - Inhalt 3",           #26
    "Beleginfo - Art 4",              #27
    "Beleginfo - Inhalt 4",           #28
    "Beleginfo - Art 5",              #29
    "Beleginfo - Inhalt 5",           #30
    "Beleginfo - Art 6",              #31
    "Beleginfo - Inhalt 6",           #32
    "Beleginfo - Art 7",              #33
    "Beleginfo - Inhalt 7",           #34
    "Beleginfo - Art 8",              #35
    "Beleginfo - Inhalt 8",           #36
    "KOST1 - Kostenstelle",           #37
    "KOST2 - Kostenstelle",           #38
    "Kost-Menge",                     #39
    "EU-Land u. UStID",               #40
    "EU-Steuersatz",                  #41
    "Abw. Versteuerungsart",          #42
    "Sachverhalt L+L",                #43
    "Funktionsergänzung L+L",         #44
    "BU 49 Hauptfunktionstyp",        #45
    "BU 49 Hauptfunktionsnummer",     #46
    "BU 49 Funktionsergänzung",       #47
    "Zusatzinformation - Art 1",      #48
    "Zusatzinformation- Inhalt 1",    #49
    "Zusatzinformation - Art 2",      #50
    "Zusatzinformation- Inhalt 2",    #51
    "Zusatzinformation - Art 3",      #52
    "Zusatzinformation- Inhalt 3",    #53
    "Zusatzinformation - Art 4",      #54
    "Zusatzinformation- Inhalt 4",    #55
    "Zusatzinformation - Art 5",      #56
    "Zusatzinformation- Inhalt 5",    #57
    "Zusatzinformation - Art 6",      #58
    "Zusatzinformation- Inhalt 6",    #59
    "Zusatzinformation - Art 7",      #60
    "Zusatzinformation- Inhalt 7",    #61
    "Zusatzinformation - Art 8",      #62
    "Zusatzinformation- Inhalt 8",    #63
    "Zusatzinformation - Art 9",      #64
    "Zusatzinformation- Inhalt 9",    #65
    "Zusatzinformation - Art 10",     #66
    "Zusatzinformation- Inhalt 10",   #67
    "Zusatzinformation - Art 11",     #68
    "Zusatzinformation- Inhalt 11",   #69
    "Zusatzinformation - Art 12",     #70
    "Zusatzinformation- Inhalt 12",   #71
    "Zusatzinformation - Art 13",     #72
    "Zusatzinformation- Inhalt 13",   #73
    "Zusatzinformation - Art 14",     #74
    "Zusatzinformation- Inhalt 14",   #75
    "Zusatzinformation - Art 15",     #76
    "Zusatzinformation- Inhalt 15",   #77
    "Zusatzinformation - Art 16",     #78
    "Zusatzinformation- Inhalt 16",   #79
    "Zusatzinformation - Art 17",     #80
    "Zusatzinformation- Inhalt 17",   #81
    "Zusatzinformation - Art 18",     #82
    "Zusatzinformation- Inhalt 18",   #83
    "Zusatzinformation - Art 19",     #84
    "Zusatzinformation- Inhalt 19",   #85
    "Zusatzinformation - Art 20",     #86
    "Zusatzinformation- Inhalt 20",   #87
    "Stück",                          #88
    "Gewicht",                        #89
    "Zahlweise",                      #90
    "Forderungsart",                  #91
    "Veranlagungsjahr",               #92
    "Zugeordnete Falligkeit",         #93
    "Skontotyp",                      #94
    "Auftragsnummer",                 #95
    "Buchungstyp",                    #96
    "USt-Schlüssel (Anzahlungen)",    #97
    "EU-Land (Anzahlungen)",          #98
    "Sachverhalt L+L (Anzahlungen)",  #99
    "EU-Steuersatz (Anzahlungen)",    #100
    "Erlöskonto (Anzahlungen)",       #101
    "Herkunft-Kz",                    #102
    "Buchungs GUID",                  #103
    "KOST-Datum",                     #104
    "SEPA-Mandatsreferenz",           #105
    "Skontosperre",                   #106
    "Gesellschaftername",             #107
    "Beteiligtennummer",              #108
    "Identifikationsnummer",          #109
    "Zeichnernummer",                 #110
    "Postensperre bis",               #111
    "Bezeichnung SoBil-Sachverhalt",  #112
    "Kennzeichen SoBil-Buchung",      #113
    "Festschreibung",                 #114
    "Leistungsdatum",                 #115
    "Datum Zuord. Steuerperiode"      #116
    ]

#
# Readable keys for mapping a semantics to a DATEV columns
#
datev_column_mapping = {
    'umsatz'            : 1,
    'soll_haben'        : 2,
    'konto'             : 7,
    'gegenkonto'        : 8,
    'bu_schluessel'     : 9,
    'belegdatum'        : 10,
    'belegfeld_1'       : 11,
    'belegfeld_2'       : 12,
    'buchungstext'      : 14,
    'sachverhalt'       : 18,
    'beleginfo_art_1'   : 21,
    'beleginfo_inhalt_1': 22,
    'beleginfo_art_2'   : 23,
    'beleginfo_inhalt_2': 24,
    'beleginfo_art_3'   : 25,
    'beleginfo_inhalt_3': 26,
    'beleginfo_art_4'   : 27,
    'beleginfo_inhalt_4': 28,
    'beleginfo_art_5'   : 29,
    'beleginfo_inhalt_5': 30,
    'beleginfo_art_6'   : 31,
    'beleginfo_inhalt_6': 32,
    'beleginfo_art_7'   : 33,
    'beleginfo_inhalt_7': 34,
    'beleginfo_art_8'   : 35,
    'beleginfo_inhalt_8': 36,
    'eu_steuersatz'     : 41,
    'stueck'            : 88,
    'zahlweise'         : 90,
    'auftragsnummer'    : 95,
    'buchungstyp'       : 96,
    'gesellschaftername': 107,
    'leistungsdatum'    : 115
}

    
#---------------------------------------------------------------------
# Auxillary functions
#---------------------------------------------------------------------

#
# Round into full euro and cents
#
def roundEuro (n):
    return round (100.0 * n + 0.0001) / 100.0

#
# Convert date string representation into Python date class
#
def toDate (text):
    return  time.strptime (text, "%Y-%m-%d %H:%M:%S")


#---------------------------------------------------------------------
# CLASS FileDatabase
#
# This class keeps the content of a single CSV file in a data frame
# like format
#---------------------------------------------------------------------

class FileDatabase:

    #
    # Constructor
    #
    # @param file Opened file containing the CSV data. File content fill be read here.
    #
    def __init__ (self, file):

        #
        # Database content (id, row content)
        #
        self._data = {}

        reader = csv.reader (io.TextIOWrapper (file, 'utf-8'), delimiter=',', quotechar='\"')

        keys = {}
        header_read = False
    
        for row in reader:
            if not header_read:
                for i in range (len (row)):
                    keys[i] = row[i]
                    
                header_read = True
            else:
                id = None
                line = {}
                
                for i in range (len (row)):
                    key = keys[i]
                    
                    if (key == 'id'):
                        id = int (row[i])
                    if (row[i] != 'NULL'):
                        line[key] = row[i]
                    else:
                        line[key] = ""

                assert (id != None)
                self._data[id] = line

    #
    # Check if the database supports the given key
    #
    # @param key Key to check 
    #
    def has (self, key):
        id = list (self._data.keys ())[0]
        return key in self._data[id];
                
    #
    # Return single cell content
    #
    # @param id  Id of the entry
    # @param key Key of the column to access
    #
    def get (self, id, key):
        assert (id in self._data)

        data = self._data[id]
        
        assert (isinstance (data, dict))
        assert (key in data)

        return data[key]

    #
    # Return range of ids present in the file database
    #
    def range (self):
        return self._data.keys ()


#---------------------------------------------------------------------
# CLASS Database
# 
# This class keeps the set of all file databases
#---------------------------------------------------------------------

class Database:

    def __init__ (self):
        self._data = {}

    #
    # Return single cell content of a file database
    #
    # @param database Name of the file database to access
    # @param id       Id of the entry
    # @param key      Key of the column to access
    #
    def get (self, database, id, key):
        assert (database in self._data)
        return self._data[database].get (id, key)
        
    #
    # Check if the database supports the given key
    #
    # @param key Key to check 
    #
    def has (self, database, key):
        assert (database in self._data)
        return self._data[database].has (key)
    
    #
    # Return range of ids present in the file database
    #
    def range (self, database):
        assert (database in self._data)
        return self._data[database].range ()
    
    #
    # Read new file database and add it to the content
    #
    def add (self, file, name):
        self._data[name] = FileDatabase (file)


#---------------------------------------------------------------------
# CLASS Invoice
#---------------------------------------------------------------------

#
# This class keeps record about a single invoice, split into different
# tax related parts
#
class Invoice:

    #
    # Constructor
    #
    # @param id Unique invoice id
    #
    def __init__ (self, database, id):

        #
        # Gather some information about the invoice itself
        #
        self._id        = id
        self._client_id = database.get ("invoices", id, "client_id")
        self._number    = database.get ("invoices", id, "number")

        date = database.get ("invoices", id, "date")
        self._date  = time.strptime (date, '%Y-%m-%d') if date else None
        self._total = float (database.get ("invoices", id, "total"))

        #
        # Collect parts of the invoice which must sum up to the total and
        # will be used to split the total into the different tax parts
        #
        if False:
            print ("Invoice: " + str (id) + " / " + self._number)
        
        self._total_products   = self.sum_content (database, "invoice_product", id)
        self._total_medication = self.sum_content (database, "invoice_medication", id)
        self._total_service    = self.sum_content (database, "invoice_service", id)

        total = sum (self._total_products.values ()) + \
                sum (self._total_medication.values ()) + \
                sum (self._total_service.values ())


        if False:
            print ("  --> total: " + str (self._total) + ", sum: " + str (total))

        assert (round (100 * self._total) == round (100 * total)) 
        

    #
    # Sum content of a database file belonging to a given invoice id
    #
    # @param database   Database we are working with
    # @param file       Database file containing the detailed items
    # @param invoice_id Id of the invoice processed
    # @return Dictionary of sums with the tax ids as keys
    #
    def sum_content (self, database, file, invoice_id):

        total = {}

        for id in database.range (file):
            if int (database.get (file, id, "invoice_id")) == invoice_id:

                amount = 1.0
                if database.has (file, "amount"):
                    amount = float (database.get (file, id, "amount"))

                factor = 1.0
                if database.has (file, "factor"):
                    factor = float (database.get (file, id, "factor"))

                count = 1.0
                if database.has (file, "count"):
                    count = float (database.get (file, id, "count"))

                tax_id = database.get (file, id, "tax_id")
                    
                price = float (database.get (file, id, "price"))

                if not tax_id in total:
                    total[tax_id] = 0.0
                
                total[tax_id] += roundEuro (amount * factor * count * price)

                if False:
                    print ("  " + file + ", " + str (amount) +
                           " * " + str (factor) +
                           " * " + str (count) +
                           " * " + str (price) +
                           " = " + str (roundEuro (amount * factor * count * price)) +
                           " (" + str (amount * factor * count * price) + ")")

        return total
                
#---------------------------------------------------------------------
# CLASS DatevEntry
#---------------------------------------------------------------------

#
# Class representing a single DATEV file entry
#
# Valid fields are:
#
# - invoice_id       - Id of the invoice
# - invoice_date     - Date of the invoice
# - payment_id       - Id of the payment itself
# - payment_date     - Date of payment
# - payment_kind     - Way of payment (cash, card, ...)
# - payment_type     - Type of payment (transfer, ...)
# - item_kind        - Kind of item applied/sold
# - item_date        - Date the item was applied/sold
# - item_description - Description of the item
# - item_tax         - Tax of the item
# - customer_id      - Id of the customer
# - amount           - Amount of money
# - remarks          - Payment remarks
# - responsible      - Name of the responsible person
# - account          - Account where the money goes to / came from
class DatevEntry:

    #
    # Constructor (regular payment)
    #
    # @param database   Database we are working with
    # @param payment_id Id of the payment to be processed
    #
    def __init__ (self, database, payment_id):

        #
        # The payment type is not known yet. Setup common fields only.
        #
        self._id               = payment_id
        self._invoice_id       = ""
        self._invoice_date     = ""
        self._payment_id       = payment_id
        self._payment_date     = toDate (self.get (database, "date"))
        self._payment_kind     = self.get (database, "method")
        self._item_kind        = self.get (database, "paymenttype")
        self._item_date        = self._payment_date
        self._item_description = self.get (database, "notes")
        self._item_tax         = ""
        self._customer_id      = ""
        self._amount           = round (float (self.get (database, "amount")))
        self._remarks          = ""
        self._responsible      = self.get (database, "username")
        self._account          = Accounts.Null
        self._payment_type     = ""

    #
    # Setup invoice based payment
    #
    def setupInvoiceEntry (self, database):
        pass
            
    #
    # Setup non invoice payment
    #
    def setupNonInvoiceEntry (self, database):
        if (self.get (database, "paymenttype").lower ().startswith ("geld auf bank")):
            self._account      = Accounts.Bank
            self._payment_type = "Einzahlung"
        else:
            self._account      = Accounts.Null
            self._payment_type = "Barausgabe"


    #
    # Setup counter entry for moving another payment via ec card onto a
    # special account for accounting purposes
    #
    def setupECCounterEntry (self):
        self._account      = Accounts.EC
        self._payment_type = "Umbuchung"
        self._remarks      = "Übertrag EC-Karten-Zahlung"
        pass

    #
    # Query database for payment entry
    #
    def get (self, database, key):
        return database.get ("payments", self._id, key)

    #
    # Convert entry into vector of DATEV rows
    #
    def toDatev (self):
        pass
    
#---------------------------------------------------------------------
# MAIN
#---------------------------------------------------------------------

#
# Database instance containg everything which was read
#
database = Database ()

#
# Parse command line arguments
#
if len (sys.argv) != 5:
    sys.stderr.write ('Format: ' + sys.argv[0] + ' <backup zip file> <month (MM)> <year (YYYY)> <output file>')
    sys.exit (1)

filename = sys.argv[1]
month    = int (sys.argv[2])
year     = int (sys.argv[3])
output   = sys.argv[4]
    
#
#
#
with zipfile.ZipFile (filename) as zip:
    for entry in zip.namelist ():
        if (entry.endswith ("clients.csv")):
            with zip.open (entry, 'rU') as file:
                database.add (file, "clients")
        elif (entry.endswith ("invoices.csv")):
            with zip.open (entry, 'rU') as file:
                database.add (file, "invoices")
        elif (entry.endswith ("invoice_medication.csv")):
            with zip.open (entry, 'rU') as file:
                database.add (file, "invoice_medication")
        elif (entry.endswith ("invoice_product.csv")):
            with zip.open (entry, 'rU') as file:
                database.add (file, "invoice_product")
        elif (entry.endswith ("invoice_service.csv")):
            with zip.open (entry, 'rU') as file:
                database.add (file, "invoice_service")
        elif (entry.endswith ("payments.csv")):
            with zip.open (entry, 'rU') as file:
                database.add (file, "payments")
        elif (entry.endswith ("tax.csv")):
            with zip.open (entry, 'rU') as file:
                database.add (file, "tax")

#
# Generate invoice dictionary
#
invoices = {}

for id in database.range ("invoices"):
    if database.get ("invoices", id, "status") == "complete":
        invoices[id] = Invoice (database, id)

#
# Process payment list for the given month
#
datev = []

for id in database.range ("payments"):
    date = toDate (database.get ("payments", id, "date"))

    #
    # Given month only
    #
    if date.tm_year == year and date.tm_mon == month:

        #
        # Accountants tax application cannot process payments with 0€ amount
        #
        if roundEuro (float (database.get ("payments", id, "amount"))) != 0:

            #
            # Skip cancelled payments
            #
            if not database.get ("payments", id, "deleted"):

                #
                # Setup general DATEV entry will all common fields initialized
                #
                entry = DatevEntry (database, id)

                #
                # Case 1: Invoice based payment
                #
                if database.get ("payments", id, "invoice_id"):
                    pass

                #
                # Case 2: Non invoice based payment
                #
                else:
                    entry.setupNonInvoiceEntry (database);
                    datev.append (entry)

                #
                # In case of EC card payments, setup counter entry
                #
                if entry._payment_kind == "ec":
                    counter_entry = copy.deepcopy (entry)
                    counter_entry.setupECCounterEntry ()
                    datev.append (counter_entry)

#
# Extract result as DATEV file
#
with open (output, 'w', newline='') as file:
    writer = csv.writer (file, dialect='excel')

    for entry in datev:
        writer.writerow (entry.toDatev ())
