#
# datevexport.py - Test for a valid DATEV conversion and export
#
# Syntax: datevexport.py <backup zip file> <month (MM)> <year (YYYY)>
#

import csv
import io
import sys
import time
import zipfile

#---------------------------------------------------------------------
# Auxillary functions
#---------------------------------------------------------------------

#
# Round into full euro and cents
#
def roundEuro (n):
    return round (100.0 * n + 0.0001) / 100.0



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
        print ("Add: " + name)
        self._data[name] = FileDatabase (file)


#---------------------------------------------------------------------
# CLASS Invoice
#---------------------------------------------------------------------

check_id = 3

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
        print ("Invoice: " + str (id) + " / " + self._number)
        
        self._total_products   = self.sum_content (database, "invoice_product", id)
        self._total_medication = self.sum_content (database, "invoice_medication", id)
        self._total_service    = self.sum_content (database, "invoice_service", id)

        sum = self._total_products + self._total_medication + self._total_service

        print ("  --> total: " + str (self._total) + ", sum: " + str (sum))

        assert (round (100 * self._total) == round (100 * sum)) 
        

    #
    # Sum content of a database file belonging to a given invoice id
    #
    # @param database   Database we are working with
    # @param file       Database file containing the detailed items
    # @param invoice_id Id of the invoice processed
    #
    def sum_content (self, database, file, invoice_id):
        total = 0.0

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
                    
                price = float (database.get (file, id, "price"))
                
                total += roundEuro (amount * factor * count * price)

                print ("  " + file + ", " + str (amount) +
                       " * " + str (factor) +
                       " * " + str (count) +
                       " * " + str (price) +
                       " = " + str (roundEuro (amount * factor * count * price)) +
                       " (" + str (amount * factor * count * price) + ")")

        return total
                
                


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
if len (sys.argv) != 4:
    sys.stderr.write ('Format: ' + sys.argv[0] + ' <backup zip file> <month (MM)> <year (YYYY)>')
    sys.exit (1)

with zipfile.ZipFile (sys.argv[1]) as zip:
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
# Generate invoice classes
#
invoices = {}

for id in database.range ("invoices"):
    if database.get ("invoices", id, "status") == "complete":
        invoices[id] = Invoice (database, id)
