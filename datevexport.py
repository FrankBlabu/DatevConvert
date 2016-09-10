#
# datevexport.py - Test for a valid DATEV conversion and export
#
# Syntax: datevexport.py <backup zip file> <month (MM)> <year (YYYY)>
#

import csv
import io
import sys
import zipfile


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
        # Database content
        #
        self._data = []

        reader = csv.reader (io.TextIOWrapper (file, 'utf-8'), delimiter=',', quotechar='\"')

        keys = {}
        header_read = False
    
        for row in reader:

            if not header_read:
                for i in range (len (row)):
                    keys[i] = row[i]
                header_read = True
            else:
                line = {}
            
                for i in range (len (row)):
                    if (row[i] != 'NULL'):
                        line[keys[i]] = row[i].encode ('utf-8')

                self._data.append (line)


#---------------------------------------------------------------------
# CLASS Database
# 
# This class keeps the set of all file databases
#---------------------------------------------------------------------

class Database:

    def __init__ (self):
        self._data = {}

    #
    # Read new file database and add it to the content
    #
    def add (self, file, name):
        print ("Add: " + name)
        self._data[name] = FileDatabase (file)


#---------------------------------------------------------------------
# MAIN
#---------------------------------------------------------------------

#
# Database containg everything which was read
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
        print (entry)
        if (entry.endswith ("clients.csv")):
            with zip.open (entry, 'rU') as file:
                database.add (file, "clients")
        elif (entry.endswith ("invoices.csv")):
            with zip.open (entry, 'rU') as file:
                database.add (file, "invoices")
        elif (entry.endswith ("invoice_medication.csv")):
            with zip.open (entry, 'rU') as file:
                database.add (file, "invoices_medication")
        elif (entry.endswith ("invoice_product.csv")):
            with zip.open (entry, 'rU') as file:
                database.add (file, "invoices_product")
        elif (entry.endswith ("invoice_service.csv")):
            with zip.open (entry, 'rU') as file:
                database.add (file, "invoices_service")
        elif (entry.endswith ("payments.csv")):
            with zip.open (entry, 'rU') as file:
                database.add (file, "payments")
        elif (entry.endswith ("tax.csv")):
            with zip.open (entry, 'rU') as file:
                database.add (file, "tax")
