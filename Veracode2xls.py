import xml.etree.ElementTree as ET
import xlwt
from sys import argv
import argparse
import os

def create_excel(Flaws):

    # Creating book and 2 sheets
    book = xlwt.Workbook(encoding="utf-8")
    stats_sheet = book.add_sheet("Stats")
    flaws_sheet = book.add_sheet("Flaws")
    modules_sheet = book.add_sheet("Modules")
    
    # Preparing Stats sheet
    stats_sheet.write(0, 0, "Severity")
    stats_sheet.write(0, 1, "DB")
    stats_sheet.write(0, 2, "3RD party")
    stats_sheet.write(0, 3, "Company")
    stats_sheet.write(0, 4, "?")
    stats_sheet.write(0, 5, "Fix By Policy")
    stats_sheet.write(0, 6, "Comments")
    
    stats_sheet.write(1, 0, 5)
    stats_sheet.write(2, 0, 4)
    stats_sheet.write(3, 0, 3)
    stats_sheet.write(4, 0, 2)
    stats_sheet.write(5, 0, 1)
    stats_sheet.write(6, 0, "Total")
    
    stats_sheet.write(1, 1, '=COUNTIFS(Flaws!F:F;"1";Flaws!G:G;"DB")')
    stats_sheet.write(2, 1, '=COUNTIFS(Flaws!F:F;"2";Flaws!G:G;"DB")')
    stats_sheet.write(3, 1, '=COUNTIFS(Flaws!F:F;"3";Flaws!G:G;"DB")')
    stats_sheet.write(4, 1, '=COUNTIFS(Flaws!F:F;"4";Flaws!G:G;"DB")')
    stats_sheet.write(5, 1, '=COUNTIFS(Flaws!F:F;"5";Flaws!G:G;"DB")')
    stats_sheet.write(6, 1, "=SUM(B2:B6)")

    stats_sheet.write(1, 2, '=COUNTIFS(Flaws!F:F;"1";Flaws!G:G;"Third Party")')
    stats_sheet.write(2, 2, '=COUNTIFS(Flaws!F:F;"2";Flaws!G:G;"Third Party")')
    stats_sheet.write(3, 2, '=COUNTIFS(Flaws!F:F;"3";Flaws!G:G;"Third Party")')
    stats_sheet.write(4, 2, '=COUNTIFS(Flaws!F:F;"4";Flaws!G:G;"Third Party")')
    stats_sheet.write(5, 2, '=COUNTIFS(Flaws!F:F;"5";Flaws!G:G;"Third Party")')
    stats_sheet.write(6, 2, "=SUM(C2:C6)")
    
    stats_sheet.write(1, 3, '=COUNTIFS(Flaws!F:F;"1";Flaws!G:G;"Company")')
    stats_sheet.write(2, 3, '=COUNTIFS(Flaws!F:F;"2";Flaws!G:G;"Company")')
    stats_sheet.write(3, 3, '=COUNTIFS(Flaws!F:F;"3";Flaws!G:G;"Company")')
    stats_sheet.write(4, 3, '=COUNTIFS(Flaws!F:F;"4";Flaws!G:G;"Company")')
    stats_sheet.write(5, 3, '=COUNTIFS(Flaws!F:F;"5";Flaws!G:G;"Company")')
    stats_sheet.write(6, 3, "=SUM(D2:D6)")
    
    stats_sheet.write(1, 4, '=COUNTIFS(Flaws!F:F;"1";Flaws!G:G;"?")')
    stats_sheet.write(2, 4, '=COUNTIFS(Flaws!F:F;"2";Flaws!G:G;"?")')
    stats_sheet.write(3, 4, '=COUNTIFS(Flaws!F:F;"3";Flaws!G:G;"?")')
    stats_sheet.write(4, 4, '=COUNTIFS(Flaws!F:F;"4";Flaws!G:G;"?")')
    stats_sheet.write(5, 4, '=COUNTIFS(Flaws!F:F;"5";Flaws!G:G;"?")')
    stats_sheet.write(6, 4, "=SUM(E2:E6)")
    
    stats_sheet.write(1, 5, 0)
    stats_sheet.write(2, 5, 0)
    stats_sheet.write(3, 5, 0)
    stats_sheet.write(4, 5, 0)
    stats_sheet.write(5, 5, 0)
    stats_sheet.write(6, 5, "=SUM(F2:F6)")
        
    # Preparing Flaws sheet 
    flaws_sheet.write(0, 0, "Flaw id")
    flaws_sheet.write(0, 1, "Category")
    flaws_sheet.write(0, 2, "Sub Category")
    flaws_sheet.write(0, 3, "File path")
    flaws_sheet.write(0, 4, "Line")
    flaws_sheet.write(0, 5, "Severity")
    flaws_sheet.write(0, 6, "Propietary")

    # Preparing Modules sheet
    modules_sheet.write(0, 0, "Module")
    modules_sheet.write(0, 1, "Propietary")
    vulnerable_modules = set()    
        
    counter = 0
    
    # Writing all the flaws
    for flaw_category in Flaws:
        counter += 1
        flaw = flaw_category[0]
        category = flaw_category[1]
        
        # Flaw id
        flaws_sheet.write(counter, 0, int(flaw.get("issueid")))
        # Category
        flaws_sheet.write(counter, 1, category) 
        # Sub Category
        flaws_sheet.write(counter, 2, flaw.get("categoryname")) 
        # File path/sourcefile
        pathfile = flaw.get("sourcefilepath") + flaw.get("sourcefile")
        flaws_sheet.write(counter, 3, pathfile)
        vulnerable_modules.add(pathfile)
        # Line
        flaws_sheet.write(counter, 4, int(flaw.get("line"))) 
        # Severity
        flaws_sheet.write(counter, 5, int(flaw.get("severity")))
        # Propietary
        flaws_sheet.write(counter, 6, "=VLOOKUP(D" + str(counter+1) + ";Modules!A:B;2;FALSE)")
    
    
    counter = 0
    
    vulnerable_modules = list(vulnerable_modules)
    vulnerable_modules.sort()

    # Writing all vulnerable modules
    for module in vulnerable_modules:
        counter += 1
        modules_sheet.write(counter, 0, module)
        
    return book




parser = argparse.ArgumentParser(description='Transforms a Veracode report in XML format to an .xls file.')

parser.add_argument('infile', help='XML report you want to parse')
parser.add_argument('outfile', help='name of the output excel file')

args = parser.parse_args()

inxml = args.infile
outxls = args.outfile

if not(os.path.isfile(inxml)):
    print inxml + "doesn't exist."
    exit()

# Preparing to parse XML
tree = ET.parse(inxml)
root = tree.getroot()

Flaws = []

print "Parsing..."
# Go through the XML and create a list of lists of flaws to be able to associate a flaw to a category. Please note the namespace when finding nodes
for severity in root.findall("{https://www.veracode.com/schema/reports/export/1.0}severity"):
    for category in severity:
        for flaw in category.iter("{https://www.veracode.com/schema/reports/export/1.0}flaw"):
            Flaws.append([flaw, category.get("categoryname")])

print "Parsing done correctly."

print "Creating excel file..."
book = create_excel(Flaws)

book.save(outxls)

print "Excel file: " + outxls + " built correctly!"

print "If any functions have not been executed, copy the column, paste it in a text file, copy it and paste it back into the excel file!"
