from itertools import count
from re import S
import xml.etree.ElementTree as ET
from sys import argv
import argparse
import os
import xlsxwriter
from xml.etree import ElementTree


def create_excel(Flaws):
    book = xlsxwriter.Workbook('Book1.xlsx')
    flaws_sheet = book.add_worksheet("Flaws")

    # Preparing Flaws sheet 
    flaws_sheet.write(0, 0, "Flaw id")
    flaws_sheet.write(0, 1, 'CWE ID')
    flaws_sheet.write(0, 2, "Category Name")
    flaws_sheet.write(0, 3, "Description")
    flaws_sheet.write(0, 4, "Affects Policy Compliance")
    flaws_sheet.write(0, 5, "Exploit (Manual)")
    flaws_sheet.write(0, 6, "Severity (Manual")
    flaws_sheet.write(0, 7, "Remediation (Manual)")
    flaws_sheet.write(0, 8, "Date First Occurrence")
    flaws_sheet.write(0, 9, 'Module')
    flaws_sheet.write(0, 10, "Source File")
    flaws_sheet.write(0, 11, "Source File Path")
    flaws_sheet.write(0, 12, "Attack Vector")
    flaws_sheet.write(0, 13, "Function Prototype")
    flaws_sheet.write(0, 14, "Line")
    flaws_sheet.write(0, 15, "Function Relative Location (%)")
    flaws_sheet.write(0, 16, "Scope")
    flaws_sheet.write(0, 17, 'Severity')
    flaws_sheet.write(0, 18, "Exploitability Adjustments")
    flaws_sheet.write(0, 19, "Grace Period Expires")
    flaws_sheet.write(0, 20, "Remediation Status")
    flaws_sheet.write(0, 21, "Mitigation Status")
    flaws_sheet.write(0, 22, "Mitigation Status Description")
    flaws_sheet.write(0, 23, "Mitigation Text")
        
    counter = 0
    
    # Writing all the flaws
    for flaw_category in Flaws:
        counter += 1
        flaw = flaw_category[0]
        category = flaw_category[1]
        
        # Flaw id
        flaws_sheet.write(counter, 0, int(flaw.get("issueid")))
        # CWE ID
        flaws_sheet.write(counter, 1, int(flaw.get('cweid'))) 
        # Category Name
        flaws_sheet.write(counter, 2, flaw.get("categoryname")) 
        # Description
        flaws_sheet.write(counter, 3, flaw.get('description'))
        # Affects Policy compliance
        flaws_sheet.write(counter, 4, flaw.get('affects_policy_compliance'))
        # Exploit Manual
        flaws_sheet.write(counter, 5, flaw.get('exploit_desc'))
        # Severity Manual 
        flaws_sheet.write(counter, 6, flaw.get('severity_desc'))
        # Remediation (Manual)
        flaws_sheet.write(counter, 7, flaw.get('remediation_desc'))
        # Date First Occurrence
        flaws_sheet.write(counter, 8, flaw.get('date_first_occurrence'))
        # Module
        flaws_sheet.write(counter, 9, flaw.get('module'))
        # Source File
        flaws_sheet.write(counter, 10, flaw.get("sourcefile"))
        # Source File Path
        flaws_sheet.write(counter, 11, flaw.get("sourcefilepath"))
        # Attack Vector
        flaws_sheet.write(counter, 12, flaw.get('type'))
        # Function Prototype
        flaws_sheet.write(counter, 13, flaw.get('functionprototype'))
        # Line
        flaws_sheet.write(counter, 14, int(flaw.get('line')))
        # Function Relative location
        flaws_sheet.write(counter, 15, int(flaw.get('functionrelativelocation')))
        # Scope
        flaws_sheet.write(counter, 16, flaw.get('scope'))
        # Severity
        if int(flaw.get('severity')) == 5:
            flaws_sheet.write(counter, 17, "5 - Very High")
        elif int(flaw.get('severity')) == 4:
            flaws_sheet.write(counter, 17, "4 - High")
        elif int(flaw.get('severity')) == 3:
            flaws_sheet.write(counter, 17, "3 - Medium")
        elif int(flaw.get('severity')) == 2:
            flaws_sheet.write(counter, 17, "2 - Low")
        elif int(flaw.get('severity')) == 1:
            flaws_sheet.write(counter, 17, "1 - Very Low")
        else:
            flaws_sheet.write(counter, 17, "0 - Informational")
        # Exploitability Adjustments
        flaws_sheet.write(counter, 18, flaw.get('exploitability_adjustments'))
        # Grace Period Expires
        flaws_sheet.write(counter, 19, flaw.get('grace_period_expires'))
        # Remediation Status
        flaws_sheet.write(counter, 20, flaw.get('remediation_status'))
        # Mitigation Status
        flaws_sheet.write(counter, 21, flaw.get('mitigation_status'))
        # Mitigation Status Description
        flaws_sheet.write(counter, 22, flaw.get('mitigation_status_desc'))
        # Mitigation Text
        flaws_sheet.write(counter, 23, flaw.get('mitigation_text'))
    
    counter = 0
        
    return book


parser = argparse.ArgumentParser(description='Transforms a Veracode report in XML format to an .xlsx file.')

parser.add_argument('-i','--infile', help='XML report you want to parse')
parser.add_argument('-o','--outfile', help='name of the output excel file')
# parser.add_argument()

args = parser.parse_args()
# print(args)

inxml = args.infile
outxls = args.outfile

if not(os.path.isfile(inxml)):
    print(inxml + "doesn't exist.")
    exit()

#Preparing to parse XML
tree = ElementTree.parse(inxml)
root = tree.getroot()

Flaws = []

print("Parsing...")
# Go through the XML and create a list of lists of flaws to be able to associate a flaw to a category. Please note the namespace when finding nodes
for severity in root.findall("{https://www.veracode.com/schema/reports/export/1.0}severity"):
    for category in severity:
        for flaw in category.iter("{https://www.veracode.com/schema/reports/export/1.0}flaw"):
            Flaws.append([flaw, category.get("categoryname")])

print("Parsing done correctly.")

print("Creating excel file...")
book = create_excel(Flaws)
 
book.close()

print("Excel file: " + outxls + " built correctly!")

print("If any functions have not been executed, copy the column, paste it in a text file, copy it and paste it back into the excel file!")