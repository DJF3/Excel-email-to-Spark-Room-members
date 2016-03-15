import json
import requests
import openpyxl
import re

TestOnly = "no" 

if TestOnly.lower() == "yes" :
    for i in range(mysheetStartRow, mysheethighestrow, 1):
        EmailFound = sheet.cell(row=i, column=mysheetStartCol).value
        if not re.match(r"[^@]+@[^@]+\.[^@]+", EmailFound):
            print "--- No valid email address found:", EmailFound
            continue
        print EmailFound
