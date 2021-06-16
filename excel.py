import glob 
import pandas as pd 
import xlswriter 
import openpyxl
from pandas import series 
import os 
import time 
import numpy as np

wb = openpyxl.load_workbook("Stocks.xlsx")
ws = wb.worksheets[0]
ws_tables = []

for table in ws._tables:
    ws_tables.append(table)
    print(table.name, table.ref)
