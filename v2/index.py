import pandas as pd
import requests
import json
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import numbers
import time

print("Queuing Update...")
file_path = "Stocks.xlsx"
wb = load_workbook(file_path)
companyList = pd.read_excel(file_path, sheet_name="CompanyList", header=None)
for company in companyList[0]:
    print("Featching data for ", company)
    companySheet = pd.read_excel(file_path, sheet_name=company)
    api_url = "https://www.investorslounge.com/Default/SendPostRequest"
    payload = {
        "url": "PriceHistory/GetPriceHistoryCompanyWise",
        "data": json.dumps(
            {
                "company": company,
                "sort": "0",
                "DateFrom": datetime.now().strftime("%d %b %Y"),
                "DateTo": datetime.now().strftime("%d %b %Y"),
                "key": "",
            }
        ),
    }

    response = requests.post(api_url, json=payload)
    if len(response.json()) > 0:
        data = response.json()[0]
        ws = wb[company]
        ws.move_range("G2:M2", rows=1)
        max_column = ws.max_column
        for col in range(7, max_column + 1):
            if ws.cell(row=3, column=col).data_type == "f":
                ws.cell(row=2, column=col).value = ws.cell(row=3, column=col).value
        ws["G2"].value = datetime.strptime(data["Date_"], "%Y-%m-%dT%H:%M:%S")
        ws["H2"].value = float(data["Open"])  # Assuming Open is a float
        ws["I2"].value = float(data["High"])  # Assuming High is a float
        ws["J2"].value = float(data["Low"])  # Assuming Low is a float
        ws["K2"].value = float(data["Close"])  # Assuming Close is a float
        ws["L2"].value = int(data["Volume"])  # Volume should be an int
        ws["M2"].value = float(data["Change"])  # Change should be a float
        ws["M2"].number_format = "0.00"
        print('Successful')
    else:
        print('Failed')


try:
    wb.save(file_path)
    print(f"Update complete for all companies")
except Exception as e:
    print(f"An error occurred: {e}")
