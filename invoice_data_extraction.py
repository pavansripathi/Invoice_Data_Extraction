import pdfplumber
import re
import json
import os
import pandas as pd

# Looping through the files
files = os.listdir(r'A:\Projects\Invoice Data Extraction\Inputs')
overall = {}
count = 1
Total_amount = {}
order_date = {}
order_number = {}
gst_num = {}
biller_name = {}
file_name = {}


class Extract:
    def __init__(self, nfile):
        self.file = nfile

    def dict_create(self, num, line_list):
        for line in text.split('\n'):
            # Searching for Total amount of the purchase
            if line.startswith('TOTAL:'):
                pattern = r'(\d*,?\d+.?\d*)$'
                match = re.search(pattern, line)
                dicts['Total_amount'] = [match.group()]
                Total_amount[num] = match.group()

            # Searching for Order date
            if line.startswith('Order Date:'):
                pattern = r'\d\d.\d\d.\d\d\d\d'
                match = re.search(pattern, line)
                dicts['order_date'] = [match.group()]
                order_date[num] = match.group()

            # Searching for Order Number
            if line.startswith('Order Number:'):
                pattern = r'\d{3}-\d{7}-\d{7}'
                match = re.search(pattern, text)
                dicts['order_number'] = [match.group()]
                order_number[num] = match.group()

            # Searching for GST number
            if line.startswith('GST Registration No:'):
                pattern = r'\w{15}'
                match = re.search(pattern, text)
                dicts['gst_num'] = [match.group()]
                gst_num[num] = match.group()

            # Searching for person name
            for i in range(len(line_list)):
                if line_list[i].endswith('Shipping Address :'):
                    if line_list[i+1].startswith('GST Registration No:'):
                        name = line_list[i+1][37:]
                    elif line_list[i+1].startswith('PAN No:'):
                        name = line_list[i+1][18:]
                    else:
                        name = line_list[i+1]
                    dicts['Name'] = name
                    biller_name[num] = name
                    dicts['File name'] = self.file
                    file_name[num] = self.file


# Creating a Json file
    def json_create(self):
        # Creating a json file
        json_file = open(rf'A:\Projects\Invoice Data Extraction\Outputs\{self.file}.json', mode='w')
        json.dump(dicts, json_file)
        json_file.close()

# Creating an Xlsx file
    def xlsx_create(self):
        # Creating an xlsx file
        df = pd.read_json(rf'A:\Projects\Invoice Data Extraction\Outputs\{self.file}.json')
        df.to_excel(rf'A:\Projects\Invoice Data Extraction\Outputs\{self.file}.xlsx')


for file in files:
    dicts = {}

    # Opening the PDF file
    with pdfplumber.open(rf'A:\Projects\Invoice Data Extraction\Inputs\{file}') as pdf:
        new_file = pdf.pages[0]
        text = new_file.extract_text()
    file = file[0:-4]
    obj = Extract(file)

    line_list = [x for x in text.split('\n')]
    obj.dict_create(count, line_list)
    obj.json_create()
    obj.xlsx_create()
    count += 1


# Adding details of all the files into a seperate Xlsx file
overall['Total_Amount'] = Total_amount
overall['Order_Date'] = order_date
overall['Order_Number'] = order_number
overall['GST_Number'] = gst_num
overall['Biller_Name'] = biller_name
overall['File_Name'] = file_name

# Adding details of all the files into a seperate Json file
json_file = open(r'A:\Projects\Invoice Data Extraction\Outputs\overall.json', 'w')
json.dump(overall, json_file)
json_file.close()

df = pd.read_json(r'A:\Projects\Invoice Data Extraction\Outputs\overall.json')
df.to_excel(r'A:\Projects\Invoice Data Extraction\Outputs\overall.xlsx')
