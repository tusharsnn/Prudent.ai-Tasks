import re
import json
import xlwt
from xlwt import Workbook
  
# saves data to corresponding workbook(wb)
def save_data(wb: str, date: str, des: str, amount: str, row: int):
    mon, day, year = date.split("/")
    wb.write(row,0, date)
    wb.write(row,1, des)
    wb.write(row,2, amount)
    wb.write(row,3, day)
    wb.write(row,4, mon)
    wb.write(row,5, "20"+year)


# initializes excel work book to store data in three sheets 
def init_sheets(wb):
    withdrawals_sheet = wb.add_sheet("WITHDRAWALS")
    deposits_sheet = wb.add_sheet("DEPOSITS")
    insights_sheet = wb.add_sheet("INSIGHTS")

    # stores headers at the top for withdrawals sheet
    withdrawals_sheet.write(0,0, "DATE")
    withdrawals_sheet.write(0,1, "DESCRIPTION")
    withdrawals_sheet.write(0,2, "AMOUNT")
    withdrawals_sheet.write(0,3, "DAY")
    withdrawals_sheet.write(0,4, "MONTH")
    withdrawals_sheet.write(0,5, "YEAR")

    # stores headers at the top for deposits sheet
    deposits_sheet.write(0,0, "DATE")
    deposits_sheet.write(0,1, "DESCRIPTION")
    deposits_sheet.write(0,2, "AMOUNT")
    deposits_sheet.write(0,3, "DAY")
    deposits_sheet.write(0,4, "MONTH")
    deposits_sheet.write(0,5, "YEAR")

    # stores headers at the top for insights sheet
    insights_sheet.write(0,0, "key")
    insights_sheet.write(0,1, "value")

    # returns reference of three sheets
    return withdrawals_sheet, deposits_sheet, insights_sheet


# reading data from json array
with open("task_input_list.json", "r") as f:
    data = json.load(f)

# re patterns for capturing data
site_pattern = r"(http(s)?://www.)?([A-Za-z])+([\w])*((\.com)|(\.in))"
email_pattern = r"[a-zA-Z]+[\w]*@(([A-Za-z])+\.(\w)+)"
amount_pattern = r"(-)?(\$)?((\d){1,2},)?(\d)+(\.)(\d)+"
phone_pattern = r"(((\(\d{3}\) )|(\d-\d{3}-))\d{3}-\d{4})|(\+(\d{1,2,3} \d{10}))"
date_pattern  = r"\d{2}/\d{2}/\d{2}"

# data structure to hold various data
sites = set()
emails= set()
amounts = list()
phones = set()

# Workbook is created
wb = Workbook()

# creating sheets
withdrawals_sheet, deposits_sheet, insights_sheet = init_sheets(wb)

# row number to write data on new row
ds_row, ws_row = 1,1

i=0
while True:
    try:
        item = data[i]
    except:
        # break when no more data
        break

    # extracting data
    site_match = re.search(site_pattern, item)
    email_match = re.search(email_pattern, item)
    phone_match = re.search(phone_pattern, item)
    date_match = re.search(date_pattern, item)

    # transactions
    if date_match:
        date = date_match.group()
        i+=1
        description = ""

        # capturing all descriptions from different lines
        while True:
            try:
                item = data[i]
            except:
                break

            amount_match = re.search(amount_pattern, item)

            # amount found
            if not amount_match is None:
                amount = amount_match.group()
                # if any dollar sign then its balance amount
                if amount[0]!="$" and amount[1]!="$":
                    # string to float
                    amount = float(amount.replace(",", ""))
                    amounts.append(amount)

                    # writing to deposit when amount is positive
                    if amount>=0.0:
                        save_data(wb=deposits_sheet, date=date, des=description, amount=amount, row=ds_row)
                        ds_row+=1

                    # writing to withdrawal when amount is negative
                    else:
                        save_data(wb=withdrawals_sheet, date=date, des=description, amount=amount, row=ws_row)
                        ws_row+=1
                    break
            
            # amout not found, data must be part of description
            else:
                description += item
            
            i+=1
        
        pass

    # sites
    if not site_match is None:
        sites.add(site_match.group())
    # emails
    if not email_match is None:
        emails.add(email_match.group())
    # phone numbers
    if not phone_match is None:
        phones.add(phone_match.group())

    i+=1

# writing other insights to excel sheet
insights_sheet.write(2,0, "email")
if not len(emails):
    insights_sheet.write(2,1, "NA")
else:
    email_data = ""
    for email in emails:
        email_data += ", "+email
    insights_sheet.write(2,1, email_data.strip(","))

insights_sheet.write(3,0, "phone_numbers")
if not len(phones):
    insights_sheet.write(3,1, "NA")
else:
    phone_data = ""
    for phone in phones:
        phone_data += ", "+ phone
    insights_sheet.write(3,1, phone_data.strip(","))
    
insights_sheet.write(1,0, "website")
if not len(sites):
    insights_sheet.write(1,1, "NA")
else:
    site_data=""
    for site in sites:
        site_data += ", "+ site
    insights_sheet.write(1,1, site_data.strip(","))

# max amount and min amount if there are any amounts
insights_sheet.write(4,0, "max amount")
insights_sheet.write(5,0, "min amount")
if len(amounts):
    insights_sheet.write(4,1, max(amounts))
    insights_sheet.write(5,1, min(amounts))
else:
    insights_sheet.write(4,1, "NA")
    insights_sheet.write(5,1, "NA")

# saving excel
wb.save('task_output.xls')

