import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *
import datetime as dt

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

#1
while True:
    directory = single_input('Input directory with ProductData.xlsx, MasterInvoice.pdf, and order files')
    if directory[-1] == '\\':
        directory = directory
    else:
        directory = directory + '\\'
    
    try:
        filenames = os.listdir(directory)

        if ('MasterInvoice.pdf' not in filenames) or ('ProductData.xlsx' not in filenames) or ('Order1.xlsx' not in filenames) or ('Order2.xlsx' not in filenames)or ('Order3.xlsx' not in filenames)or ('Order4.xlsx' not in filenames): 
            show_message('Error',"Directory doesn't have correct files. Try again.")
            
        else:
            break
    except:
        show_message('Error','Not a valid directory. Try again.')

#2
df1 = pd.read_excel(directory + 'ProductData.xlsx',sheet_name='Descriptions')
df2 = pd.read_excel(directory + 'ProductData.xlsx',sheet_name='Prices')

product_descriptions = {}
product_prices = {}

for index, row in df1.iterrows():
    unit = row['Unit number']
    desc = row['Description']
    product_descriptions[unit] = desc

for index, row in df2.iterrows():
    unit = row['Unit number']
    price = row['Price per unit']
    product_prices[unit] = price

#3
def get_order_data(invoice_num, filename, directory):
    
    invoice = openpyxl.load_workbook(directory + filename,)
    selected_sheet = invoice['OrderForm']

    date1 = selected_sheet['C9'].value
    year = date1.year
    month = date1.month
    day = date1.day
    date = f'{month}/{day}/{year}'

    address1 = selected_sheet['C16'].value
    city = selected_sheet['C17'].value
    state = selected_sheet['C18'].value
    zip = selected_sheet['C19'].value
    address2 = f'{city}, {state} {zip}'

    name = selected_sheet['C14'].value
    comp = selected_sheet['c15'].value

    unit0 = selected_sheet['B24'].value
    price0 = product_prices[unit0]
    desc0 = product_descriptions[unit0]
    quant0 = selected_sheet['C24'].value
    amount0 = price0 * quant0
    price0 = f'${price0:,.2f}'


    if selected_sheet['B25'].value is not None:

        unit1 = selected_sheet['B25'].value
        price1 = round(product_prices[unit1],2)
        quant1 = selected_sheet['C25'].value
        desc1 = product_descriptions[unit1]
        amount1 = round(price1 * quant1,2)
        price1 = f'${price1:,.2f}'

    else:
        amount1 = 0

    if selected_sheet['B26'].value is not None:

        unit2 = selected_sheet['B26'].value
        price2 = product_prices[unit2]
        desc2 = product_descriptions[unit2]
        quant2 = selected_sheet['C26'].value
        amount2 = price2 * quant2
        price2 = f'${price2:,.2f}'

    else:
        amount2 = 0

    if selected_sheet['B27'].value is not None:
        unit3 = selected_sheet['B27'].value
        price3 = product_prices[unit3]
        desc3 = product_descriptions[unit2]
        quant3 = selected_sheet['C27'].value
        amount3 = price3 * quant3
        price3 = f'${price3:,.2f}'

    else:
        amount3 = 0

    total = amount0 + amount1 + amount2 + amount3
    total = f'${total:,.2f}'

    amount0 = f'${amount0:,.2f}'

    if amount1 == 0:
        amount1 = ''
    else:
        amount1 = f'${amount1:,.2f}'

    if amount2 == 0:
        amount2 = ''
    else:
        amount2 = f'${amount2:,.2f}'

    if amount3 == 0:
        amount3 = ''
    else:
        amount3 = f'${amount3:,.2f}'

    update_dict = {}
    update_dict['InvoiceNum'] = invoice_num
    update_dict['Date'] = date
    update_dict['CustomerAddress1'] = address1
    update_dict['CustomerAddress2'] = address2
    update_dict['CustomerName'] = name
    update_dict['CustomerCompany'] = comp
    update_dict['Total'] = total

    update_dict['Price0'] = price0
    update_dict['Quantity0'] = quant0
    update_dict['Amount0'] = amount0
    update_dict['Description0'] = desc0

    if amount1 != '':
        update_dict['Price1'] = price1
        update_dict['Quantity1'] = quant1
        update_dict['Amount1'] = amount1
        update_dict['Description1'] = desc1

    if amount2 != '':
        update_dict['Price2'] = price2
        update_dict['Quantity2'] = quant2
        update_dict['Amount2'] = amount2
        update_dict['Description2'] = desc2

    if amount3 != '':
        update_dict['Price3'] = price3
        update_dict['Quantity3'] = quant3
        update_dict['Amount3'] = amount3
        update_dict['Description3'] = desc3

    email = selected_sheet['C13'].value

    return update_dict, name, email, date

order_data, name, email, date = get_order_data(8515, "Order1.xlsx", directory)

#4
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', f'{directory}Invoice {invoice_num}.pdf', order_data)

create_invoice(8515, order_data, directory)

#5
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()
gmail_user = "ethanrsmith45@gmail.com"

def send_email(invoice_num, name, email, date, directory):

    recipients = [email]

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    body = f"""Dear {name},
    
Attached is your invoice dated {date}.

Thank you for your business.
    
PythonTech Co.
10 Python Road
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    path_to_file = f'{directory}Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message) 
        
    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    mail_server.send_message(message)
    mail_server.quit()

#6
filenames = os.listdir(directory)
filenames

invoice_num = 8515
for file in filenames:
    if 'Order' in file:
        order_data, name, email, date = get_order_data(invoice_num, file, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num,name,email,date,directory)
        invoice_num = invoice_num + 1
