# Insert your python code here
import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

# Question 1

while True:
    directory = single_input("Enter the folder path.")

    if os.path.exists(directory):

        if not directory.endswith('/'):
            directory = directory + '/'

        filenames = os.listdir(directory)

        has_files = (
            'ProductData.xlsx' in filenames
            and 'MasterInvoice.pdf' in filenames
            and 'Order1.xlsx' in filenames
            and 'Order2.xlsx' in filenames
            and 'Order3.xlsx' in filenames
            and 'Order4.xlsx' in filenames)

        if has_files:
            break
        else:
            show_message("Missing files","The folder does not contain all required files.")

    else:
        show_message("Invalid folder","The folder does not exist. Please try again.")

# Question 2

# INSERT ANSWER FOR QUESTION 2 HERE
df = pd.read_excel(directory+'ProductData.xlsx',sheet_name='Descriptions')

product_descriptions = {}
for index, row in df.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

df1 = pd.read_excel(directory+'ProductData.xlsx',sheet_name='Prices')

product_prices = {}
for index, row in df1.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

show_message('Question 2 Complete','Continue')
# Question 3

# INSERT ANSWER FOR QUESTION 3 HERE
df = get_pdf(directory+ 'MasterInvoice.pdf')
display(df)

def get_order_data(invoice_num,filename,directory):
    path = directory + filename
    workbook = openpyxl.load_workbook(path)
    sheet = workbook['OrderForm']
    dictionary = {}
    dictionary['InvoiceNum'] = str(invoice_num)
    dictionary['Date'] = sheet['C9'].value.strftime('%m/%d/%Y')
    date = sheet['C9'].value.strftime('%m/%d/%Y')
    dictionary['CustomerAddress1'] = sheet['C16'].value
    dictionary['CustomerName'] = sheet['C14'].value
    name = sheet['C14'].value
    email = sheet['C13'].value
    dictionary['CustomerCompany'] = sheet['C15'].value
    dictionary['CustomerAddress2'] = f'{sheet['C17'].value}, {sheet['C18'].value} {sheet['C19'].value}'
    if pd.isna(sheet['B24'].value) == False:
        dictionary['Description0'] = str(product_descriptions[sheet['B24'].value])
        dictionary['Price0'] = str(product_prices[sheet['B24'].value])
        dictionary['Quantity0'] = str(sheet['C24'].value)
        dictionary['Amount0'] = f'${(product_prices[sheet['B24'].value] * sheet['C24'].value):,.2f}'
    if pd.isna(sheet['B25'].value) == False:
        dictionary['Description1'] = str(product_descriptions[sheet['B25'].value])
        dictionary['Price1'] = str(product_prices[sheet['B25'].value])
        dictionary['Quantity1'] = str(sheet['C25'].value)
        dictionary['Amount1'] = f'${(product_prices[sheet['B25'].value] * sheet['C25'].value):,.2f}'
    if pd.isna(sheet['B26'].value) == False:
        dictionary['Description2'] = str(product_descriptions[sheet['B26'].value])
        dictionary['Price2'] = str(product_prices[sheet['B26'].value])
        dictionary['Quantity2'] = str(sheet['C26'].value)
        dictionary['Amount2'] = f'${(product_prices[sheet['B26'].value] * sheet['C26'].value):,.2f}'
    if pd.isna(sheet['B27'].value) == False:
        dictionary['Description3'] = str(product_descriptions[sheet['B27'].value])
        dictionary['Price3'] = str(product_prices[sheet['B27'].value])
        dictionary['Quantity3'] = str(sheet['C27'].value)
        dictionary['Amount3'] = f'${(product_prices[sheet['B27'].value] * sheet['C27'].value):,.2f}'
    if pd.isna(sheet['B28'].value) == False:
        dictionary['Description4'] = str(product_descriptions[sheet['B28'].value])
        dictionary['Price4'] = str(product_prices[sheet['B28'].value])
        dictionary['Quantity4'] = str(sheet['C28'].value)
        dictionary['Amount4'] = f'${(product_prices[sheet['B28'].value] * sheet['C28'].value):,.2f}'

    total = 0

    for row in [24, 25, 26, 27, 28]:
        unit_num = sheet[f'B{row}'].value

        if pd.isna(unit_num) == False:
            quantity = sheet[f'C{row}'].value
            total = total + product_prices[unit_num] * quantity

    dictionary['Total'] = f'${total:,.2f}'
    
    

    return dictionary, name, email, date

dictionary, n, e, d = get_order_data(12831, 'Order1.xlsx', '/Users/dylanrosvall/Documents/ACC 525/17 - Project 3 - Robotic Process Automation/')
print(dictionary, n, e, d)

show_message('Question 3 Complete','Continue')

# Question 4

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf',order_data)

show_message('Question 4 Complete','Continue')


# Question 5

# INSERT ANSWER FOR QUESTION 5 HERE

# gmail_pasword_path = 
gmail_password_path = select_file('Please select the gmail password file')

def send_email(invoice_num, name, email, date, directory):
    # Enter your gmail user name and password

    # gmail_user = 'sales@pythontechco.com'
    gmail_user = "dylanrosvall@gmail.com"
    gmail_password = open(gmail_password_path).read()

    # List of recipients

    recipients = email

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    name = f'{name}'

    body = f"""Dear {name},
        
Attached is your invoice dated {date},

Thank you for your business.
    
PythonTech Co

10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    # path_to_file = os.getcwd()+'\\Invoice Template.xlsx'
    path_to_file = directory + f'Invoice {invoice_num}.pdf' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()


send_email(8515,'Thomas Holt', 'dylanrosvall@gmail.com','11/12/2023',directory)

show_message('Question 5 Complete','Continue')
# Question 6

# INSERT ANSWER FOR QUESTION 6 HERE

filenames = os.listdir(directory)
filenames = sorted(filenames)

invoice_num = 8515

for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num,order_data,directory)
        send_email(invoice_num,name,email,date,directory)
        invoice_num = invoice_num + 1

show_message('Question 6 Complete','Continue')
