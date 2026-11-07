
pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

while True:
    directory = single_input("Please input the directory where the file is located.", width=400, height=200)
    directory = directory.strip("'")
    if directory.endswith('/'):
        pass
    else:
        directory = directory + '/'
    file_list = os.listdir(directory)
    if 'ProductData.xlsx' in file_list and 'MasterInvoice.pdf' in file_list and 'Order1.xlsx' in file_list and 'Order2.xlsx' in file_list and 'Order3.xlsx' in file_list and 'Order4.xlsx' in file_list:
        break
    else: 
        show_message('Error: Your directory is missing files.','Please re-enter the directory.')# Insert your python code here

ddf = pd.read_excel(directory+'ProductData.xlsx',sheet_name='Descriptions')

product_descriptions = {}
for index, row in ddf.iterrows():
    product_descriptions[row['Unit number']] = row['Description']
product_descriptions

pdf = pd.read_excel(directory+'ProductData.xlsx',sheet_name='Prices')

product_prices = {}
for index, row in pdf.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']
product_prices

def get_order_data(invoice_num, filename, directory):
    template = openpyxl.load_workbook(directory + filename)
    sheet = template['OrderForm']
    name = sheet['C14'].value
    email = sheet['B7'].value.strip('E-mail:')
    date = sheet['C9'].value.strftime('%m/%d/%Y')
    order_data = {}
    order_data['InvoiceNum'] = invoice_num
    order_data['Date'] = sheet['C9'].value.strftime('%m/%d/%Y')
    order_data['CustomerAddress1'] = sheet['C16'].value
    order_data['CustomerName'] = sheet['C14'].value
    order_data['CustomerCompany'] = sheet['C15'].value
    order_data['CustomerAddress2'] = f'{sheet['C17'].value}, {sheet['C18'].value} {sheet['C19'].value}'
    num = 0
    row = 24
    total = 0
    while True:
          if sheet[f'B{row}'].value == None:
                break
          else: 
            price = product_prices[sheet[f'B{row}'].value] 
            order_data[f'Price{num}'] = str(f'${price:,.2f}')
            quantity = sheet[f'C{row}'].value
            order_data[f'Quantity{num}'] = str(quantity)
            amount = product_prices[sheet[f'B{row}'].value]*sheet[f'C{row}'].value
            order_data[f'Amount{num}'] = str(f'${amount:,.2f}')
            order_data[f'Description{num}'] = product_descriptions[sheet[f'B{row}'].value]
            num = num+1
            row = row+1
            total = total + amount
    order_data['Total'] = f'${total:,.2f}'

    return order_data, name, email, date
    

order_data, name, email, date = get_order_data(8515, filename = 'Order1.xlsx', directory = directory)

get_order_data(8515, 'Order1.xlsx', directory)
invoice_num = 8515

def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory +  f'Invoice {invoice_num}.pdf', order_data)

create_invoice(8515, order_data, directory)

def send_email(invoice_num, name, email, date, directory):
    # Enter your gmail user name and password

    gmail_user = "22dtaylor2@gmail.com"
    gmail_password = open('gmail_password.txt').read()

    # List of recipients

    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    name = name

    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.
        
    Python Tech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory+f'/Invoice {invoice_num}.pdf'
    #path_to_file = os.getcwd()+'/Invoice Template.xlsx' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

send_email(8515, 'Thomas Holt', 'dtaylo18@byu.edu', '11/13/2023', directory)

filenames = os.listdir(directory)
filenames.sort()
for file in filenames:
    invoice_num = 8515
    if 'Order' in file and 'xlsx' in file:
        get_order_data(invoice_num, file, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
    invoice_num += 1


