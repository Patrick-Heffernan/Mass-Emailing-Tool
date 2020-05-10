import openpyxl, ezgmail

wb = openpyxl.load_workbook('Email Addresses.xlsx')
sheet = wb['Sheet1']

emails = {}
for i in range(1,sheet.max_row-1):
    name = sheet.cell(i, 1).value
    email_address = sheet.cell(i,2).value
    emails[name] = email_address

subject = "Testing!"

for k, v in emails.items():
   sendTo = v
   content = "Hello, " + str(k) + ". This is a test"
   ezgmail.send(sendTo, subject, content)



