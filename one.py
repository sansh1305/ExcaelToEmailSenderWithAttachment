import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os

# Path to the Excel file containing client information
path = "./helooooo.xlsx"

workbook = openpyxl.load_workbook(path)
sheet = workbook["Sheet1"]

# Lists to store email addresses, amounts, and client names
mail_list = []
name = []

# Extract relevant data from the Excel sheet
for row in sheet.iter_rows(min_row=2, values_only=True):
    client, email,count_amount = row
    mail_list.append(email)
    name.append(client)



# Gmail credentials
email_address = "sanshita.goyal2022@vitstudent.ac.in"  
password = "aaum eshh xqib xvre"  

# Connect to Gmail server
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(email_address, password)

# Attach the Word document
word_file_path = "helooooo.docx"

# Send personalized emails to clients who haven't paid
for mail_to, clientName in zip(mail_list, name):
    subject = f"{clientName}, you have a new email"
    message = f"Dear {clientName},\n\n" \
              f"We inform you that you owe.\n\n" \
              "Best Regards"

    msg = MIMEMultipart()
    msg["From"] = email_address
    msg["Subject"] = subject
    msg.attach(MIMEText(message, "plain"))

    with open(word_file_path, "rb") as file:
        attachment = MIMEApplication(file.read(), Name=os.path.basename(word_file_path))
        attachment["Content-Disposition"] = f"attachment; filename={os.path.basename(word_file_path)}"
        msg.attach(attachment)

    print(f"Sending email to {clientName}...")
    server.sendmail(email_address, mail_to, msg.as_string())

# Close the server connection
server.quit()
print("Process is finished!")