import pandas as pd
import gspread
import smtplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
import os
import openpyxl
from openpyxl.styles import Alignment

# Csv file downloaded from PowerBi
pbi_csv = r"Path to csv"
pbi_data= pd.read_csv(pbi_csv)

poc_csv= r"Path to csv"
poc_data= pd.read_csv(poc_csv)

req_pbi = pbi_data.loc[:, ["Company name", "Last 10 Days Avg Redemption(INR)","Redemption Wallet Balance(INR)","Redeemable Balance(INR)"]]
req_poc=poc_data.loc[:, ["Company name","Sales POC", "Sales Email ID"]]

merged= pd.merge(req_pbi, req_poc, on="Company name",how="left")

# Iterate over the merged DataFrame and create the dictionary
sales_email_id_dict = {}
sales_email_id_list=[]
for index, row in merged.iterrows():
    sales_email_id = row['Sales Email ID']
    company_name = row['Company name']
    # Create a dictionary of company details
    company_details = {
        'Last 10 Days Avg Redemption(INR)': row['Last 10 Days Avg Redemption(INR)'],
        'Redemption Wallet Balance(INR)': row['Redemption Wallet Balance(INR)'],
        'Redeemable Balance(INR)': row['Redeemable Balance(INR)'],
    }
    # Check if the sales email ID is already in the dictionary
    if sales_email_id in sales_email_id_dict:
        sales_email_id_dict[sales_email_id]['Companies'].append((company_name, company_details))
    else: 
        sales_email_id_dict[sales_email_id] = {
            'Companies': [(company_name, company_details)]
        }
        sales_email_id_list.append(sales_email_id)

def expected_value(x, y, z):
    result = []
    for i in range(len(x)):
        if y[i] < 0:
            result.append(abs(y[i]) + z[i])
        else:
            result.append(max(x[i], y[i], z[i]))
    return result
    
# Export the dictionaries to Excel files
for sales_email_id, data in sales_email_id_dict.items():
    final_df = pd.DataFrame({'Company Name': [x[0] for x in data['Companies']],
                           **{key: [x[1][key] for x in data['Companies']] for key in ['Last 10 Days Avg Redemption(INR)', 'Redemption Wallet Balance(INR)', 'Redeemable Balance(INR)']}})

    #expected recharge value calculation
    final_df["Expected recharge value (INR)"] = expected_value(
        final_df['Last 10 Days Avg Redemption(INR)'],
        final_df['Redemption Wallet Balance(INR)'],
        final_df['Redeemable Balance(INR)'])

    #Creating excel files 
    excel_name=(str(sales_email_id)).split('@')[0]
    output_file = f"Path to the folder where the individual excel sheets will be generated\\{excel_name}.xlsx"
    final_df.to_excel(output_file, sheet_name='Company_details', index=False)

    #text wrapping
    workbook = openpyxl.load_workbook(output_file)
    worksheet = workbook['Company_details']
    for row in worksheet.iter_rows():
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(wrap_text = True)
    workbook.save(output_file)

#send out the mails 
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
sender_mail= "****"
password= "****"
server.login(sender_mail, password)
for email_id in sales_email_id_list:
    poc_excel=(str(email_id)).split('@')[0]
    msg = MIMEMultipart()
    msg["From"] = f"XYZ <{sender_mail}>"
    msg["To"] = str(email_id)
    msg["Subject"] = "Balance is running low!"
    msg["Cc"] = "abc@gmail.com, xyz@gmail.com, ghj@gmail.com"
    body = f"Hi {poc_excel},\nPlease reach out to the client"
    msg.attach(MIMEText(body, "plain"))
    company_sheet = MIMEBase('application', 'octet-stream')     
    attachment_file= f"Path to the folder where the individual excel sheets are generated\\{poc_excel}.xlsx"
    company_sheet.set_payload(open(attachment_file, 'rb').read())
    email.encoders.encode_base64(company_sheet)
    company_sheet.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(attachment_file)}"')
    msg.attach(company_sheet)
    server.sendmail(sender_mail, [str(email_id), 'abc@gmail.com', 'xyz@gmail.com', 'ghj@gmail.com'], msg.as_string())
server.quit()


