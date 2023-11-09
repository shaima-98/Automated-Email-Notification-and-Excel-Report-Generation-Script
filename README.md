# Automated Email Notification and Excel Report Generation Script

1. Overview:
   
This script automates the process of generating Excel reports and sending email notifications to sales representatives based on data from Power BI and CSV files. It uses the SMTP protocol to send email notifications to sales representatives. It logs in using a sender's email and password, then iterates through the list of sales representatives' email IDs. Each email includes a personalized message and an attached Excel report.

3. Requirements:
   
- Python 3
- pandas
- gspread
- smtplib
- email
- openpyxl
- os

5. Code Structure:
   
- Importing Libraries
- The required libraries are imported for data manipulation, Excel handling, and email functionalities.
- Data Loading
- Power BI and CSV files are loaded into Pandas DataFrames.
- Data Merging
- Relevant columns from both DataFrames are selected and merged on the "Company name" column.
- Sales Email ID Dictionary Creation
- The script iterates through the merged DataFrame and creates a dictionary (sales_email_id_dict) that organizes data by sales representatives' email IDs.
- Expected Value Calculation Function
- The function expected_value calculates the expected recharge value based on specific conditions.
- Export to Excel Files
- The script exports data to Excel files for each sales representative, including the calculated expected recharge value. Text wrapping is applied to enhance readability.
