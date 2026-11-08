from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import os
from dotenv import load_dotenv
import pandas as pd
from openpyxl.chart import BarChart, Reference
# from openpyxl.chart import ...: Because Pandas doesn't 
# know how to draw charts natively, 
# we pull in explicit chart tools directly from our openpyxl library.
# BarChart: The blueprint for creating a bar/column graph.
# Reference: A tool used to tell Python which specific cells 
# on the Excel grid the chart should look at for its data.

# Load environment variables from the .env file
load_dotenv()

data = {
    'Region': ['North', 'South', 'East', 'West', 'North', 'South', 'East', 'West'],
    'Salesperson': ['Alice', 'Bob', 'Charlie', 'Diana', 'Eve', 'Frank', 'Grace', 'Hank'],
    'Product': ['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'Laptop', 'Monitor', 'Keyboard', 'Mouse'],
    'Units_Sold': [15, 45, 30, 10, 25, 12, 40, 50],
    'Revenue': [15000, 1350, 2100, 3000, 25000, 3600, 2800, 1500]
}
# A DataFrame is basically a supercharged, programmable Excel table stored in Python's memory. This is the exact format we would get if we pulled data from a real SQL Server!
df = pd.DataFrame(data)

# 2. CREATE A SUMMARY / PIVOT TABLE (Grouping data by Region)
# This sums up the Revenue and Units Sold for each Region
# df.groupby('Region'): This groups our rows together by the Region column (just like dragging a field into the "Rows" area of an Excel Pivot Table).
# [['Units_Sold', 'Revenue']]: This tells Python which specific columns we want to calculate.
# .sum(): This adds up all the units sold and revenue for each region.
# .reset_index(): This cleans up the table layout so Region stays a normal column instead of turning into an index label.
summary_df = df.groupby('Region')[['Units_Sold', 'Revenue']].sum().reset_index()

print("Summary Data (Pivot Table style):")
print(summary_df)

# len(df) counts how many rows are in our table.
# df.head() grabs just the first 5 rows of the table so you can preview it without flooding your screen with too much text
print("\n\nSuccessfully generated mock data!")
print(f"Total rows: {len(df)}")
print("\nData Preview:")
print(df.head())

# 3. Export the DataFrame directly to an Excel file
file_name = 'sales_report.xlsx'

# # index=False tells pandas not to save the boring row numbers on the far left
# df.to_excel(file_name, sheet_name='Sales Summary', index=False)
# 'ExcelWriter' lets us write more than one table into different tabs of the same file
# ExcelWriter acts as a multi-sheet container. The with statement ensures the file safely opens, writes our data, and automatically closes when finished.

with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Raw Data', index=False)
    summary_df.to_excel(writer, sheet_name='Region Summary', index=False)
    # New day:-
    # Grabs the active 
    # Excel workbook that our ExcelWriter is currently building.
    sheets = writer.book
    # Pulls out just the 'Region Summary' sheet
    worksheet = writer.sheets['Region Summary']
    chart = BarChart()
    chart.type = 'col'
    chart.title = 'Total Revenue by Region'
    # Applies a built-in professional color theme so it doesn't look completely plain.
    chart.style = 10
    chart.y_axis.title = 'Revenue ($)'
    chart.x_axis.title = 'Region'
    # This tells the chart where the numbers are.
    data_ref = Reference(
        worksheet, min_col=3, min_row=1, max_col=3, max_row=len(summary_df) + 1
    )
    # This tells the chart what text labels to put on the bottom of the graph.
    cats_ref = Reference(
        worksheet, min_col=1, min_row=2, max_col=1, max_row=len(summary_df) + 1
    )
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    worksheet.add_chart(chart, 'E2')

print(f"\nSuccess! Your Excel report has been generated: {file_name}")

# Phase 4 the email automation
sender_email = "mbazil540@gmail.com"
receiver_email = "afj084@gmail.com"
password = os.getenv("ETLpass")

# Initializes our blank email envelope.
msg = MIMEMultipart()
msg['From'] = sender_email
msg['To'] = receiver_email
msg['Subject'] = 'Automated Sales Report - Region Summary'

body = "Hello User, Here is your ETL pipeline report"
# Inserts our text message into the email body as plain text.
msg.attach(MIMEText(body, 'plain'))

with open(file_name, 'rb') as attachment:
    #   builds a specialized container box with a label that says "Raw Binary File Inside".
    part = MIMEBase('application', 'octet-stream')
    #  takes the contents of your sales_report.xlsx file and packs them right inside that box.
    part.set_payload(attachment.read())

# attachment.read(): Remember earlier we opened our file with open(file_name, 'rb'). 
# The .read() part literally reads all the raw data (all the bytes) out of that Excel file from your computer's hard drive.

# part.set_payload(...): This takes all those raw bytes we just read 
# and places them inside our shipping box (part) that we created in Line 1.
encoders.encode_base64(part)
part.add_header('Content-Disposition', f'attachment; filename= {file_name}')
msg.attach(part)

try:
  print('Connecting to email server...')
  server = smtplib.SMTP('smtp.gmail.com', 587)
#   encrypt the connection
  server.starttls()
  server.login(sender_email, password)
  server.sendmail(sender_email, receiver_email, msg.as_string())
  server.quit()
  print('Email sent successfully to stakeholders!')
except Exception as e:
  print(f'Failed to send email: {e}')
