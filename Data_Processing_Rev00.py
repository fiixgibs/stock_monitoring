import pandas as pd
import requests
import io
import base64
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib

# Step 1: Define the function to convert OneDrive share link to direct download link
def create_onedrive_directdownload(onedrive_link):
    data_bytes64 = base64.b64encode(bytes(onedrive_link, 'utf-8'))
    data_bytes64_string = data_bytes64.decode('utf-8').replace('/', '_').replace('+', '-').rstrip("=")
    result_url = f"https://api.onedrive.com/v1.0/shares/u!{data_bytes64_string}/root/content"
    return result_url

# Step 2: Define OneDrive share links for your files
erp_file_link = "https://1drv.ms/x/s!AmKK1G3zXcbEbjC0wAbMfCREMOU?e=ykM0ex"
whs_file_link = "https://1drv.ms/x/s!AmKK1G3zXcbEbdkGNXzqP_Ylm6o?e=eTE1pg"

# Step 3: Convert the OneDrive share links to direct download links
erp_download_link = create_onedrive_directdownload(erp_file_link)
whs_download_link = create_onedrive_directdownload(whs_file_link)

# Step 4: Function to download the files from OneDrive using the direct download link
def download_file_from_onedrive(link):
    response = requests.get(link)
    if response.status_code == 200:
        print("File downloaded successfully!")
        return io.BytesIO(response.content)
    else:
        raise Exception(f"Failed to download file. Status code: {response.status_code}")

# Step 5: Download both files from OneDrive
erp_data_file = download_file_from_onedrive(erp_download_link)
whs_data_file = download_file_from_onedrive(whs_download_link)

# Step 6: Load the Excel data into pandas DataFrames (specifying the engine)
erp_data = pd.read_excel(erp_data_file, engine='openpyxl')  # Present Quantity data
whs_data = pd.read_excel(whs_data_file, engine='openpyxl', skiprows=5)  # Maintenance data

# Step 7: Merge the two datasets based on 'Resource' in erp_data and 'Item Code' in whs_data
# Also include 'Class' from whs_data (from 102WHS) in the merged data
merged_data = pd.merge(erp_data, whs_data, how='inner', left_on='Resource', right_on='Item Code')

# Step 8: Use correct column names for filtering
filtered_data = merged_data[merged_data['Stock'] < merged_data['Minimum \nAmount']]

# Step 9: Select relevant columns for your final report, including 'Class'
final_report = filtered_data[['Item Code', 'Stock', 'Minimum \nAmount', 'Description', 'Class']]

# Step 9.5: Rename 'Minimum \nAmount' to 'Minimum Amount' for clarity
final_report.rename(columns={'Minimum \nAmount': 'Minimum Amount'}, inplace=True)

# Step 9.6: Add the 'to order' column
final_report['to order'] = final_report['Minimum Amount'] - final_report['Stock']
final_report['to order'] = final_report['to order'].apply(lambda x: x if x > 0 else 0)  # Ensure no negative values

# Step 10: Save the final report to a new Excel file
final_report_file = 'Low_Stock_Report.xlsx'
final_report.to_excel(final_report_file, index=False)

print("Data processed and Low Stock Report generated successfully!")

# Step 12: Function to send email with attachment via Outlook
def send_email_with_outlook(to_email, subject, body, file_path):
	from_email = "fix.gibs.banting@gmail.com"  # Your Outlook email
	password = "nsyh kspa jevv rhbg"  # Your Outlook App Password (if 2FA is enabled)

	# Set up the email
	msg = MIMEMultipart()
	msg['From'] = from_email
	msg['To'] = to_email
	msg['Subject'] = subject

	# Add the email body
	msg.attach(MIMEText(body, 'plain'))

	# Attach the file
	with open(file_path, "rb") as attachment:
			part = MIMEBase('application', 'octet-stream')
			part.set_payload(attachment.read())
	encoders.encode_base64(part)
	part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(file_path)}")
	msg.attach(part)

	# Connect to Gmail's SMTP server
	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.starttls()  # Upgrade connection to secure
	server.login(from_email, password)

	# Send the email
	text = msg.as_string()
	server.sendmail(from_email, to_email, text)
	server.quit()      

# Example usage:
file_path = final_report_file
# file_path = 'Low_Stock_Report.xlsx' -- or this one, if above didnt work
send_email_with_outlook("GIBSmaintenance@gamuda-ibs.com.my", "Low Stock Report", "Please find attached the latest low stock report.", file_path)