#!/usr/bin/env python
# coding: utf-8

# In[127]:

"""
import os
from datetime import datetime

import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


# In[128]:


subject = "An email with attachment from Python"
body = "This is an email with attachment sent from Python"
sender_email = "jose.bustamante@dex.co"
receiver_email = "jose.bustamante@dex.co" #supervisoresbarbosa@duratexsa.onmicrosoft.com
password = "Dexco2024..."


# In[129]:


# Create a multipart message and set headers
message = MIMEMultipart()
message["From"] = sender_email
message["To"] = receiver_email
message["Subject"] = subject
message["Bcc"] = receiver_email  # Recommended for mass emails


# In[130]:


def image_getter(root_folder):
    """
    Finds an image file in the given folder whose name contains today's date.

    Args:
        root_folder (str): The root folder to search in.

    Returns:
        str: Absolute path of the image file if found, otherwise None.
    """
    # Get today's date in the format 'YYYY-MM-DD'
    today_date = datetime.today().strftime('%Y-%m-%d')
    files_found=[] #list to store path of selected files
    # Walk through all files in the root folder and subfolders
    for root, dirs, files in os.walk(root_folder):
        for file in files:
            # Check if the file is an image and contains today's date
            if file.endswith(('.png', '.jpg', '.jpeg')) and today_date in file:
                # Return the absolute path of the image
                files_found.append(os.path.abspath(os.path.join(root, file)))
    # Return None if no matching file is found
    if len(files_found)<1:
        return None
    else:
        return files_found


# In[131]:


# Add body to email
message.attach(MIMEText(body, "plain"))

root_folder = r"C:\Users\JDBUSTAMANTE\OneDrive - Duratex SA\reports_visualizacion_data_produccion\data_plots\imgs_reports_daily"  # Replace with your folder path
filenames = image_getter(root_folder)  # In same directory as script
for filename in filenames:
    # Open file in binary mode
    with open(filename, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    # Encode file in ASCII characters to send by email    
    encoders.encode_base64(part)
    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )
    # Add attachment to message and convert message to string
    message.attach(part)
text= message.as_string()


# In[134]:


# Log in to server using secure context and send email
context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp-mail.outlook.com", 587) as server: #, context=context
    server.login(sender_email, password)
    server.sendmail(sender_email, receiver_email, text)

"""