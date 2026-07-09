import smtplib
from email.message import EmailMessage
import pandas as pd
import os
import time
import random
from dotenv import load_dotenv

# Load environment variables from a .env file
load_dotenv()

def send_emails(excel_path, sender_email, app_password, email_subject, email_body_template, resume_path):
    # Read the Excel file
    try:
        df = pd.read_excel(excel_path)
        # Standardize column names to lowercase to prevent case-sensitivity issues
        df.columns = df.columns.str.lower()
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    # Check if resume exists
    if not os.path.exists(resume_path):
        print(f"Error: Resume file not found at {resume_path}")
        return

    # Set up the SMTP server
    smtp_server = "smtp.gmail.com"
    smtp_port = 465 # SSL port
    
    # Load already sent emails from a log file so we can resume if interrupted
    sent_log_file = "sent_emails_log.txt"
    sent_emails = set()
    if os.path.exists(sent_log_file):
        with open(sent_log_file, "r") as f:
            for line in f:
                sent_emails.add(line.strip())
    
    try:
        # Connect to the server
        print("Connecting to the email server...")
        server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        server.login(sender_email, app_password)
        print("Successfully logged in.")
        
        sent_this_session = 0
        max_to_send = 400

        for index, row in df.iterrows():
            if sent_this_session >= max_to_send:
                print(f"\nReached the limit of {max_to_send} emails for this session. Stopping.")
                break
            # Extract data from the Excel row
            # Adjust these column names to match exactly what is in your Excel file
            recipient_email = row.get('email')
            recipient_name = row.get('name')
            recipient_role = row.get('title')
            company_name = row.get('company')

            # Skip if there's no email or if we already sent to this email
            if pd.isna(recipient_email) or str(recipient_email).strip() == "":
                continue
                
            email_address_str = str(recipient_email).strip()
            if email_address_str in sent_emails:
                print(f"[{index + 1}/{len(df)}] Skipping {email_address_str} (already sent in a previous run)")
                continue

            # Create the email message
            msg = EmailMessage()
            
            # Format subject
            try:
                formatted_subject = email_subject.format(
                    name=recipient_name,
                    title=recipient_role,
                    company=company_name
                )
            except KeyError:
                formatted_subject = email_subject
            msg['Subject'] = formatted_subject
            
            msg['From'] = sender_email
            msg['To'] = recipient_email

            # Personalize the email body
            # We replace placeholders like {name}, {title}, {company} in the template
            personalized_body = email_body_template.format(
                name=recipient_name,
                title=recipient_role,
                company=company_name
            )
            msg.set_content(personalized_body)

            # Attach the resume
            with open(resume_path, 'rb') as f:
                resume_data = f.read()
                resume_name = os.path.basename(resume_path)
                # Assuming the resume is a PDF. If it's a docx, change the maintype and subtype.
                msg.add_attachment(resume_data, maintype='application', subtype='pdf', filename=resume_name)

            # Send the email
            try:
                server.send_message(msg)
                print(f"[{index + 1}/{len(df)}] Successfully sent email to {recipient_name} at {company_name} ({email_address_str})")
                
                # Record successful send in the log file
                with open(sent_log_file, "a") as f:
                    f.write(email_address_str + "\n")
                sent_emails.add(email_address_str)
                
                sent_this_session += 1
                
            except Exception as send_err:
                print(f"[{index + 1}/{len(df)}] Failed to send to {email_address_str}. Error: {send_err}")
            
            # Add a fixed 2-second delay between each mail
            print("Waiting for 2.00 seconds before next email...")
            time.sleep(2)

    except Exception as e:
        print(f"An error occurred with the email server: {e}")
    finally:
        try:
            server.quit()
        except:
            pass

if __name__ == "__main__":
    # --- Configuration ---
    
    # 1. Path to your Excel file
    EXCEL_FILE_PATH = "contacts.xlsx" 
    
    # 2. Your credentials (now loaded safely from the .env file)
    MY_EMAIL = os.getenv("SENDER_EMAIL")
    APP_PASSWORD = os.getenv("APP_PASSWORD") 
    
    if not MY_EMAIL or not APP_PASSWORD:
        print("Error: Missing credentials. Please make sure you have a .env file with SENDER_EMAIL and APP_PASSWORD.")
        exit(1)
    
    # 3. Email details
    # You can use placeholders {company}, {name} in the subject as well
    EMAIL_SUBJECT = "Software Engineer Opportunity at {company}"
    
    # The email body. Use {name} and {company} where you want the script to insert the Excel data.
    EMAIL_BODY = """Dear {name},

I hope you're doing well.

I came across your profile and wanted to reach out regarding Software Engineering opportunities at {company}.
I am currently working as a Junior Software Engineer at CronLabs Solutions, where I am building scalable healthcare application using React, Node.js, Prisma, SQL, Django, Celery, and RabbitMQ. 
Currently, I am working on MyMedScreen, a DOT medical certification platform, where I develop scalable backend services, implement complex workflow automation, and build asynchronous processing using Django, Celery, RabbitMQ, React, and SQL.
I also have experience building full-stack MERN applications and have solved 350+ DSA problems on LeetCode and GeeksforGeeks.
If your team is hiring, I would be grateful if you could consider my profile or refer me to the appropriate hiring team. 
I've attached my resume for your review. Thank you for your time, and I look forward to hearing from you.


Best regards,

Anuj Mishra
+91-9229409080
anujm8918@gmail.com
GitHub: github.com/anuj8918
LeetCode: leetcode.com/u/anujm8918
"""
    
    # 4. Path to your resume
    RESUME_FILE_PATH = "AnujMishraResume.pdf"
    
    # --- Run the function ---
    print("Starting the mass email process...")
    send_emails(EXCEL_FILE_PATH, MY_EMAIL, APP_PASSWORD, EMAIL_SUBJECT, EMAIL_BODY, RESUME_FILE_PATH)
    print("Process finished!")
