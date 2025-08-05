import pandas as pd
import smtplib
import ssl
from email.message import EmailMessage
import time

# ================================
# CONFIGURATION SECTION - MODIFY THESE
# ================================

# Email Settings
SENDER_EMAIL = "prabhavgupta05@gmail.com"  # Replace with your email
SENDER_NAME = "Prabhav Gupta"          # Replace with your name
APP_PASSWORD = "password"      # Replace with your Gmail app password
SUBJECT = "Application for Potential Opportunities"
RESUME_FILE = "PrabhavResume.pdf"         # Replace with your resume file name

# Excel File Settings
EXCEL_FILE = "demoemail.xlsx"           # Replace with your .xlsx file name
EMAIL_COLUMN = "Email "                  # Replace with the column name that contains emails
NAME_COLUMN = "Name"                    # Optional: column with HR person names (use None if not available)

# Email Template
EMAIL_TEMPLATE = """
Dear {hr_name},

I hope this email finds you well.

I am writing to express my interest in potential opportunities within your esteemed organization. 
I have attached my resume for your review and would welcome the chance to discuss how my skills 
and experience could contribute to your team.

Thank you for your time and consideration. I look forward to hearing from you.

Best regards,
{sender_name}
Phone: 9997153223
LinkedIn: https://www.linkedin.com/in/gupta-prabhav/
"""

# ================================
# FUNCTIONS
# ================================

def load_emails_from_excel(file_path, email_col, name_col=None):
    """
    Load emails from Excel file
    """
    try:
        # Read Excel file - handles both .xlsx and .xls
        df = pd.read_excel(file_path, engine='openpyxl')
        
        print(f"✅ Successfully loaded Excel file: {file_path}")
        print(f"📊 Found {len(df)} rows")
        print(f"📋 Columns available: {list(df.columns)}")
        
        # Check if email column exists
        if email_col not in df.columns:
            print(f"❌ Column '{email_col}' not found in Excel file")
            print(f"Available columns: {list(df.columns)}")
            return []
        
        # Extract emails and names
        emails = df[email_col].dropna().tolist()  # Remove empty cells
        names = []
        
        if name_col and name_col in df.columns:
            names = df[name_col].fillna("Sir/Madam").tolist()
        else:
            names = ["Sir/Madam"] * len(emails)
        
        # Combine emails and names
        hr_contacts = list(zip(emails, names))
        
        print(f"📧 Found {len(hr_contacts)} valid email addresses")
        return hr_contacts
        
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        return []

def send_email_with_resume(receiver_email, receiver_name, sender_email, sender_name, subject, resume_path):
    """
    Send email with resume attachment
    """
    try:
        # Create email message
        msg = EmailMessage()
        msg['From'] = f"{sender_name} <{sender_email}>"
        msg['To'] = receiver_email
        msg['Subject'] = subject
        
        # Personalize email body
        email_body = EMAIL_TEMPLATE.format(
            hr_name=receiver_name,
            sender_name=sender_name
        )
        msg.set_content(email_body)
        
        # Add resume attachment
        try:
            with open(resume_path, 'rb') as f:
                resume_data = f.read()
                resume_filename = resume_path.split('/')[-1]  # Get just filename
                msg.add_attachment(resume_data, 
                                 maintype='application', 
                                 subtype='pdf', 
                                 filename=resume_filename)
        except FileNotFoundError:
            print(f"⚠️  Resume file '{resume_path}' not found. Email will be sent without attachment.")
        
        # Send email using Gmail SMTP
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as server:
            server.login(sender_email, APP_PASSWORD)
            server.send_message(msg)
        
        print(f"✅ Email sent successfully to {receiver_name} ({receiver_email})")
        return True
        
    except Exception as e:
        print(f"❌ Failed to send email to {receiver_email}: {e}")
        return False

def main():
    """
    Main function to run the email campaign
    """
    print("🚀 Starting Email Campaign...")
    print("=" * 50)
    
    # Load emails from Excel
    hr_contacts = load_emails_from_excel(EXCEL_FILE, EMAIL_COLUMN, NAME_COLUMN)
    
    if not hr_contacts:
        print("❌ No valid email addresses found. Please check your Excel file.")
        return
    
    # Display contacts to be emailed
    print(f"\n📋 Contacts to email:")
    for i, (email, name) in enumerate(hr_contacts[:5], 1):  # Show first 5
        print(f"   {i}. {name} - {email}")
    if len(hr_contacts) > 5:
        print(f"   ... and {len(hr_contacts) - 5} more")
    
    # Confirm before sending
    proceed = input(f"\n🤔 Do you want to send emails to {len(hr_contacts)} recipients? (yes/no): ")
    if proceed.lower() not in ['yes', 'y']:
        print("❌ Email campaign cancelled.")
        return
    
    # Send emails
    print(f"\n📧 Sending emails...")
    successful = 0
    failed = 0
    
    for i, (email, name) in enumerate(hr_contacts, 1):
        print(f"\n📤 Sending email {i}/{len(hr_contacts)} to {name} ({email})")
        
        success = send_email_with_resume(
            receiver_email=email,
            receiver_name=name,
            sender_email=SENDER_EMAIL,
            sender_name=SENDER_NAME,
            subject=SUBJECT,
            resume_path=RESUME_FILE
        )
        
        if success:
            successful += 1
        else:
            failed += 1
        
        # Add delay between emails to avoid being flagged as spam
        if i < len(hr_contacts):  # Don't delay after last email
            print("⏳ Waiting 5 seconds before next email...")
            time.sleep(5)
    
    # Final summary
    print("\n" + "=" * 50)
    print("📊 EMAIL CAMPAIGN SUMMARY")
    print("=" * 50)
    print(f"✅ Successful: {successful}")
    print(f"❌ Failed: {failed}")
    print(f"📧 Total: {len(hr_contacts)}")
    print("🎉 Email campaign completed!")

if __name__ == "__main__":
    main()
