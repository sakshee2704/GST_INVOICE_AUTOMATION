import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Function to send an email with an attachment
def send_email_with_attachment(filename):
    # Sender's email credentials
    email_user = 'patilsakshee41@gmail.com'
    email_password = 'bkfx ssoj rvdy aiid'  # App-specific password
    # Receiver's email address
    email_send = 'patilsakshee41@gmail.com'
    subject = 'GST Invoice for policy'

    # Create the email using MIMEMultipart for adding attachments
    msg = MIMEMultipart()
    msg['From'] = email_user
    msg['To'] = email_send
    msg['Subject'] = subject

    # Email body content
    body = '''Dear Sir/Madam,<br>
    Please find the attached GST invoice for the captioned Policy number.<br>
    To view the policy simply<br>
    1.Click on the attachment<br>
    2.It will prompt for password.The password is the risk commencement day of yor policy in DDMM format. 
      For example, if the risk commencement date is 07June, then the password would be 0706.
    '''
    msg.attach(MIMEText(body, 'html'))

    try:
        # Attach the file
        with open(filename, 'rb') as content_file:
            content = content_file.read()
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(content)
            encoders.encode_base64(part)
            # Add header for attachment with filename
            part.add_header('Content-Disposition', f'attachment; filename={filename.split("/")[-1]}')
            msg.attach(part)

        # Convert the message to a string
        text = msg.as_string()

        # Send the email
        server = smtplib.SMTP('smtp.gmail.com', 587) # Gmail's SMTP server and port
        server.starttls()  # Secure the connection
        server.login(email_user, email_password)
        server.sendmail(email_user, email_send, text) # Send the email
        server.quit()
        print("Email has been sent successfully!")

    except FileNotFoundError:
        print(f"Error: File '{filename}' not found.")
    except smtplib.SMTPAuthenticationError:
        print("Authentication failed. Check your email or app-specific password.")
    except Exception as e:
        print(f"Failed to send email: {e}")


