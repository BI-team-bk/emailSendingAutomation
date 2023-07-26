import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def send_email(sender_email, sender_password, receiver_email, subject, body, file_path):
    # Create a multipart message
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject

    # Add body to the email
    message.attach(MIMEText(body, "plain"))

    # Open the file in binary mode
    with open(file_path, "rb") as attachment:
        # Add file as application/octet-stream
        # Email clients can usually download this automatically as an attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode the file in ASCII characters to send by email
    encoders.encode_base64(part)

    # Add header as key/value pair to the attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {file_path}",
    )

    # Add attachment to the message and convert the message to a string
    message.attach(part)
    text = message.as_string()

    # Log in to the email server
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(sender_email, sender_password)

    # Send the email
    server.sendmail(sender_email, receiver_email, text)
    server.quit()

# Example usage
sender_email = "elhabimana@bk.rw"
sender_password = "Greece16@"
receiver_email = "elhabimana@bk.rw"
subject = "Automated File"
body = "Please find the attached CSV file."
file_path = "C:/Users/elhabimana/Downloads/bkapp-automated-reports/bkapp.csv"

send_email(sender_email, sender_password, receiver_email, subject, body, file_path)
