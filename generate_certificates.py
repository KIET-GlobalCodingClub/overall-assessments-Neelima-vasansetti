import pandas as pd
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML  # ✅ instead of pdfkit
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
import os

# Setup Jinja2 environment (current folder)
env = Environment(loader=FileSystemLoader('.'))
template = env.get_template("certificate_template.html")

# Certificates folder
OUTPUT_DIR = "certificates"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Function to create certificate PDF
def create_certificate(name, roll_number, department):
    html_content = template.render(Name=name, Roll_number=roll_number, Department=department)
    file_name = os.path.join(OUTPUT_DIR, f"{name.replace(' ', '_')}_certificate.pdf")

    # ✅ Generate PDF using WeasyPrint
    HTML(string=html_content, base_url=".").write_pdf(file_name)

    return file_name

# Function to send email
def send_email(to_email, subject, body, attachment):
    from_email = "neelimavasanasetti@gmail.com"  # 🔴 change this
    from_password = "saca ggqb sbpn ufys"  # 🔴 use Gmail App Password

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Attach PDF
    with open(attachment, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(attachment)}")
        msg.attach(part)

    # Send Email
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(from_email, from_password)
        server.send_message(msg)

# Main
def main():
    df = pd.read_excel('students.xlsx')  # Must have Name, Roll_number, Department, Email columns

    for _, row in df.iterrows():
        name = row['Name']
        roll_number = row['Roll_number']
        department = row['Department']
        email = row['Email']

        # Generate certificate
        cert_file = create_certificate(name, roll_number, department)

        # Send email
        subject = "Your Certificate of Participation"
        body = f"Dear {name},\n\nPlease find attached your certificate of participation.\n\nBest regards,\nML Hackathon Team"
        send_email(email, subject, body, cert_file)

if __name__ == "__main__":
    main()
