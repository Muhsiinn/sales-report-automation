import pandas as pd
from email.message import EmailMessage
from email.utils import formataddr
from dotenv import load_dotenv
import os
import smtplib
from pathlib import Path
dotenv_path = Path(__file__).resolve().parent.parent/ '.env'
print("EMAIL_ADDRESS loaded as:", os.getenv("EMAIL_ADDRESS"))
print("EMAIL_PASSWORD loaded as:", os.getenv("EMAIL_PASSWORD"))

load_dotenv(dotenv_path)
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Path to report
BASE_DIR = Path(__file__).resolve().parent.parent
report_path = BASE_DIR / "reports" / "ecommerce_Report.xlsx"

# Load sheets
sheets = pd.read_excel(report_path, sheet_name=None)

# Convert to HTML
def df_to_html_table(df, title):
    return f"<h2>{title}</h2>{df.to_html(index=False)}<br><br>"

html_parts = [df_to_html_table(df.head(10), name) for name, df in sheets.items()]

html_content = f"""
<html><body>
<h1>📊Sales Report</h1>
{''.join(html_parts)}
<p>– Automated by Sales Report Automation</p>
</body></html>
"""

def send_email(subject, recipient, content_html):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = formataddr(("Sales Report Automation", EMAIL_ADDRESS))
    msg['To'] = recipient
    msg.set_content("HTML report attached. Please view in an HTML-compatible email client.")
    msg.add_alternative(content_html, subtype='html')

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.send_message(msg)
        print("Email sent successfully!")

if __name__ == "__main__":
    send_email("📈 Weekly Sales Report", "noobmuhsin@gmail.com", html_content)
