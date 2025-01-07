import os
import sys
from dotenv import load_dotenv
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import json
try:
    now = datetime.now()
    formatDateTime = now.strftime("%d-%m-%Y-%H:%M")
except Exception as e:
    with open ('logfile.log', 'a') as file:
        file.write(f"""Problem with date - {str(e)}\n""")

try: 
    local_file = f"test.xlsx"
    # Load env variables
    load_dotenv()
    site_url = os.getenv('site_url')
    file_url = os.getenv('file_url')
    office_username = os.getenv('office_username')
    password = os.getenv('password')

    # db
    db_password = os.getenv('db_password')
    db_user = os.getenv('db_user')
    db_server = os.getenv('db_server')
    db_driver = os.getenv('db_driver')
    db_db = os.getenv('db_db')

    # Email
    from_address = os.getenv('from_address')
    to_address_str = os.getenv('to_address')
    password = os.getenv('password')
except Exception as e:
    with open ('logfile.log', 'a') as file:
        file.write(f"""{formatDateTime} Problem with importing .env data - {str(e)}\n""")  
    sys.exit(0)
    
    
def read_db():
    try:
        import pyodbc
        conn = pyodbc.connect(f'DRIVER={db_driver};'
                        f'SERVER={db_server};'
                        f'DATABASE={db_db};'
                        f'UID={db_user};'
                        f'PWD={db_password}')
        cursor = conn.cursor()
        own_email_code = cursor.execute("SELECT * FROM contact_details_own WHERE code != LEFT(email, 4);").fetchall()
        agent_email_code = cursor.execute("SELECT * FROM contact_details_agent WHERE code != LEFT(salon_email, 4);").fetchall()
        own_phone = cursor.execute("SELECT *  FROM contact_details_own WHERE NOT (phone_number LIKE '[0-9][0-9][0-9]-[0-9][0-9][0-9]-[0-9][0-9][0-9]' OR phone_number LIKE '[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]');").fetchall()
        agent_phone = cursor.execute("SELECT *  FROM contact_details_agent WHERE NOT (phone_number LIKE '[0-9][0-9][0-9]-[0-9][0-9][0-9]-[0-9][0-9][0-9]' OR phone_number LIKE '[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]');").fetchall()

        return True,own_email_code,agent_email_code,own_phone,agent_phone
    except Exception as e:
        with open ('logfile.log', 'a') as file:
            file.write(f"""{formatDateTime} Problem with inserting data to db- {str(e)}\n""")
        return False,_,_,_,_

def send_mail(own_email_code,agent_email_code,own_phone,agent_phone):
    to_address = json.loads(to_address_str)
    msg = MIMEMultipart()
    msg['From'] = from_address
    msg["To"] = ", ".join(to_address)
    msg['Subject'] = f"Sprawdzenie integralności danych teleadresowych: {formatDateTime}."
    # Adding data to email message
    body = ''
    if own_email_code:
        body += "<p>Błędne dane salony własne Email - Kod salonu:\n</p>"
        results = [(row[0], row[-1]) for row in own_email_code]
        for row in results:
            body += f'<p>{row[0]} - {row[-1]}\n</p>'
        body += "\n"
            
    if agent_email_code:
        body += "<p>Błędne dane salony agencyjne Email - Kod salonu:\n<p>"
        results = [(row[0], row[-1]) for row in agent_email_code]
        for row in results:
            body += f'<p>{row[0]} - {row[-1]}\n</p>'
        body += "\n"
        
    if own_phone:
        body += "<p>Zły format telefonu salony włane:\n</p>"
        results = [(row[0], row[2]) for row in own_phone]
        for row in results:
            body += f'<p>{row[0]} - {row[-1]}\n</p>'
        body += "\n"
        
    if agent_phone:
        body += "<p>Zły format telefonu salony agencyjne:\n</p>"
        results = [(row[0], row[2]) for row in agent_phone]
        for row in results:
            body += f'<p>{row[0]} - {row[-1]}\n</p>'
        body += "\n"

    if not body:
        body = "<p>Brak błędów.</p>"
    msg.attach(MIMEText(body, 'html'))

    try:
        server = smtplib.SMTP('smtp-mail.outlook.com', 587)
        server.starttls()
        server.login(from_address, password)
        text = msg.as_string()
        server.sendmail(from_address, to_address, text)
        server.quit()               
    except Exception as e:
        with open ('logfile.log', 'a') as file:
            file.write(f"""{formatDateTime} Problem z wysłaniem na maile\n{str(e)}\n""")
            
    

status,own_email_code,agent_email_code,own_phone,agent_phone=read_db()
if status:
    send_mail(own_email_code,agent_email_code,own_phone,agent_phone)