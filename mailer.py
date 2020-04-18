from csv import DictReader
import win32com.client as win32
from datetime import datetime as dt
from pathlib import Path

def seperator():
    return "===================================     " + dt.now().isoformat() + "\n"

def read_csv():
    with open(input("CSV File Path: "), 'r') as fh:
        return [x for x in DictReader(fh, skipinitialspace=True)]

def mail_send(email_set, subj, html_msg, inv_num, attachment=None):
    try:
        o = win32.Dispatch("Outlook.Application")
    except:
        print("ERROR: Unable to connect to Outlook.")
    mail = o.CreateItem(0)
    for email in email_set:
        mail.Recipients.Add(email)
    mail.Subject = subj
    mail.HTMLBody = html_msg
    try:
        mail.Attachments.Add(attachment)  # attachment must be a path to the file
    except:
        pass
    mail.Send()
    print("Sent email(s) to {emails} for invoice: {invoice}".format(emails=email_set, invoice=inv_num))

def main():
    while True:
        print("\nMain Menu\n" + seperator())
        print("Please select what you would like to do from the following options:")
        print("  1. Send Mail")
        print("  2. PDF Splitter (Unavailable)")
        print("  3. Exit")
        selector = input("Choice: ").lower()
        if selector in ("1", "one", "first", "mail", "send mail", "mailer"):

            mailing()
        elif selector in ("2", "two", "pdf", "split", "pdfsplit", "splitter"):
            print("\nPDF Splitter")
            pass
        elif selector in ("3", "three", "exit", "quit", "q", "e"):
            print("All contents of this log will be lost.")
            if input("Are you sure you would like to quit?(Y/N) ").lower() in ("y", "yes"):
                quit()
        else:
            print("Invalid selection, please try again.")

def mailing():
    while True:
        print("\nSend Mail")
        print(seperator() + "\n")
        print("Please select what type of invoice you are sending:")
        print("  1. UPS")
        print("  2. DHL")
        selector = input("Choice: ").lower()
        if selector in ("1", "ups"):
            template = "templates/UPS/"
            break
        elif selector in ("2", "dhl"):
            template = "templates/DHL/"
            break
    with open(template+'subject.txt', 'r') as fh:
        subject = fh.read()
    with open(template+'body.html', 'r') as fh:
        body_html = fh.read()
    inv_location = input("Invoice Path: ") + '\\{invoice}.pdf'
    for r in read_csv():
        mail_send(
            email_set=set(r['Email'].split(";")),
            subj=subject.format(company=r['Company'], invoice=r['Invoice Number']),
            html_msg=body_html.format(company=r['Company'], invoice=r['Invoice Number']),
            attachment=inv_location.format(invoice=r['Invoice Number']),
            inv_num=r['Invoice Number']
        )

if __name__ == "__main__":
    if int(dt.now().strftime("%H")) >= 12:
        print("Good Afternoon, Kim")
    else:
        print("Good Morning, Kim")
    main()
