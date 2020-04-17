from csv import DictReader
import win32com.client as win32
o = win32.Dispatch('outlook.application')

def read_csv():
    with open(input("File: "), 'r') as fh:
        return [x for x in DictReader(fh, skipinitialspace=True)]

def mail_send(to, subj, html_msg, attachment=None):
    print(subj)
    print(html_msg)
    mail = o.CreateItem(0)
    mail.To = to
    mail.Subject = subj
    mail.Body = 'Filler'
    mail.HTMLBody = html_msg
    if(attachment):
        mail.Attachments.Add(attachment) # attachment must be a path to the file
    mail.Send()
        

def main():
    with open('templates/subject.txt', 'r') as fh:
        subject = fh.read()
    with open('templates/body.html', 'r') as fh:
        body_html = fh.read()
    for r in read_csv():
        mail_send(
            to=r['email'],
            subj=subject.format(invoice=r['invoice']),
            html_msg=body_html.format(name=r['name'], invoice=r['invoice']),
            attachment=None
        )

if __name__ == '__main__':
    main()