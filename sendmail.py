import smtplib
from pathlib import Path
from string import Template
from datetime import datetime as dt
from email.utils import formatdate
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


def login(settings):
    # print("[info]", "Connect 2 SMTP Server %s:%s" %
    #      (settings['smtp_server'], settings['port']))
    smtp = smtplib.SMTP(settings['smtp_server'], settings['port'])
    # print("[info]", "Start EHLO")
    smtp.ehlo()
    # print("[info]", "Start TLS")
    smtp.starttls()
    print("[info]", "login 2 %s" % (settings['mail']))
    smtp.login(settings['mail'], settings['pswd'])

    return smtp


def send(letter):
    print("[info]", "send letter from %s 2 %s" %
          (letter['From'], letter['To']))
    status = smtp.sendmail(letter['From'], letter['To'], letter.as_string())
    if status == {}:
        print("[info]", "Ok!")
    else:
        print("[info]", "Not OK!")


def close_smtp():
    smtp.quit()


def form_letter(letter_data):
    print("[info]", "Form Letter %s" % letter_data['subject'])
    msg = MIMEMultipart()
    msg["Subject"] = letter_data['subject']
    msg["From"] = letter_data['from_addr']
    msg["To"] = letter_data['to_addr']
    msg["Date"] = formatdate(localtime=True)

    body = MIMEText(letter_data['txt'], "plain")
    msg.attach(body)

    img = MIMEImage(Path("rat.png").read_bytes())
    msg.attach(img)

    # print(msg.as_string())

    return msg


def form_html_letter(letter_data, user):
    print("[info]", "Form Letter %s" % letter_data['subject'])
    msg = MIMEMultipart()
    msg["Subject"] = letter_data['subject']
    msg["From"] = letter_data['from_addr']
    msg["To"] = letter_data['to_addr']
    msg["Date"] = formatdate(localtime=True)
    #print(letter_data)

    # template = Template(Path("template-2.html").read_text(encoding='utf-8'))
    # html = template.substitute({"user": "Messy"})
    #print("[info]", "Form Letter HTML content")
    template = Template(
        Path("template-timer.html").read_text(encoding='utf-8'))
    #print(template)
    html = template.substitute({"user": user, 'date': dt.now().isoformat()})
    body = MIMEText(html, "html", 'utf-8')
    # body = MIMEText(letter_data['txt'], "plain")
    msg.attach(body)

    #print("[info]", "Add Image 2 Letter")
    img = Path("rat.png").read_bytes()
    img = MIMEApplication(img, Name='myrat.png')
    msg.attach(img)

    # print(msg.as_string())

    return msg


def load_pswd():
    fn = "../fuxksquid.pswd"
    pswd = Path(fn).read_text()

    return pswd


def load_recievers():
    blk_li = ['[', ']', ' ', '\t']
    fn = "li.txt"
    sep = '$'
    li = Path(fn).read_text(encoding='utf-8')
    for b in blk_li:
        li = li.replace(b, '')
    li = li.split('\n')
    li = [l.split(sep) for l in li if l != '']

    return li


settings = {
    'smtp_server': 'smtp.gmail.com',
    'port': 587,
    'mail': 'fuxksquid@gmail.com',
    'pswd': load_pswd()
}


user = 'Stanley'
letter_data = {
    'from_addr': settings['mail'],
    'to_addr': 'starcyh3@gmail.com',
    'subject': 'Hi %s!' % user,
    'txt': 'Hello World!\ni am doing cool things~~~~~~\nhow are you today',
    'image': 'rat.png'
}


recievers = load_recievers()
'''
recievers = [['starcyh@gmail.com', 'Stanley'],
             ['starcyh2@gmail.com', 'star222'],
             ['40847030s@gapps.ntnu.edu.tw', 'Messy']]
'''
print(recievers)

recievers = recievers[-6:]  # [['guojialin1206@gmail.com', '郭家麟']]
recievers = [['lingziha@gmail.com', 'Linz']]
print(recievers)

smtp = login(settings)


for rx in recievers:
    letter_data['to_addr'] = rx[0]
    user = rx[1]
    letter_data['subject'] = 'Hi %s! This is 4.7th TEST' % user
    letter = form_html_letter(letter_data, user)
    send(letter)
close_smtp()

'''
letter = form_letter(letter_data)
send(letter)
'''
