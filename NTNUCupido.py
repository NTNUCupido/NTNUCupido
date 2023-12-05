from time import sleep
from datetime import datetime as dt

import hashlib
import smtplib
from pathlib import Path
from string import Template
from datetime import datetime as dt
from email.utils import formatdate
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


from quickstart import GooDocX


#################### BASE FUNCTIONS ###################


############### send mail ###############
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


def send(smtp, letter):
    print("[info]", "send letter from %s 2 %s" %
          (letter['From'], letter['To']))
    status = smtp.sendmail(letter['From'], letter['To'], letter.as_string())
    if status == {}:
        print("[info]", "Ok!")
    else:
        print("[info]", "Not OK!")

    return status


def close_smtp(smtp):
    smtp.quit()


def load_pswd():
    fn = "../pswd/fuxksquid.pswd"
    pswd = Path(fn).read_text()

    return pswd


def form_html_letter(letter_data):

    print("[info]", "Form Letter %s" % letter_data['subject'])
    letter = MIMEMultipart()
    letter["Subject"] = letter_data['subject']
    letter["From"] = letter_data['from_addr']
    letter["To"] = letter_data['to_addr']
    letter["Date"] = formatdate(localtime=True)

    template = Template(
        Path(letter_data['template']).read_text(encoding='utf-8'))
    html = template.substitute(letter_data['substitute'])
    body = MIMEText(html, "html", 'utf-8')
    letter.attach(body)

    return letter


settings = {
    'smtp_server': 'smtp.gmail.com',
    'port': 587,
    'mail': 'fuxksquid@gmail.com',
    'pswd': load_pswd()
}

############### hash ###############


def sha256(sth):
    sth = sth.encode()
    return hashlib.sha256(sth).hexdigest()


################### Gears ###################


def get_sheet_values(sheet_id, range_name):
    res = sheets.spreadsheets().values().get(
        spreadsheetId=sheet_id, range=range_name).execute()
    rows = res.get('values')  # ['values']

    return rows


def add_sheet_values(sheet_id, range_name, values):
    body = {"values": values}
    res = sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id, range=range_name,
        valueInputOption="USER_ENTERED", body=body).execute()

    # ValueInputOption Values:
    #   INPUT_VALUE_OPTION_UNSPECIFIED  : Default input value. This value must not be used.
    #   RAW                             : The values the user has entered will not be parsed and will be stored as-is.
    #   USER_ENTERED                    : same as User Input, could be a function or others.

    return res


def check_position(sheet_id, table, row_id_offset, row_id, mail):
    data = get_sheet_values(sheet_id, table)[1:][row_id-row_id_offset]
    if mail in data:
        return True
    else:
        print('[ERROR]:', 'check_position error')
        return False


def search_position(sheet_id, table, col_name, target):
    target = [target]
    col_name = col_name.upper()
    range_name = "'%s'!%s1:%s999" % (table, col_name, col_name)
    res = sheets.spreadsheets().values().get(
        spreadsheetId=sheet_id, range=range_name).execute()
    rows = res.get('values')  # ['values']
    print(rows)
    try:
        ### search first match: could not found new data already exist in sheets. ###
        # idx = rows.index(target) + 1
        # pos = col_name + str(idx)

        ### search last match: could fix the above problem temporary. ###
        idx = len(rows) - 1 - rows[::-1].index(target)
        idx = str(idx)

    except:
        idx = ""
       # print('Not Found!')

    return idx


def smtp_login(smtp_settings):
    print('[info]', 'Connecting to SMTP Server.....')
    smtp_data = login(smtp_settings), smtp_settings['mail']

    return smtp_data


def send_verify_mail(smtp_data, mail, verify_code, verify_form_id):
    smtp, my_mail = smtp_data
    form_link = "https://docs.google.com/forms/d/e/%s/viewform?usp=pp_url&entry.759435213=%s&entry.1600004906=%s"
    form_link %= (verify_form_id, mail, verify_code)
    letter_data = {
        'from_addr': my_mail,
        'to_addr': mail,
        'subject': 'Hi 2 don\'t know who U R, But Here is your Verify Code',  # % user,
        'template': "letter_template/template-Verify-Letter.html",
        'substitute': {"vcode": verify_code, "vlink": form_link, 'date': dt.now().isoformat().split('.')[0]}
    }
    letter = form_html_letter(letter_data)
    status = send(smtp, letter)

    return status


def send_verify_res_mail(smtp_data, mail, status):
    smtp, my_mail = smtp_data
    letter_data = {
        'from_addr': my_mail,
        'to_addr': mail,
        'subject': 'Congratulations! Your Mail has been SUCCESSFULLY verified',  # % user,
        'template': "letter_template/template-Verify-res-Letter.html",
        'substitute': {"urmail": mail, "date": dt.now().isoformat().split('.')[0]}
    }
    if status == 'SUCCESSFUL':
        letter_data['subject'] = 'Congratulations! Your Mail has been SUCCESSFULLY verified'
        letter_data['template'] = "letter_template/template-Verify-successful-Letter.html"
        letter_data['substitute'] = {"urmail": mail,
                                     "date": dt.now().isoformat().split('.')[0]}
    else:
        form_link = "https://docs.google.com/forms/d/e/1FAIpQLSejrkf5XffIUCl1PRWtB5-jO95gBQSofFoB2fdxyE12uyBYnQ/viewform?usp=pp_url&entry.759435213=%s&entry.1600004906=%s"
        form_link %= (mail, gen_verify_code(mail))
        letter_data['subject'] = 'Verification Failed!'
        letter_data['template'] = "letter_template/template-Verify-failed-Letter.html"
        letter_data['substitute'] = {"urmail": mail, "vlink": form_link,
                                     "date": dt.now().isoformat().split('.')[0]}

    letter = form_html_letter(letter_data)
    status = send(smtp, letter)

    return status


def send_psychometric_test_mail(smtp_data, mail, status):
    smtp, my_mail = smtp_data
    form_link = "https://docs.google.com/forms/d/e/1FAIpQLSejrkf5XffIUCl1PRWtB5-jO95gBQSofFoB2fdxyE12uyBYnQ/viewform?usp=pp_url&entry.759435213=%s&entry.1600004906=%s"
    form_link %= (mail, gen_verify_code(mail))
    letter_data = {
        'from_addr': my_mail,
        'to_addr': mail,
        'subject': 'You are successfully verified your mail, and this psychometric_test_mail is for you',  # % user,
        'template': "letter_template/template-Verify-Letter.html",
        'substitute': {"vcode": verify_code, "vlink": form_link, 'date': dt.now().isoformat()}
    }
    letter = form_html_letter(letter_data)
    status = send(smtp, letter)

    return status


def checked_mail(sheet_id, table, checked_col_name, row_id_offset, row_id, mail):
    # pos = search_position(sheet_id, table, mail_col_name, mail)
    status = check_position(
        sheet_id, table, row_id_offset, row_id, mail)
    if status:
        range_name = "'%s'!" % table + checked_col_name + str(row_id)
        print(range_name)
        res = add_sheet_values(sheet_id, range_name, [['X']])
    else:
        print('[search_position]:Target "%s" Not Found!' % mail)

    return res


def checked_verify_code(sheet_id, table, checked_col, verify_col, row_id_offset, row_id, mail, chk_status):
    # pos = search_position(sheet_id, table, mail_col_name, mail)
    status = check_position(sheet_id, table, row_id_offset, row_id, mail)

    if status:
        crange_name = "'%s'!" % table + checked_col + str(row_id)
        print(crange_name)
        res = add_sheet_values(sheet_id, crange_name, [['X']])

        vrange_name = "'%s'!" % table + verify_col + str(row_id)
        print(vrange_name)
        res = add_sheet_values(sheet_id, vrange_name, [[chk_status]])

    else:
        print('[search_position]: Target "%s" Not Found!' % mail)

    return res


def check_new_mails_in(sheet_id, table, row_id_offset):
    res = get_sheet_values(sheet_id, table)#[1:]
    col_len = len(res[0])
    res = res[1:]
    # print(res)
    new_mails = []
    for i in range(len(res)):
        if len(res[i]) != col_len:  # 當欄位長度比資料長度長得時候，代表還每打果勾勾代表新資料
            new_mails.append([i + row_id_offset, res[i][1]])
    # new_mails = [r[0] for r in res if len(r) == 1]

    return new_mails


def check_new_verify_code_in(sheet_id, table, row_id_offset):
    res = get_sheet_values(sheet_id, table)[1:]
    new_vcodes = []
    for i in range(len(res)):
        if len(res[i]) == 3:
            new_vcodes.append([i + row_id_offset, res[i][1], res[i][2]])
    #new_vcodes = [r for r in res if len(r) == 2]

    return new_vcodes


def gen_verify_code(sth):
    hash_times = 5
    sth = str(sth)
    salt = 'I Love Cupido!!'
    nothing = "dEstr0yR@1nB0wTAb1es"
    sth = sha256(sth + salt + nothing)
    for _ in range(hash_times-1):
        sth = sha256(sth)

    return sth


################### Processes ###################


def verify_school_mail_process(register_sheet_setting, verify_sheet_setting, smtp_settings):

    ### init ###
    register_sheet_id = register_sheet_setting['sheet_id']
    register_table = register_sheet_setting['table']
    register_mail_col = register_sheet_setting['mail_col']
    register_checked_col = register_sheet_setting['checked_col']
    register_row_id_offset = register_sheet_setting['row_id_offset']

    verify_sheet_id = verify_sheet_setting['sheet_id']
    verify_table = verify_sheet_setting['table']
    verify_mail_col = verify_sheet_setting['mail_col']
    verify_checked_col = verify_sheet_setting['checked_col']
    verify_verify_col = verify_sheet_setting['verify_col']
    verify_row_id_offset = verify_sheet_setting['row_id_offset']
    verify_form_id = verify_sheet_setting['form_id']

    smtp_data = smtp_login(smtp_settings)
    while 1:
        ### Send verify code to registed mail ###
        try:
            datas = check_new_mails_in(
                register_sheet_id, register_table, register_row_id_offset)
        except:
            sleep(5)
            continue

        if datas != []:
            for data in datas:
                row_id, mail = data
                verify_code = gen_verify_code(mail)
                status = send_verify_mail(smtp_data, mail, verify_code, verify_form_id)
                if status == {}:
                    checked_mail(register_sheet_id, register_table,
                                 register_checked_col, register_row_id_offset, row_id, mail)

        ### Check if verify code is right ###
        try:
            datas = check_new_verify_code_in(
                verify_sheet_id, verify_table, verify_row_id_offset)
        except:
            sleep(5)
            continue

        if datas != []:
            for data in datas:
                row_id, mail, vcode = data
                if gen_verify_code(mail) == vcode:
                    status = 'SUCCESSFUL'

                else:
                    status = 'FAILED'
                checked_verify_code(verify_sheet_id, verify_table, verify_checked_col,
                                    verify_verify_col, verify_row_id_offset, row_id, mail, status)
                send_verify_res_mail(smtp_data, mail, status)
                #send_psychometric_test_mail(smtp_data, mail, status)

        sleep(3)

'''
def make_verified_mail_list_proccess(verify_sheet_setting):
    cmd = """={UNIQUE(QUERY(IMPORTRANGE("https://docs.google.com/spreadsheets/d/%s", "表單回應 1!B:E"), "SELECT Col1 WHERE Col4='SUCCESSFUL'"))}"""
    cmd %= verify_sheet_setting['sheet_id']

    spreadsheet = {
        'properties': {
            'title': "Verified_mail_list"
        }
    }
    spreadsheet = sheets.spreadsheets().create(
        body=spreadsheet, fields='spreadsheetId').execute()
    # print(f"Spreadsheet ID: {(spreadsheet.get('spreadsheetId'))}")
    verified_mails_sheet_setting = {"sheet_id": spreadsheet.get('spreadsheetId'), 'table': "工作表1", "mail_col": "A"}

    add_sheet_values(verified_mails_sheet_setting['sheet_id'], "%s!A1"%verified_mails_sheet_setting['table'], [[cmd]])

    return verified_mails_sheet_setting
'''


def test_and_match_proccess(register_sheet_setting, verify_sheet_setting, smtp_settings):
    pass


################### init phase ###################
drive, sheets, docs = GooDocX()

smtp_settings = {
    'smtp_server': 'smtp.gmail.com',
    'port': 587,
    'mail': 'fuxksquid@gmail.com',
    'pswd': load_pswd()
}


register_sheet_setting = {"sheet_id": "1hTfsAKIabmOhyoFUuvUlPdGgTIPzMC2l-HDPPOC4tSA",#####
                          "table": "表單回應 1",
                          "mail_col": "B",
                          "checked_col": "I",
                          "row_id_offset": 2}

verify_sheet_setting = {"sheet_id": "1iV7N8BWQ34WjQh-sC39SMEWuvBmoIYG3XAwM-ULZ0lk",#####
                        "table": "表單回應 1",
                        "mail_col": "B",
                        "checked_col": "D",
                        "verify_col": "E",
                        "row_id_offset": 2,
                        "form_id": "1FAIpQLSdZmzZciV6rFEXry8O2xWv-UaxfKkO5BpXv8pp8-9FMFQLRAQ"}#####

verified_mails_sheet_setting = {"sheet_id": "1AQm8p1ydDiWzHc02vRUD4imyNUdXyiSHqgYw3pOVrcU",
                                "table": "工作表1",
                                "mail_col": "A",
                                "checked_col": "B",
                                "row_id_offset": 1}


################### Main Function ###################

debug = 0

if debug:
    #smtp_data = smtp_login(smtp_settings)
    test = make_verified_mail_list_proccess

    res = test(verify_sheet_setting)
    print(res)
    #for r in res:
    #    print(r)

if not debug:
    verify_school_mail_process(register_sheet_setting,
                               verify_sheet_setting, smtp_settings)

################### 2 D0 ###################

# [all]: ADD terminal info
# [checked_mail]: need the new way for mail that has been registered more than once.
