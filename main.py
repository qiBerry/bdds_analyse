import smtplib as smtp
import csv

from tabulate import tabulate
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd


def makeDiffFile(bdds, plan, percent_diff):
    # write file part
    write_df = pd.DataFrame(
        columns=['Категория', 'Разница в %', 'Разница в абсолюте (миллионы рублей)', 'Причина разницы (заполнить)'])

    for index, row in bdds.iterrows():
        if (index > 6):
            bdds_value = row[6]
            plan_value = plan._get_value(index, 'Unnamed: 6')
            diff_value = (bdds_value - plan_value)
            if (not pd.isnull(row[1]) and not pd.isnull(row[6]) and abs(diff_value) > percent_diff and bdds_value > 0):
                diff_value_perc = (bdds_value - plan_value) / bdds_value
                print('diff', ' ', round(diff_value * 100, 2))
                print(row[6], ' ', plan._get_value(index, 'Unnamed: 6'))
                line = {'Категория': row[1], 'Разница в %': str(round(diff_value_perc * 100, 2)),
                        'Разница в абсолюте (миллионы рублей)': str(round(diff_value, 2)),
                        'Причина разницы (заполнить)': 'заполнить'}
                write_df = pd.concat([write_df, pd.DataFrame([line])], ignore_index=True)

    write_df.to_csv("output.csv", index=False)



def sendEmail(email, password, dest_email, text, html, filename_table):

    with open(filename_table) as input_file:
        reader = csv.reader(input_file)
        data = list(reader)

    text = text.format(table=tabulate(data, headers="firstrow", tablefmt="grid"))
    html = html.format(table=tabulate(data, headers="firstrow", tablefmt="html"))

    message = MIMEMultipart(
        "alternative", None, [MIMEText(text), MIMEText(html, 'html')])
    message['Subject'] = "Сформирован сводный отчет по БДДС"
    message['From'] = dest_email
    message['To'] = dest_email

    server = smtp.SMTP_SSL('smtp.yandex.com')
    server.set_debuglevel(1)
    server.ehlo(email)
    server.login(email, password)
    server.auth_plain()
    server.sendmail(email, dest_email, message.as_string())
    server.quit()

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    #Конфигурация параметров
    percent_diff = 0.05
    bdds = pd.read_excel('bdds.xlsx')
    plan = pd.read_excel('plan.xlsx')
    email = "misistestrpa@yandex.ru"
    password = "sessbpnrfneidsil"
    dest_email = "misistestrpa@yandex.ru"
    text = """
    Здравствуйте! Был сформирован отчет БДДС. В данном письме отображен сводный отчет с разницей плана и факта в 5%.
    В ответном письме, пожалуйста, заполните причины разницы. 
    
    {table}
    
    С уважением,
    ERP"""

    html = """
    <html><body><p>Здравствуйте! Был сформирован отчет БДДС. В данном письме отображен сводный отчет с разницей плана и факта в 5%.
    В ответном письме, пожалуйста, заполните причины разницы. </p>
    {table}
    <p>С уважением,</p>
    <p>ERP</p>
    </body></html>
    """
    filename_table = "output.csv"

    #формирование сводного отчета
    makeDiffFile(bdds, plan, percent_diff)

    #отправка отчета на почту
    sendEmail(email, password, dest_email, text, html, filename_table)
