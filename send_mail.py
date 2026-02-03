import os
import smtplib
import pandas as pd
from email.message import EmailMessage
from dotenv import load_dotenv


load_dotenv()
# --- НАСТРОЙКИ ---
# Путь к файлу со списком сотрудников (Табельный номер и почта)
employees_file = 'employees.xlsx'
# Папка, где лежат созданные файлы (имена файлов должны совпадать с "Табельным номером")
folder_path = 'calculations/'
# Название столбца в Excel, которое совпадает с названием файла
id_column = 'ID'
email_column = 'Email'

# Настройки почты (SMTP)
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = os.getenv("SMTP_PORT")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")

# Загружаем данные сотрудников
df_employees = pd.read_excel(employees_file)


def send_email(recipient_email, subject, body, attachment_path):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = SENDER_EMAIL
    msg['To'] = recipient_email
    msg.set_content(body)

    # Добавляем вложение
    with open(attachment_path, 'rb') as f:
        file_data = f.read()
        file_name = os.path.basename(attachment_path)
        msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)


# Проходим по списку сотрудников
def mail_for_employees():
    message_from_send_mail = []
    for index, row in df_employees.iterrows():
        employee_id = str(row[id_column])
        recipient = row[email_column]

        # Формируем путь к файлу
        file_name = f"{employee_id}.xlsx"
        full_path = os.path.join(folder_path, file_name)

        if os.path.exists(full_path):
            try:
                send_email(
                    recipient,
                    "Ваш отчет по ЗП",
                    f"Здравствуйте! Во вложении ваш файл: {file_name}",
                    full_path
                )
                message_from_send_mail.append(f"Отправлено: {file_name} для {recipient}")
            except Exception as e:
                message_from_send_mail.append(f"Ошибка при отправке для {recipient}: {e}")
        else:
            message_from_send_mail.append(f"Файл не найден: {full_path}")
    # print(message_from_send_mail)
    return message_from_send_mail
