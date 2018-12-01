import imaplib, smtplib, email, os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


to_addr = []

def sendMail(TO, SUBJECT, text):
    TEXT = 'Subject: {}\n\n{}'.format(SUBJECT, text)
    FROM = 'shapovo.kvit@yandex.ru'

    body = MIMEMultipart()
    body['From'] = FROM
    body['To'] = TO
    body['Subject'] = SUBJECT
    body.attach(MIMEText(text, 'plain'))

    server = smtplib.SMTP_SSL('smtp.yandex.ru')
    server.login('shapovo.kvit@yandex.ru', 'Shapovo')
    server.sendmail(FROM, [TO], body.as_string())


def parse_from_addr(s):
    result = list(email.utils.parseaddr(s))
    return result[1]


def decode_header(header):
    parts = []
    for part in email.header.decode_header(header):
        header_string, charset = part
        if charset:
            decoded_part = header_string.decode(charset)
        else:
            decoded_part = header_string
        parts.append(decoded_part)
    return "".join(parts)


def getMessage():
    mail = imaplib.IMAP4_SSL('imap.yandex.ru')  # данные сервера
    mail.login('shapovo.kvit@yandex.ru', 'Shapovo')  # логин на почту
    mail.list()  # список папок ящика
    mail.select('inbox')  # выбираем папку входящие
    result, data = mail.search(None, 'UNSEEN')  # UNSEEN
    # собираем письма в список и обрабатываем
    for num in data[0].split():
        result, data = mail.fetch(num, '(RFC822)')
        email_message = email.message_from_bytes(data[0][1])
        # чекаем на наличие вложений
        if email_message.get_content_maintype() != 'multipart':
            '''sendMail((parse_from_addr(msg['From'])), 'An error occurred while sending receipts',
                     'please send the file in one of the following formats: '
                     'pdf, png, jpeg, gif, doc(MS Word document).')'''
            print('content maintype != multipart')
            name = parse_from_addr(email_message['From'])
            if name not in to_addr:
                to_addr.append(name)
            mail.store(num, '+FLAGS', '\\Deleted')
            continue
        mail.store(num, '+FLAGS', '\\Seen')
        yield email_message


if __name__ == '__main__':
    counter = 1
    print('Поехали!')
    for msg in getMessage():
        for part in msg.walk():
            if part.get_content_type() == 'application/pdf' or 'image/png' \
                    or 'image/gif' or 'image/jpeg' or 'application/msword':

                if part.get('Content-Disposition') is None:
                    continue

                detach_dir = "F:\\KVIT\\"

                if not os.path.isdir(detach_dir):
                    os.makedirs(detach_dir)
                    continue

                filename = decode_header(part.get_filename())
                '''
                # if there is no filename, we create one with a counter to avoid duplicates
                if not filename:
                    filename = 'part-%03d%s' % (counter, 'bin')
                    counter += 1'''

                att_path = os.path.join(detach_dir, filename)

                # Check if its already there
                if not os.path.isfile(att_path):
                    # finally write the stuff
                    fp = open(att_path, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()

            else:
                print('non standart file format')
                name = parse_from_addr(msg['From'])
                if name not in to_addr:
                    to_addr.append(name)
    print('Еще немного....')
    for name in to_addr:
        print(name)
        sendMail(name, 'An error occurred while sending receipts',
                 'Please send the file in one of the following formats: pdf, png, jpeg, gif, doc (MS Office Word document).')

    print('Готово!')
