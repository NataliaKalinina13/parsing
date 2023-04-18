import os
import csv
import email
from email.header import decode_header

# -------------декодируем заголовок---------------------
def get_decoded_header(header):
    value, charset = decode_header(header)[0]
    if charset is None:
        return value
    else:
        return value.decode(charset)

# ------------достаем необходимые поля---------------------------------------------
def parse_eml_file(filename):
    with open(filename, "rb") as f:
        msg = email.message_from_binary_file(f)

        date = get_decoded_header(msg["Date"])
        subject = get_decoded_header(msg["Subject"])
        sender = get_decoded_header(msg["From"])
        recipient = get_decoded_header(msg["To"])

        body = ""
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                body = part.get_payload(decode=True).decode("utf-8")
                break

        return [date, sender, recipient, subject, body]

# -------------------------анализируем все файлы и сохраняем csv------------------------------------------------------
def parse_eml_directory(directory):
    with open("парсинг_почты.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Date", "From", "To", "Subject", "Body"])

        for filename in os.listdir(directory):
            if filename.endswith(".eml"):
                filepath = os.path.join(directory, filename)
                try:
                    fields = parse_eml_file(filepath)
                    writer.writerow(fields)
                except Exception as e:
                    print(f"Error parsing file {filepath}: {e}")


if __name__ == "__main__":
    parse_eml_directory("/Users/nataliakalinina/Desktop/messages")



