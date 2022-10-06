import csv
import gspread
import configparser
import telegram
import pprint
from oauth2client.service_account import ServiceAccountCredentials
from datetime import date

def checker(secretFile, scope, spreadsheetName, worksheetName, headers):
  # Get current date
  today_formatted = date.today().strftime('%d/%m/%Y')

  # Credential
  creds = ServiceAccountCredentials.from_json_keyfile_name(secretFile, scope)
  client = gspread.authorize(creds)

  # Load spreadsheet
  sheet = client.open(spreadsheetName)
  worksheet = sheet.worksheet(worksheetName)
  data = worksheet.get_all_records()

  expired_users=[]

  # Check expired user
  for i in data:
    if today_formatted == i["Due Date Access"]:
      expired_users.append(i)

  # Delete unused key
  expired_users = [{ key: item[key] for key in headers } for item in expired_users]
  return expired_users



def csv_generator(headers, data, exportedFilePath):
  with open(exportedFilePath, 'w', encoding='UTF8', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=headers)
    writer.writeheader()
    writer.writerows(data)


def send(msg, chatId, token, document):
  bot = telegram.Bot(token=token)
  bot.send_document(chatId, document=document, caption=msg)


def main():
  config = configparser.ConfigParser()
  config.read('setting.conf')

  telegram_token = config.get('telegram', 'bot_token')
  chat_id = config.get('telegram', 'chat_id')
  scope = config.get('credential', 'scope').split(',')
  secret = config.get('credential','secret_file')
  spreadsheet_name = config.get('spreadsheet', 'spreadsheet_name')
  worksheet_name = config.get('spreadsheet', 'worksheet_name')
  headers = config.get('spreadsheet', 'keep_column').split(',')
  data_list = checker(secret, scope, spreadsheet_name, worksheet_name, headers)
  file_name = config.get('storage', 'filename')
  msg = config.get('telegram', 'caption')

  if len(data_list):
    csv_generator(headers, data_list, file_name)
    file_document = open(file_name, 'rb')
    send(msg, chat_id, telegram_token, file_document)

if __name__ == "__main__":
  main()
