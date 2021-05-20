import telebot
import openpyxl

bot = telebot.TeleBot('1745813569:AAFqiQHg5cszsEa0D5CU9YnEwxpSz70OlrY')

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    bot.reply_to(message, f'Я бот. Приятно познакомиться, {message.from_user.first_name}')

file = openpyxl.open("goods4bot.xlsx", read_only=True)

sheet = file.active

#print(sheet[1][0].value)
print(sheet["B2"].value)

cells = sheet["A1":"C10"]
for type, coll, roll in cells:
    print(type.value, coll.value, roll.value)

#for row in range(1, sheet.max_row + 1):
#    type = sheet[row][0].value
#    collection = sheet[row][1].value
#    roll = sheet[row][2].value
#    print(row, type, collection, roll)
