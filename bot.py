import telebot
from telebot import types
import openpyxl
import sqlite3

bot = telebot.TeleBot('6031193706:AAEpiyTY6IJRDe-chpGJT14dFQQkd17DKnE')

# открываем файл с результатами и считываем с него данные в worksheet
wookbook = openpyxl.open("ex.xlsx", read_only=True)
worksheet = wookbook.active
questions, fraction, description, protoabilities, abilities, directions = [''] * 35, [''] * 7, [''] * 7, [''] * 7, [''] * 7, [''] * 7
for j in range(1, 8):
    fraction[j-1] += worksheet[8][j].value
    description[j-1] += worksheet[15][j].value
    protoabilities[j-1] += worksheet[14][j].value
c = 0
#(9, 14), (1, 8)
for i in range(9, 14):
    for j in range(1, 8):
        questions[c] += (worksheet[i][j].value)
        c += 1
welcome = worksheet[1][1].value
final = worksheet[16][1].value
for i in range(1, 8):
    proto = []
    for j in range(18, 25):
        if worksheet[j][i].value != None:
            proto.append(worksheet[j][i].value)
    directions[i-1] = proto
wookbook.close()

score = [0] * 7
position = 0


# создаем кнопки
btn1 = types.InlineKeyboardButton('1', callback_data='1')
btn2 = types.InlineKeyboardButton('2', callback_data='2')
btn3 = types.InlineKeyboardButton('3', callback_data='3')
btn4 = types.InlineKeyboardButton('4', callback_data='4')
btn5 = types.InlineKeyboardButton('5', callback_data='5')

#создаём переменные для работы с бд
id, name, surname, password, email, phone_number = 0, '', '', '', '', ''

@bot.message_handler(commands=['start'])
def start(message):
    conn = sqlite3.connect('DataBase.sql')
    cur = conn.cursor()
    #cur.execute('DROP TABLE IF EXISTS users')
    #cur.execute('DROP TABLE IF EXISTS results')
    cur.execute('CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY AUTOINCREMENT, name varchar(20), surname varchar(20), pass varchar(20), email varchar(20), phone_number varchar(15))')
    cur.execute('CREATE TABLE IF NOT EXISTS results(id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER REFERENCES users(id), result varchar(20))')
    conn.commit()
    cur.close()
    conn.close()
    bot.send_message(message.chat.id,'Привет, давай зарегистрируемся. Введи имя')
    bot.register_next_step_handler(message, user_name)


def user_name(message):
    global name
    name = message.text.strip()
    bot.send_message(message.chat.id,'Введи фамилию')
    bot.register_next_step_handler(message, user_surname)

def user_surname(message):
    global surname
    surname = message.text.strip()
    bot.send_message(message.chat.id,'Введи пароль')
    bot.register_next_step_handler(message, user_password)

def user_password(message):
    global password
    password = message.text.strip()
    bot.send_message(message.chat.id,'Введи email')
    bot.register_next_step_handler(message, user_email)

def user_email(message):
    global email
    email = message.text.strip()
    bot.send_message(message.chat.id,'Введи телефон')
    bot.register_next_step_handler(message, user_phone_number)

def user_phone_number(message):
    markup1 = types.InlineKeyboardMarkup()
    markup1.add(types.InlineKeyboardButton('Начать', callback_data='start'))    

    global phone_number
    phone_number = message.text.strip()

    conn = sqlite3.connect('DataBase.sql')
    cur = conn.cursor()

    cur.execute(f"SELECT id FROM users WHERE name = '%s' and surname = '%s' and pass = '%s' and email = '%s' and phone_number = '%s'"% (name, surname, password, email, phone_number))
    id = cur.fetchall()
    if id != []:
        bot.send_message(message.chat.id, 'Вы уже зарегистрированы')
    else: 
        cur.execute(f"INSERT INTO users(name, surname, pass, email, phone_number) VALUES('%s', '%s', '%s', '%s', '%s')" % (name, surname, password, email, phone_number))
        bot.send_message(message.chat.id, 'Пользователь зарегистрирован')
        conn.commit()
    cur.close()
    conn.close()
    bot.send_message(message.chat.id,f'{welcome}', reply_markup=markup1 )
    

@bot.callback_query_handler(func=lambda callback: True)
def callback_message(callback):
    if callback.data == 'start':
        question(callback.message)
    elif callback.data == 'directions':
        for i in range(len(directions[index])):
            bot.send_message(callback.message.chat.id, {directions[index][i]})
    else:
        if callback.data == '1':
            counting(1)
        elif callback.data == '2':
            counting(2)
        elif callback.data == '3':
            counting(3)
        elif callback.data == '4':
            counting(4)
        elif callback.data == '5':
            counting(5)
        question(callback.message)


#функция создания вопросов и подсчёта результатов
def question(message):
    markup = types.InlineKeyboardMarkup()
    markup.row(btn1, btn2, btn3, btn4, btn5)

    global position
    global score
    global index

    if len(questions) > position:
        bot.send_message(message.chat.id, f'{questions[position]}', reply_markup=markup)
        position += 1
    else:
        index = score.index(max(score))
        bot.send_message(message.chat.id, f'{final} {fraction[index]}')
        bot.send_message(message.chat.id, f'{description[index]}')
        for i in range(7):
            abilities[i] = str(protoabilities[i]) + ' - ' + str(score[i])
        bot.send_message(message.chat.id,'Ваши способности:')
        bot.send_message(message.chat.id,{"\n".join(abilities)})
        result = fraction[index]
        conn = sqlite3.connect('DataBase.sql')
        cur = conn.cursor()
        cur.execute(
            f"SELECT id FROM users WHERE name = '%s' and surname = '%s' and pass = '%s' and email = '%s' and phone_number = '%s'"% (name, surname, password, email, phone_number))
        id = cur.fetchall()
        
        cur.execute(
            f"INSERT INTO results(user_id, result) VALUES('%s', '%s')" % (id, result))
        cur.execute(
            f"SELECT result FROM results WHERE user_id = '%s'"% (id))
        results = cur.fetchall()
        conn.commit()
        cur.close()
        conn.close()

        print_results = ''
        for i in range(len(results)):
            if i != 0:
                print_results += ', ' + ''.join(results[i]) 
            else:
                print_results += ''.join(results[i]) 
        bot.send_message(message.chat.id, f'Ваши предыдущие результаты: {print_results}')
        markup1 = types.InlineKeyboardMarkup()
        markup1.add(types.InlineKeyboardButton('Вывести', callback_data='directions'))
        bot.send_message(message.chat.id, f'Вывести рекомендуемые направления?', reply_markup=markup1)
        markup2 = types.InlineKeyboardMarkup()
        markup2.add(types.InlineKeyboardButton('Да', callback_data='start'))
        bot.send_message(message.chat.id, 'Пройти тест ещё раз?', reply_markup=markup2)

        position = 0
        score = [0] * 7


#функция подсчёта ответов
def counting(number):
    global position
    sc = (position - 1) % 7
    for i in range(7):
        if sc == i:
            score[sc] += number
            break

bot.polling(none_stop=True)