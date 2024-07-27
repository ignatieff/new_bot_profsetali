import telebot
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from telebot import types

# Настройка Google Sheets API
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("C:/pynew/json/credentials.json", scope)
client = gspread.authorize(creds)

# Откройте таблицу по URL
spreadsheet_url = "https://docs.google.com/spreadsheets/d/15XAjqmV9a977G1NLdgl-F-CIxewqCf35GuSMhmiBCRY"
spreadsheet = client.open_by_url(spreadsheet_url)

# Имя фиксированного листа
FIXED_SHEET_NAME = "TGBOT"  # Замените на имя вашего листа

# Инициализация Telegram Bot
TOKEN = '6734116838:AAGrgQW5zRRf5WDkI5qCpKyynSRYpZ8Nz3Y'
bot = telebot.TeleBot(TOKEN)

# Глобальные переменные для хранения состояния админ-панели и поиска
admin_ids = [1267722744, 382647732]  # Замените на ID администраторов
current_step = {}
column_names = {}

# Функция для экранирования MarkdownV2
def escape_markdown_v2(text):
    return text.replace('_', '\\_').replace('*', '\\*').replace('[', '\\[').replace(']', '\\]').replace('(', '\\(').replace(')', '\\)').replace('~', '\\~').replace('`', '\\`').replace('>', '\\>').replace('#', '\\#').replace('+', '\\+').replace('-', '\\-').replace('=', '\\=').replace('|', '\\|').replace('{', '\\{').replace('}', '\\}').replace('.', '\\.')

# Функция для создания начальной клавиатуры
def create_start_keyboard():
    keyboard = types.ReplyKeyboardMarkup(row_width=2, resize_keyboard=True)
    keyboard.add(types.KeyboardButton("Поиск"), types.KeyboardButton("Добавление"))
    return keyboard

# Функция для создания клавиатуры админ-панели
def create_admin_keyboard():
    keyboard = types.ReplyKeyboardMarkup(row_width=2, resize_keyboard=True)
    keyboard.add(types.KeyboardButton("Записать значение"), types.KeyboardButton("Назад"))
    return keyboard

# Функция для создания клавиатуры с кнопками "Назад" и "Отмена"
def create_navigation_keyboard():
    keyboard = types.ReplyKeyboardMarkup(row_width=2, resize_keyboard=True)
    keyboard.add(types.KeyboardButton("Назад"), types.KeyboardButton("Отмена"))
    return keyboard

# Функция для записи данных в Google Sheets
def write_to_sheet(sheet_name, column_name, value, row_idx):
    sheet = spreadsheet.worksheet(sheet_name)
    headers = sheet.row_values(1)
    try:
        col_idx = headers.index(column_name) + 1  # Индекс столбца (1-based)
    except ValueError:
        return "Столбец не найден."
    
    # Запись значения в указанную строку и столбец
    sheet.update_cell(row_idx, col_idx, value)
    return "Данные успешно записаны."

# Функция для получения заголовков столбцов
def get_column_headers(sheet_name):
    sheet = spreadsheet.worksheet(sheet_name)
    headers = sheet.row_values(1)
    return headers

# Обработчик команды /start
@bot.message_handler(commands=['start'])
def send_start(message):
    bot.send_message(message.chat.id, "Добро пожаловать! Выберите действие:", reply_markup=create_start_keyboard())

# Функция поиска
def search_sheet(query):
    sheet = spreadsheet.worksheet(FIXED_SHEET_NAME)
    matches = []
    # Преобразование запроса в нижний регистр
    lower_query = query.lower()
    
    # Поиск совпадений по всем ячейкам
    cells = sheet.get_all_values()
    for row_idx, row in enumerate(cells):
        for col_idx, cell in enumerate(row):
            if cell.lower() == lower_query:
                matches.append((row_idx + 1, col_idx + 1))  # Сохранение индексов строки и столбца
    
    if matches:
        results = []
        headers = sheet.row_values(1)
        for match in matches:
            row = sheet.row_values(match[0])
            result = []
            for i, cell_value in enumerate(row):
                header = escape_markdown_v2(headers[i])
                cell_value = escape_markdown_v2(cell_value)
                result.append(f"{header}: {cell_value}")
                
            results.append('\n'.join(result))
        
        return results
    else:
        return ["Результаты поиска не найдены."]

# Обработчик текстовых сообщений
@bot.message_handler(func=lambda message: True)
def handle_messages(message):
    chat_id = message.chat.id
    text = message.text

    # Логирование для отладки
    print(f"Received message: '{text}' from chat_id: {chat_id}, current_step: {current_step}")

    if text == "Поиск":
        bot.send_message(chat_id, "Введите запрос для поиска:", reply_markup=create_navigation_keyboard())
        current_step[chat_id] = {"action": "search_query"}

    elif text == "Добавление":
        if message.from_user.id in admin_ids:
            headers = get_column_headers(FIXED_SHEET_NAME)
            if len(headers) > 2:
                headers = headers[:2]
            column_names[chat_id] = headers
            current_step[chat_id] = {'action': 'input_values', 'row_idx': len(spreadsheet.worksheet(FIXED_SHEET_NAME).col_values(1)) + 1, 'column_idx': 0}
            bot.send_message(chat_id, f"Выбран лист '{FIXED_SHEET_NAME}'. Введите значения для двух первых столбцов по порядку:\n{', '.join(escape_markdown_v2(header) for header in headers)}", reply_markup=create_navigation_keyboard())
        else:
            bot.send_message(chat_id, "У вас нет доступа к админ-панели.")

    elif text == "Записать значение":
        if message.from_user.id in admin_ids:
            headers = get_column_headers(FIXED_SHEET_NAME)
            if len(headers) > 2:
                headers = headers[:2]
            column_names[chat_id] = headers
            current_step[chat_id] = {'action': 'input_values', 'row_idx': len(spreadsheet.worksheet(FIXED_SHEET_NAME).col_values(1)) + 1, 'column_idx': 0}
            bot.send_message(chat_id, f"Выбран лист '{FIXED_SHEET_NAME}'. Введите значения для двух первых столбцов по порядку:\n{', '.join(escape_markdown_v2(header) for header in headers)}", reply_markup=create_navigation_keyboard())
        else:
            bot.send_message(chat_id, "У вас нет доступа к админ-панели.")

    elif text == "Назад":
        if "action" in current_step.get(chat_id, {}):
            action = current_step[chat_id]['action']
            if action == "search_query":
                bot.send_message(chat_id, "Выберите действие:", reply_markup=create_start_keyboard())
                del current_step[chat_id]
            elif action in ["input_values"]:
                bot.send_message(chat_id, "Выберите действие:", reply_markup=create_admin_keyboard())
                del current_step[chat_id]
        else:
            bot.send_message(chat_id, "Выберите действие:", reply_markup=create_start_keyboard())

    elif text == "Отмена":
        bot.send_message(chat_id, "Действие отменено. Выберите действие:", reply_markup=create_start_keyboard())
        if chat_id in current_step:
            del current_step[chat_id]

    elif "action" in current_step.get(chat_id, {}):
        action = current_step[chat_id]['action']
        if action == "search_query":
            query = text
            try:
                results = search_sheet(query)
                for result in results:
                    if len(result) > 4096:
                        parts = [result[i:i + 4096] for i in range(0, len(result), 4096)]
                        for part in parts:
                            bot.send_message(chat_id, part, parse_mode='MarkdownV2')
                    else:
                        bot.send_message(chat_id, result, parse_mode='MarkdownV2')
            except Exception as e:
                print(f"Error sending message: {e}")
                bot.send_message(chat_id, "Произошла ошибка при отправке сообщения. Пожалуйста, повторите запрос.")
        elif action == "input_values":
            column_idx = current_step[chat_id]['column_idx']
            headers = column_names.get(chat_id, [])
            if column_idx < len(headers):
                header = headers[column_idx]
                value = text
                row_idx = current_step[chat_id]['row_idx']
                result = write_to_sheet(FIXED_SHEET_NAME, header, value, row_idx)
                bot.send_message(chat_id, result)
                column_idx += 1
                current_step[chat_id]['column_idx'] = column_idx
                if column_idx < len(headers):
                    next_header = headers[column_idx]
                    bot.send_message(chat_id, f"Введите значение для столбца '{escape_markdown_v2(next_header)}':", reply_markup=create_navigation_keyboard())
                else:
                    bot.send_message(chat_id, "Все данные успешно записаны.", reply_markup=create_admin_keyboard())
                    del current_step[chat_id]
    else:
        bot.send_message(chat_id, "Неизвестная команда. Используйте кнопки для навигации.")

# Запуск бота
bot.polling()
