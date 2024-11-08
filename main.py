import requests
from io import BytesIO
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, ContextTypes
from docx import Document

# Токен бота
TOKEN = '7559997572:AAFUYsLavFcPga0zjVyyw007CCFeqhkDOgg'

print("Bot True")

# URL файла .docx
DOCX_URL = 'http://www.bobruisk.belstu.by/uploads/b1/s/8/648/basic/113/836/Zamena_SAYT.docx?t=1731070138'

# Список доступных групп
GROUPS = ["МД23", "РС02-24", "ПМ04-23", "ДП03-24", "ПО6", "ТМ3", "ЛХ16", "МД22", "ПО5", "ТДМ2с"]


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Привет! Выбери команду /select_group для выбора группы.")


async def select_group(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Создаем кнопки для выбора группы из предопределённого списка
    keyboard = [[InlineKeyboardButton(group, callback_data=group)] for group in GROUPS]
    reply_markup = InlineKeyboardMarkup(keyboard)

    await update.message.reply_text("Выбери группу:", reply_markup=reply_markup)


def extract_info_from_docx(group):
    # Скачиваем и открываем .docx файл
    response = requests.get(DOCX_URL)
    file = BytesIO(response.content)
    document = Document(file)

    response_text = ""  # Начальный текст ответа, пока пустой
    collecting_data = False  # Флаг для сбора данных по группе
    data_found = False  # Флаг, чтобы добавить заголовок группы только один раз

    # Проходим по таблицам документа
    for table in document.tables:
        for row in table.rows:
            cells = row.cells

            # Проверяем, что строка содержит достаточное количество ячеек для извлечения данных
            if len(cells) >= 4:
                first_cell_text = cells[0].text.strip()

                # Если текущая строка содержит выбранную группу, начинаем сбор данных
                if first_cell_text == group and not data_found:
                    # Добавляем заголовок группы только один раз
                    response_text += f"Группа: {group}\n"
                    data_found = True
                    collecting_data = True  # Начинаем собирать данные для этой группы

                    # Получаем данные для строки
                    para = cells[1].text.strip() or ""  # Пара
                    aud = cells[2].text.strip() or ""  # Аудитория
                    subject = cells[3].text.strip() or ""  # Учебный предмет
                    response_text += f"Пара: {para}\nАуд: {aud}\nУчебный предмет: {subject}\n"

                # Если строка содержит "-//-" и мы уже собираем данные для выбранной группы
                elif first_cell_text == "-//-" and collecting_data:
                    para = cells[1].text.strip() or ""
                    aud = cells[2].text.strip() or ""
                    subject = cells[3].text.strip() or ""
                    response_text += f"Пара: {para}\nАуд: {aud}\nУчебный предмет: {subject}\n"

                # Если строка не относится к текущей группе, прекращаем сбор данных
                elif first_cell_text != group and collecting_data:
                    collecting_data = False  # Прекращаем сбор данных для этой группы

    # Если ничего не найдено для группы
    if not data_found:
        return "Данные для выбранной группы не найдены."

    return response_text.strip()


async def group_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    group = query.data  # Получаем выбранную группу
    await query.answer()  # Отвечаем на запрос

    # Извлекаем информацию из .docx файла по выбранной группе
    response_text = extract_info_from_docx(group)

    # Отправляем результат пользователю
    await query.edit_message_text(response_text)


def main():
    # Создаем приложение с токеном
    application = Application.builder().token(TOKEN).build()

    # Регистрируем обработчики команд
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("select_group", select_group))
    application.add_handler(CallbackQueryHandler(group_selected))

    # Запускаем бота
    application.run_polling()


if __name__ == '__main__':
    main()
