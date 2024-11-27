import os
import json
import shutil
import glob
import threading
import time
import asyncio
import aiohttp
from aiohttp import web
from telegram.ext import Application
from datetime import datetime, timedelta
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    CommandHandler,
    CallbackQueryHandler,
    ContextTypes,
    MessageHandler,
    filters,
)
from openpyxl import load_workbook

# Константы
TOKEN = os.getenv('TELEGRAM_BOT_TOKEN', '7514978498:AAF3uWbaKRRaUTrY6g8McYMVsJes1kL6hT4')
if not TOKEN:
    raise ValueError("Необходимо установить TELEGRAM_BOT_TOKEN в переменные окружения.")
PORT = int(os.getenv("PORT", 8443))
WEBHOOK_PATH = "/webhook"
WEBHOOK_URL = f"https://tbot-1-k0fj.onrender.com{WEBHOOK_PATH}"

ADMIN_IDS = [476571220, 39897938]  # ID админов

# Директории и файлы
current_directory = os.path.dirname(os.path.abspath(__file__))
template_file = os.path.join(current_directory, "template.xlsx")
output_file = os.path.join(current_directory, "output.xlsx")
log_file = os.path.join(current_directory, "logs.json")
archive_dir = os.path.join(current_directory, "archive")

# Временные настройки
reset_interval = timedelta(hours=36)
last_reset_time = datetime.now()

temp_data = {}

QUESTIONS = [
    ("Введите Описания.", "description"),
    ("Введите количество.", "quantity"),
    ("Введите количество Диск отрезной 125х2.5мм", "disks_125"),
    ("Введите количество Диск отрезной 180х2.5мм", "disks_180"),
    ("Введите количество Диск шлифовальный d-125", "grinding_125"),
    ("Введите количество Диск шлифовальный d-180.", "grinding_180"),
    ("Введите количество электродов 3 мм.", "electrodes_3mm"),
    ("Введите количество электродов ЛЭЗ УОНИ 13/55 Д-2,5 мм.", "electrodes_uoni"),
]
def log_action(username: str, success: bool):
    """Логирует действие пользователя."""
    now = datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    day_of_week = now.strftime("%A")
    time_str = now.strftime("%H:%M:%S")

    status = "Успешно" if success else "Ошибка"

    log_data = {}
    if os.path.exists(log_file):
        with open(log_file, "r", encoding="utf-8") as f:
            log_data = json.load(f)

    if date_str not in log_data:
        log_data[date_str] = {day_of_week: []}
    elif day_of_week not in log_data[date_str]:
        log_data[date_str][day_of_week] = []

    log_data[date_str][day_of_week].append({
        "username": username,
        "time": time_str,
        "status": status
    })

    with open(log_file, "w", encoding="utf-8") as f:
        json.dump(log_data, f, ensure_ascii=False, indent=2)

def archive_old_file():
    """Архивирует старый файл."""
    if not os.path.exists(output_file):
        return
    if not os.path.exists(archive_dir):
        os.makedirs(archive_dir)
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    archive_name = os.path.join(archive_dir, f"output_{timestamp}.xlsx")
    shutil.copy2(output_file, archive_name)

def clean_old_archives(retain_days: int = 7):
    """Удаляет старые архивы."""
    cutoff_time = time.time() - (retain_days * 86400)
    for file_path in glob.glob(os.path.join(archive_dir, "output_*.xlsx")):
        if os.path.getmtime(file_path) < cutoff_time:
            os.remove(file_path)

def clean_old_logs(retain_days: int = 30):
    """Удаляет старые логи."""
    if not os.path.exists(log_file):
        return
    with open(log_file, "r", encoding="utf-8") as f:
        log_data = json.load(f)
    cutoff_date = (datetime.now() - timedelta(days=retain_days)).strftime("%Y-%m-%d")
    filtered_data = {date: logs for date, logs in log_data.items() if date >= cutoff_date}
    with open(log_file, "w", encoding="utf-8") as f:
        json.dump(filtered_data, f, ensure_ascii=False, indent=2)
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик команды /start."""
    username = update.effective_user.username or "Unknown"
    user_id = update.effective_user.id
    greeting = (
        f"Привет, {username}!\n"
        "Я бот для заполнения таблиц. Вы можете ввести данные для последующего экспорта в Excel."
    )
    await update.message.reply_text(greeting)

async def add_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Начинает процесс добавления данных."""
    username = update.effective_user.username or "Unknown"
    user_id = update.effective_user.id

    if user_id not in temp_data:
        temp_data[user_id] = {"step": 0, "data": {}}

    step = temp_data[user_id]["step"]
    question, field_name = QUESTIONS[step]

    await update.message.reply_text(question)
    temp_data[user_id]["field_name"] = field_name

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает ответы пользователя."""
    username = update.effective_user.username or "Unknown"
    user_id = update.effective_user.id
    message_text = update.message.text

    if user_id not in temp_data:
        await update.message.reply_text("Введите /add для начала ввода данных.")
        return

    step_data = temp_data[user_id]
    field_name = step_data["field_name"]
    step_data["data"][field_name] = message_text

    if step_data["step"] + 1 < len(QUESTIONS):
        step_data["step"] += 1
        next_question, field_name = QUESTIONS[step_data["step"]]
        step_data["field_name"] = field_name
        await update.message.reply_text(next_question)
    else:
        # Завершение ввода данных
        await save_data_to_excel(step_data["data"], username)
        log_action(username, success=True)
        del temp_data[user_id]
        await update.message.reply_text("Данные успешно сохранены в таблице!")

async def save_data_to_excel(data, username):
    """Сохраняет данные пользователя в Excel."""
    if not os.path.exists(output_file):
        shutil.copy2(template_file, output_file)

    workbook = load_workbook(output_file)
    sheet = workbook.active
    next_row = sheet.max_row + 1

    sheet.cell(row=next_row, column=1, value=username)
    for col_index, (_, field_value) in enumerate(data.items(), start=2):
        sheet.cell(row=next_row, column=col_index, value=field_value)

    workbook.save(output_file)
    workbook.close()

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Отмена текущего процесса ввода данных."""
    user_id = update.effective_user.id
    if user_id in temp_data:
        del temp_data[user_id]
        await update.message.reply_text("Ввод данных отменён.")
    else:
        await update.message.reply_text("Нет активного процесса ввода данных.")
async def keep_alive():
    """Keep-alive запросы для Render."""
    while True:
        try:
            async with aiohttp.ClientSession() as session:
                async with session.get(WEBHOOK_URL) as response:
                    if response.status == 200:
                        print("Keep-alive успешен!")
        except Exception as e:
            print(f"Ошибка keep-alive: {e}")
        await asyncio.sleep(120)  # Период keep-alive запросов

async def webhook_handler(request):
    """Обработчик входящих запросов для вебхука."""
    data = await request.json()
    await application.update_queue.put(Update.de_json(data, application.bot))
    return web.Response(text="OK")

def setup_webhook(app):
    """Настройка вебхука."""
    app.router.add_post(WEBHOOK_PATH, webhook_handler)
if __name__ == "__main__":
    # Очистка старых архивов и логов
    archive_old_file()
    clean_old_archives()
    clean_old_logs()

    # Настройка Telegram Bot
    application = Application.builder().token(TOKEN).build()

    # Обработчики команд
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("add", add_data))
    application.add_handler(CommandHandler("cancel", cancel))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    # Запуск вебхука
    loop = asyncio.get_event_loop()
    loop.create_task(keep_alive())
    web_app = web.Application()
    setup_webhook(web_app)

    web.run_app(web_app, host="0.0.0.0", port=PORT)
