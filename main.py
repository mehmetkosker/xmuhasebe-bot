import os
import logging
import pytesseract
from PIL import Image
from telegram import Update
from telegram.ext import Updater, MessageHandler, Filters, CallbackContext
from openpyxl import Workbook
from datetime import datetime
import threading
from flask import Flask

# Logging
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    level=logging.INFO)
logger = logging.getLogger(__name__)

# Flask dummy port (Render requirement)
app = Flask(__name__)

@app.route('/')
def home():
    return "xmuhasebe-bot is running!"

def run_flask():
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 10000)))

# Telegram Bot Token (Environment variable)
TOKEN = os.getenv("BOT_TOKEN")

def process_photo(update: Update, context: CallbackContext):
    try:
        photo_file = update.message.photo[-1].get_file()
        file_path = "photo.jpg"
        photo_file.download(file_path)

        text = pytesseract.image_to_string(Image.open(file_path), lang='eng')
        update.message.reply_text(f"📄 OCR Result:\n{text}")

        # Save to Excel
        wb = Workbook()
        ws = wb.active
        ws.title = "OCR Results"
        ws.append(["Timestamp", "User", "Extracted Text"])
        ws.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), update.message.from_user.username, text])
        wb.save("ocr_results.xlsx")

        update.message.reply_text("✅ Saved to ocr_results.xlsx")

    except Exception as e:
        logger.error(f"Error: {e}")
        update.message.reply_text("⚠️ Error while processing the photo.")

def main():
    updater = Updater(TOKEN, use_context=True)
    dp = updater.dispatcher

    dp.add_handler(MessageHandler(Filters.photo, process_photo))

    # Start Flask and Telegram bot in parallel
    threading.Thread(target=run_flask).start()
    updater.start_polling()
    updater.idle()

if __name__ == "__main__":
    main()
