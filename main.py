import os
from telegram.ext import Updater, MessageHandler, Filters, CommandHandler
from telegram import Update
from telegram.ext.callbackcontext import CallbackContext
from PIL import Image
import pytesseract
import openpyxl


# --- CONFIG ---
TOKEN = os.getenv("BOT_TOKEN")  # Render Environment Variables'da ayarladığın token

if not TOKEN:
    raise ValueError("❌ BOT_TOKEN tanımlı değil! Render Environment sekmesinde BOT_TOKEN ekle.")


# --- OCR FUNCTION ---
def extract_text_from_image(image_path):
    """Fotoğraftan metni OCR ile çıkarır."""
    try:
        text = pytesseract.image_to_string(Image.open(image_path), lang='eng')
        return text.strip() if text else "Metin tespit edilemedi."
    except Exception as e:
        return f"Hata: {e}"


# --- COMMAND HANDLERS ---
def start(update: Update, context: CallbackContext):
    update.message.reply_text("Merhaba 👋\nBir fotoğraf gönder, içindeki metni OCR ile çıkarayım.")


# --- PHOTO HANDLER ---
def handle_photo(update: Update, context: CallbackContext):
    photo = update.message.photo[-1].get_file()
    photo_path = "received_photo.jpg"
    photo.download(photo_path)

    extracted_text = extract_text_from_image(photo_path)

    # Mesaj olarak gönder
    update.message.reply_text(f"📄 OCR Sonucu:\n\n{extracted_text}")

    # Excel’e kaydet (opsiyonel)
    save_to_excel(extracted_text)


def save_to_excel(text):
    """OCR sonuçlarını excel dosyasına kaydeder (yoksa oluşturur)."""
    file_path = "ocr_results.xlsx"
    if os.path.exists(file_path):
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["OCR Sonucu"])

    ws.append([text])
    wb.save(file_path)


# --- MAIN FUNCTION ---
def main():
    updater = Updater(TOKEN, use_context=True)
    dp = updater.dispatcher

    dp.add_handler(CommandHandler("start", start))
    dp.add_handler(MessageHandler(Filters.photo, handle_photo))

    updater.start_polling()
    print("🤖 Bot çalışıyor...")
    updater.idle()


if __name__ == "__main__":
    main()
