import os
import re
import pytesseract
import shutil
from PIL import Image
from datetime import datetime
from openpyxl import Workbook, load_workbook
from telegram.ext import Updater, MessageHandler, Filters

# ========================
# 📂 AYARLAR
# ========================
EXCEL_PATH = "XMuhasebe.xlsx"
IMAGE_DIR = "Fisler"
os.makedirs(IMAGE_DIR, exist_ok=True)

# ========================
# 📘 EXCEL DOSYASINI OLUŞTUR
# ========================
if not os.path.exists(EXCEL_PATH):
    wb = Workbook()
    ws = wb.active
    ws.title = "Yönetici"
    ws.append(["ID", "Yükleme Tarihi", "Belge Tarihi", "Firma", "Tutar", "Döviz", "Kullanıcı", "Fotoğraf Yolu"])
    wb.save(EXCEL_PATH)

# ========================
# 🔍 OCR FONKSİYONU
# ========================
def ocr_parse(image_path):
    text = pytesseract.image_to_string(Image.open(image_path), lang='tur')
    date_match = re.search(r'(\d{2}[./-]\d{2}[./-]\d{4})', text)
    doc_date = date_match.group(1) if date_match else ""
    amount_match = re.search(r'(\d+[.,]\d{2})(?!\d)', text)
    amount = amount_match.group(1) if amount_match else ""
    firm_match = re.search(r'[A-ZÇĞİÖŞÜ][A-Za-zÇĞİÖŞÜçğıöşü\s.-]{2,}', text)
    firm = firm_match.group(0).strip() if firm_match else ""
    return text, doc_date, amount, firm

# ========================
# 📸 FOTOĞRAF YÜKLEME İŞLEMLERİ
# ========================
def handle_photo(update, context):
    user = update.message.from_user.username or update.message.from_user.first_name
    file = context.bot.getFile(update.message.photo[-1].file_id)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    image_path = f"{IMAGE_DIR}/fis_{timestamp}.jpg"
    file.download(image_path)

    text, doc_date, amount, firm = ocr_parse(image_path)

    wb = load_workbook(EXCEL_PATH)
    ws = wb.active
    next_id = ws.max_row
    ws.append([
        next_id,
        datetime.now().strftime("%Y-%m-%d %H:%M"),
        doc_date,
        firm,
        amount,
        "TRY",
        user,
        image_path
    ])
    wb.save(EXCEL_PATH)

    update.message.reply_text(
        f"✅ Fiş kaydedildi!\n\n📆 Belge Tarihi: {doc_date}\n🏢 Firma: {firm}\n💰 Tutar: {amount} TRY"
    )

# ========================
# 🤖 TELEGRAM BOT BAŞLATMA
# ========================
def main():
    TOKEN = "YOUR_TELEGRAM_BOT_TOKEN_HERE"  # 🔒 Buraya kendi bot token'ını yaz
    updater = Updater(TOKEN, use_context=True)
    dp = updater.dispatcher

    dp.add_handler(MessageHandler(Filters.photo, handle_photo))
    updater.start_polling()
    print("🤖 XMuhasebe bot aktif (Render). Telegram'dan fotoğraf gönder!")
    updater.idle()

if __name__ == "__main__":
    main()
