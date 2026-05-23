import os
import telebot
import gspread
import json
from oauth2client.service_account import ServiceAccountCredentials
from groq import Groq
import google.generativeai as genai

# ==========================================
# 1. KONFIGURASI API DARI RENDER
# ==========================================
TOKEN = os.environ.get("TELEGRAM_TOKEN")
GROQ_KEY = os.environ.get("GROQ_KEY")
GEMINI_KEY = os.environ.get("GEMINI_KEY")

bot = telebot.TeleBot(TOKEN)
genai.configure(api_key=GEMINI_KEY)
groq_client = Groq(api_key=GROQ_KEY)
model = genai.GenerativeModel('gemini-1.5-flash')

# ==========================================
# 2. SETUP GOOGLE SHEETS
# ==========================================
try:
    gcp_credentials = json.loads(os.environ.get("GCP_CREDENTIALS"))
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_dict(gcp_credentials, scope)
    client = gspread.authorize(creds)
    
    # Masukkan ID Spreadsheet lo di sini
    SPREADSHEET_ID = "1zXr7An_gdrRBTdpvYX6fPUX7gxVQsIJp-TpnxlF8Tm4"
    sheet = client.open_by_key(SPREADSHEET_ID).sheet1
except Exception as e:
    print(f"Gagal koneksi ke Google Sheets. Error: {e}")

# ==========================================
# 3. PENERIMA PESAN & TRANSKRIP (VOICE)
# ==========================================
@bot.message_handler(content_types=['voice'])
def handle_voice(message):
    bot.reply_to(message, "🎧 Dengerin VN bentar...")
    file_info = bot.get_file(message.voice.file_id)
    downloaded_file = bot.download_file(file_info.file_path)
    
    with open('voice.ogg', 'wb') as new_file:
        new_file.write(downloaded_file)
    
    with open('voice.ogg', 'rb') as audio_file:
        transcription = groq_client.audio.transcriptions.create(
            file=audio_file,
            model="whisper-large-v3",
        )
    
    proses_data(message, transcription.text)

# ==========================================
# 4. PENERIMA PESAN (TEXT)
# ==========================================
@bot.message_handler(func=lambda m: True)
def handle_text(message):
    proses_data(message, message.text)

# ==========================================
# 5. OTAK AGENTIC KEUANGAN (MONEY LOVER)
# ==========================================
def proses_data(message, teks):
    bot.reply_to(message, "🧠 Menganalisis pengeluaran...")
    
    prompt = f"""
    Lo adalah asisten pencatat keuangan pribadi. Ekstrak data dari teks berikut ke dalam format yang spesifik.
    Teks: "{teks}"
    
    Tugas:
    1. Ekstrak menjadi format list dipisahkan lambang pipe (|): Tanggal(YYYY-MM-DD) | Waktu(HH:MM) | Tipe | Kategori Utama | Sub-Kategori | Nominal (angka) | Catatan.
    2. Tipe HANYA BOLEH: Expense, Income, atau Debt/Loan.
    3. Kategori Utama dan Sub-Kategori HARUS DIPILIH dari daftar berikut (jika tidak cocok, gunakan "Other Expense" atau "Uncategorized"):
    
    DAFTAR KATEGORI (Expense):
    - Bills & Utilities (Sub: Electricity Bill, Gas Bill, Internet Bill, Other Utility Bills, Phone Bill, Rentals, Television Bill, Water Bill)
    - Education
    - Entertainment (Sub: Fun Money, Streaming Service)
    - Family (Sub: Home Maintenance, Home Services, Pets)
    - Food & Beverage
    - Gifts & Donations
    - Health & Fitness (Sub: Fitness, Medical Check-up)
    - Insurances
    - Investment
    - Shopping (Sub: Houseware, Makeup, Personal Items)
    - Transportation (Sub: Vehicle Maintenance)
    - Other Expense
    - Outgoing transfer
    - Pay Interest
    
    DAFTAR KATEGORI (Income):
    - Collect Interest
    - Incoming transfer
    - Other Income
    - Salary
    
    4. Jika tidak ada Sub-Kategori, tulis "None". Nominal murni angka tanpa titik.
    Output HANYA list dengan pipe (|), tanpa teks tambahan apapun.
    """
    
    try:
        res = model.generate_content(prompt).text.strip()
        data = [item.strip() for item in res.split('|')]
        
        sheet.append_row(data)
        
        reply_msg = f"✅ **Sukses Dicatat!**\n\nTipe: {data[2]}\nKategori: {data[3]}\n"
        if data[4] != "None":
            reply_msg += f"Sub: {data[4]}\n"
        reply_msg += f"Nominal: Rp {int(data[5]):,}\nCatatan: {data[6]}"
        
        bot.send_message(message.chat.id, reply_msg, parse_mode='Markdown')
        
    except Exception as e:
        bot.send_message(message.chat.id, f"❌ Gagal mencatat. Error: {str(e)}")

print("Bot Keuangan is Online 🚀")
bot.infinity_polling()