import os
import io
import re
import json
import time
import base64
import telebot
import gspread
import pdfplumber
from pdf2image import convert_from_bytes
from telebot import types
from gspread.exceptions import WorksheetNotFound
from oauth2client.service_account import ServiceAccountCredentials
from groq import Groq
from datetime import datetime, timedelta
from collections import defaultdict

# ==========================================
# 1. KONFIGURASI API
# ==========================================
TOKEN          = "8648856895:AAHU6P-ameMsCxX7NI1NHxHGPSZE8DixcPc"
GROQ_KEY       = "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"

bot         = telebot.TeleBot(TOKEN)
groq_client = Groq(api_key=GROQ_KEY)

# ==========================================
# 2. MASTER DATA KATEGORI (Money Lover style)
# ==========================================
KATEGORI_MAP = {
    "Expense": {
        "Bills & Utilities":  ["Electricity Bill","Gas Bill","Internet Bill","Phone Bill","Rentals","Television Bill","Water Bill","Other Utility Bills"],
        "Education":          ["Tuition","Books & Supplies","Course","Other Education"],
        "Entertainment":      ["Fun Money","Streaming Service","Games","Movies","Other Entertainment"],
        "Family":             ["Home Maintenance","Home Services","Pets","Baby & Kids","Other Family"],
        "Food & Beverage":    ["Dining Out","Groceries","Coffee","Snacks","Other Food"],
        "Gifts & Donations":  ["Gift","Donation","Charity","Other Gifts"],
        "Health & Fitness":   ["Fitness","Medical Check-up","Medicine","Hospital","Other Health"],
        "Insurance":          ["Life Insurance","Health Insurance","Vehicle Insurance","Other Insurance"],
        "Investment":         ["Stocks","Crypto","Mutual Fund","Gold","Other Investment"],
        "Shopping":           ["Clothes","Electronics","Houseware","Makeup","Personal Items","Other Shopping"],
        "Transportation":     ["Fuel","Parking","Public Transport","Ride-hailing","Vehicle Maintenance","Other Transport"],
        "Other Expense":      ["Uncategorized"],
    },
    "Income": {
        "Salary":             ["Monthly Salary","Weekly Wage","Other Salary"],
        "Freelance":          ["Project","Gig","Commission","Other Freelance"],
        "Business":           ["Sales","Service Revenue","Other Business"],
        "Bonus":              ["Performance Bonus","Holiday Bonus","Other Bonus"],
        "Investment Return":  ["Dividend","Interest","Capital Gain","Other Return"],
        "Other Income":       ["Gift Received","Refund","Other Income"],
    },
    "Debt/Loan": {
        "Borrow Money":       ["From Friend","From Family","From Bank","Other Borrow"],
        "Lend Money":         ["To Friend","To Family","Other Lend"],
        "Debt Repayment":     ["Repay Friend","Repay Family","Repay Bank","Other Repay"],
    }
}

DOMPET_LIST = ["Cash","Bank BCA","Bank Mandiri","Bank BRI","Bank BNI","Bank BSI","GoPay","OVO","Dana","ShopeePay","Jago","Jenius","Other"]
HEADERS     = ["📅 Tanggal","⏰ Waktu","📂 Tipe","🏷️ Kategori","🔖 Sub-Kategori","💰 Nominal (Rp)","📝 Catatan","🏦 Dompet"]
COL_WIDTHS  = [110, 80, 110, 170, 170, 140, 260, 130]

# ==========================================
# 3. SETUP GOOGLE SHEETS
# ==========================================
BASE_DIR        = os.path.dirname(os.path.abspath(__file__))
file_kredensial = os.path.join(BASE_DIR, 'credentials.json')
spreadsheet_obj = None

def get_credentials():
    global file_kredensial
    if not os.path.exists(file_kredensial):
        alt = os.path.join(BASE_DIR, 'credentials.json.json')
        if os.path.exists(alt):
            file_kredensial = alt
        else:
            raise FileNotFoundError(f"credentials.json tidak ditemukan di: {BASE_DIR}")
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(file_kredensial, scope)
    return gspread.authorize(creds)

def setup_sheet(ws, spreadsheet):
    ws.clear()
    ws.append_row(HEADERS)
    sid      = ws.id
    requests = []

    requests.append({"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1},
        "cell": {"userEnteredFormat": {
            "backgroundColor": {"red": 0.13, "green": 0.27, "blue": 0.53},
            "textFormat": {"bold": True, "fontSize": 11,
                           "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE"
        }},
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
    }})
    requests.append({"updateSheetProperties": {
        "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": 1}},
        "fields": "gridProperties.frozenRowCount"
    }})
    for i, w in enumerate(COL_WIDTHS):
        requests.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": i, "endIndex": i+1},
            "properties": {"pixelSize": w},
            "fields": "pixelSize"
        }})
    requests.append({"setBasicFilter": {
        "filter": {"range": {"sheetId": sid, "startRowIndex": 0, "startColumnIndex": 0, "endColumnIndex": 8}}
    }})

    tipe_values = [{"userEnteredValue": t} for t in KATEGORI_MAP.keys()]
    requests.append({"setDataValidation": {
        "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": 1000, "startColumnIndex": 2, "endColumnIndex": 3},
        "rule": {"condition": {"type": "ONE_OF_LIST", "values": tipe_values}, "strict": True, "showCustomUi": True}
    }})

    all_cats = []
    for cats in KATEGORI_MAP.values():
        all_cats.extend(cats.keys())
    cat_values = [{"userEnteredValue": c} for c in sorted(set(all_cats))]
    requests.append({"setDataValidation": {
        "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": 1000, "startColumnIndex": 3, "endColumnIndex": 4},
        "rule": {"condition": {"type": "ONE_OF_LIST", "values": cat_values}, "strict": False, "showCustomUi": True}
    }})

    all_subs = []
    for cats in KATEGORI_MAP.values():
        for subs in cats.values():
            all_subs.extend(subs)
    sub_values = [{"userEnteredValue": s} for s in sorted(set(all_subs))]
    requests.append({"setDataValidation": {
        "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": 1000, "startColumnIndex": 4, "endColumnIndex": 5},
        "rule": {"condition": {"type": "ONE_OF_LIST", "values": sub_values}, "strict": False, "showCustomUi": True}
    }})

    dompet_values = [{"userEnteredValue": d} for d in DOMPET_LIST]
    requests.append({"setDataValidation": {
        "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": 1000, "startColumnIndex": 7, "endColumnIndex": 8},
        "rule": {"condition": {"type": "ONE_OF_LIST", "values": dompet_values}, "strict": False, "showCustomUi": True}
    }})

    spreadsheet.batch_update({"requests": requests})

def get_or_create_sheet(spreadsheet, bulan_label):
    try:
        ws = spreadsheet.worksheet(bulan_label)
    except WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=bulan_label, rows=1000, cols=10)
        setup_sheet(ws, spreadsheet)
        print(f"✅ Sheet baru dibuat: {bulan_label}")
    return ws

def init_spreadsheet():
    global spreadsheet_obj
    try:
        client          = get_credentials()
        SPREADSHEET_ID  = "1zXr7An_gdrRBTdpvYX6fPUX7gxVQsIJp-TpnxlF8Tm4"
        spreadsheet_obj = client.open_by_key(SPREADSHEET_ID)
        bulan_ini       = datetime.now().strftime("%B %Y")
        get_or_create_sheet(spreadsheet_obj, bulan_ini)
        print(f"✅ Berhasil terhubung ke Google Sheets! Sheet aktif: {bulan_ini}")
    except Exception as e:
        print(f"❌ Gagal koneksi ke Google Sheets: {e}")
        import traceback; traceback.print_exc()

init_spreadsheet()

# ==========================================
# 4. WARNA BARIS OTOMATIS
# ==========================================
def warnai_baris(spreadsheet, ws, row_index, tipe):
    warna = {
        "income":    {"red": 0.85, "green": 0.96, "blue": 0.85},
        "expense":   {"red": 1.00, "green": 0.90, "blue": 0.90},
        "debt/loan": {"red": 1.00, "green": 0.96, "blue": 0.80},
    }.get(tipe.lower(), {"red": 1, "green": 1, "blue": 1})
    try:
        spreadsheet.batch_update({"requests": [{"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": row_index-1, "endRowIndex": row_index,
                      "startColumnIndex": 0, "endColumnIndex": 8},
            "cell": {"userEnteredFormat": {"backgroundColor": warna}},
            "fields": "userEnteredFormat.backgroundColor"
        }}]})
    except Exception as e:
        print(f"Gagal warnai baris: {e}")

# ==========================================
# 5. PENDING STATE (untuk konfirmasi foto/dokumen)
# ==========================================
# { chat_id: { "transactions": [...], "source": "screenshot|PDF|CSV" } }
pending_transactions: dict = {}

# { chat_id: True }  → sedang menunggu teks koreksi dari user
waiting_edit: dict = {}

# ==========================================
# 6. HELPER: GROQ VISION (OCR gambar)
# ==========================================
def _gemini_vision(image_bytes: bytes, mime: str = "image/jpeg", waktu_kirim: str = None) -> str:
    """Kirim gambar ke Groq Llama-4 Scout (vision), kembalikan raw JSON string transaksi."""
    b64 = base64.b64encode(image_bytes).decode()
    hari_ini   = datetime.now().strftime("%Y-%m-%d")
    tahun_ini  = datetime.now().strftime("%Y")
    waktu_kini = waktu_kirim if waktu_kirim else datetime.now().strftime("%H:%M")
    prompt_text = (
        "Kamu adalah asisten ekstraksi transaksi keuangan.\n"
        f"Hari ini: {hari_ini}. Tahun sekarang: {tahun_ini}. Waktu sekarang: {waktu_kini} WIB.\n"
        "Baca gambar ini (bisa screenshot Shopee/Tokopedia/Lazada/Blibli, "
        "bukti transfer, mutasi GoPay/OVO/Dana/ShopeePay, QRIS, struk, invoice, dll).\n\n"
        "Ekstrak SEMUA transaksi keuangan yang ada.\n"
        "Balas HANYA JSON array valid, tanpa markdown, tanpa backtick:\n"
        "[\n"
        "  {\n"
        '    "tanggal": "YYYY-MM-DD",\n'
        '    "waktu": "HH:MM",\n'
        '    "deskripsi": "nama item / merchant / keterangan asli dari gambar",\n'
        '    "nominal": 12345,\n'
        '    "tipe_raw": "debit|kredit|pembelian|pengembalian",\n'
        '    "sumber": "Shopee|Tokopedia|BCA|Mandiri|GoPay|OVO|Dana|dll"\n'
        "  }\n"
        "]\n\n"
        "Aturan:\n"
        "- nominal selalu angka positif tanpa titik/koma\n"
        f"- tanggal: ambil dari gambar; kalau tahun tidak tercantum gunakan {tahun_ini}\n"
        f"- kalau tanggal tidak ada sama sekali, pakai {hari_ini}\n"
        "- JANGAN pernah pakai tahun sebelum 2024 kecuali gambar memang secara eksplisit menampilkan tahun lama\n"
        f"- waktu: ambil dari gambar kalau ada; kalau tidak ada atau tidak jelas, gunakan {waktu_kini}\n"
        "- JANGAN pakai 00:00 sebagai waktu default\n"
        "- kalau banyak item dalam satu order, pisah per item\n"
        "- kalau tidak ada transaksi sama sekali, balas: []"
    )
    resp = groq_client.chat.completions.create(
        model="meta-llama/llama-4-scout-17b-16e-instruct",
        messages=[{
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
                {"type": "text", "text": prompt_text}
            ]
        }],
        temperature=0.1,
        max_tokens=2048,
    )
    return resp.choices[0].message.content.strip()

# ==========================================
# 7. HELPER: KATEGORISASI via Groq
# ==========================================
# 7a. RULE-BASED KATEGORISASI (0 token, instant)
# ==========================================
# Format: keyword (lowercase) -> (tipe, kategori, sub_kategori, dompet)
_RULES: list[tuple[str, str, str, str, str]] = [
    # ── INCOME ──────────────────────────────────────────────────────────
    ("bunga tabungan",      "Income",    "Investment Return", "Interest",        "Bank BCA"),
    ("bunga",               "Income",    "Investment Return", "Interest",        "Other"),
    ("cashback",            "Income",    "Other Income",      "Refund",          "Other"),
    ("refund",              "Income",    "Other Income",      "Refund",          "Other"),
    ("pengembalian",        "Income",    "Other Income",      "Refund",          "Other"),
    ("gaji",                "Income",    "Salary",            "Monthly Salary",  "Other"),
    ("salary",              "Income",    "Salary",            "Monthly Salary",  "Other"),
    ("dividen",             "Income",    "Investment Return", "Dividend",        "Other"),
    ("transfer masuk",      "Income",    "Other Income",      "Other Income",    "Other"),
    ("top up",              "Income",    "Other Income",      "Other Income",    "Other"),

    # ── FOOD & BEVERAGE ─────────────────────────────────────────────────
    ("mcdonald",            "Expense",   "Food & Beverage",   "Dining Out",      "Other"),
    ("kfc",                 "Expense",   "Food & Beverage",   "Dining Out",      "Other"),
    ("burger",              "Expense",   "Food & Beverage",   "Dining Out",      "Other"),
    ("pizza",               "Expense",   "Food & Beverage",   "Dining Out",      "Other"),
    ("starbucks",           "Expense",   "Food & Beverage",   "Coffee",          "Other"),
    ("kopi",                "Expense",   "Food & Beverage",   "Coffee",          "Other"),
    ("coffee",              "Expense",   "Food & Beverage",   "Coffee",          "Other"),
    ("cafe",                "Expense",   "Food & Beverage",   "Coffee",          "Other"),
    ("resto",               "Expense",   "Food & Beverage",   "Dining Out",      "Other"),
    ("restoran",            "Expense",   "Food & Beverage",   "Dining Out",      "Other"),
    ("warung",              "Expense",   "Food & Beverage",   "Dining Out",      "Other"),
    ("makan",               "Expense",   "Food & Beverage",   "Dining Out",      "Other"),
    ("food",                "Expense",   "Food & Beverage",   "Dining Out",      "Other"),
    ("grab food",           "Expense",   "Food & Beverage",   "Dining Out",      "Other"),
    ("gofood",              "Expense",   "Food & Beverage",   "Dining Out",      "Other"),
    ("shopee food",         "Expense",   "Food & Beverage",   "Dining Out",      "Other"),
    ("indomaret",           "Expense",   "Food & Beverage",   "Groceries",       "Other"),
    ("alfamart",            "Expense",   "Food & Beverage",   "Groceries",       "Other"),
    ("alfamidi",            "Expense",   "Food & Beverage",   "Groceries",       "Other"),
    ("supermarket",         "Expense",   "Food & Beverage",   "Groceries",       "Other"),
    ("hypermart",           "Expense",   "Food & Beverage",   "Groceries",       "Other"),
    ("hero",                "Expense",   "Food & Beverage",   "Groceries",       "Other"),
    ("giant",               "Expense",   "Food & Beverage",   "Groceries",       "Other"),
    ("ranch market",        "Expense",   "Food & Beverage",   "Groceries",       "Other"),
    ("sayur",               "Expense",   "Food & Beverage",   "Groceries",       "Other"),
    ("groceries",           "Expense",   "Food & Beverage",   "Groceries",       "Other"),
    ("snack",               "Expense",   "Food & Beverage",   "Snacks",          "Other"),
    ("jajan",               "Expense",   "Food & Beverage",   "Snacks",          "Other"),
    ("minuman",             "Expense",   "Food & Beverage",   "Snacks",          "Other"),
    ("boba",                "Expense",   "Food & Beverage",   "Coffee",          "Other"),
    ("chatime",             "Expense",   "Food & Beverage",   "Coffee",          "Other"),

    # ── TRANSPORTATION ──────────────────────────────────────────────────
    ("grab",                "Expense",   "Transportation",    "Ride-hailing",    "Other"),
    ("gojek",               "Expense",   "Transportation",    "Ride-hailing",    "Other"),
    ("goride",              "Expense",   "Transportation",    "Ride-hailing",    "Other"),
    ("gocar",               "Expense",   "Transportation",    "Ride-hailing",    "Other"),
    ("maxim",               "Expense",   "Transportation",    "Ride-hailing",    "Other"),
    ("indriver",            "Expense",   "Transportation",    "Ride-hailing",    "Other"),
    ("ojek",                "Expense",   "Transportation",    "Ride-hailing",    "Other"),
    ("taxi",                "Expense",   "Transportation",    "Ride-hailing",    "Other"),
    ("transjakarta",        "Expense",   "Transportation",    "Public Transport", "Other"),
    ("commuter",            "Expense",   "Transportation",    "Public Transport", "Other"),
    ("krl",                 "Expense",   "Transportation",    "Public Transport", "Other"),
    ("mrt",                 "Expense",   "Transportation",    "Public Transport", "Other"),
    ("lrt",                 "Expense",   "Transportation",    "Public Transport", "Other"),
    ("kereta",              "Expense",   "Transportation",    "Public Transport", "Other"),
    ("kai",                 "Expense",   "Transportation",    "Public Transport", "Other"),
    ("bus",                 "Expense",   "Transportation",    "Public Transport", "Other"),
    ("bbm",                 "Expense",   "Transportation",    "Fuel",            "Other"),
    ("bensin",              "Expense",   "Transportation",    "Fuel",            "Other"),
    ("pertamina",           "Expense",   "Transportation",    "Fuel",            "Other"),
    ("shell",               "Expense",   "Transportation",    "Fuel",            "Other"),
    ("spbu",                "Expense",   "Transportation",    "Fuel",            "Other"),
    ("parkir",              "Expense",   "Transportation",    "Parking",         "Other"),
    ("parking",             "Expense",   "Transportation",    "Parking",         "Other"),
    ("tol",                 "Expense",   "Transportation",    "Other Transport",  "Other"),
    ("etoll",               "Expense",   "Transportation",    "Other Transport",  "Other"),
    ("e-toll",              "Expense",   "Transportation",    "Other Transport",  "Other"),
    ("damri",               "Expense",   "Transportation",    "Public Transport", "Other"),

    # ── SHOPPING ────────────────────────────────────────────────────────
    ("shopee",              "Expense",   "Shopping",          "Other Shopping",   "ShopeePay"),
    ("tokopedia",           "Expense",   "Shopping",          "Other Shopping",   "Other"),
    ("lazada",              "Expense",   "Shopping",          "Other Shopping",   "Other"),
    ("blibli",              "Expense",   "Shopping",          "Other Shopping",   "Other"),
    ("tiktok shop",         "Expense",   "Shopping",          "Other Shopping",   "Other"),
    ("bukalapak",           "Expense",   "Shopping",          "Other Shopping",   "Other"),
    ("amazon",              "Expense",   "Shopping",          "Other Shopping",   "Other"),
    ("fashion",             "Expense",   "Shopping",          "Clothes",          "Other"),
    ("baju",                "Expense",   "Shopping",          "Clothes",          "Other"),
    ("sepatu",              "Expense",   "Shopping",          "Clothes",          "Other"),
    ("tas",                 "Expense",   "Shopping",          "Clothes",          "Other"),
    ("kosmetik",            "Expense",   "Shopping",          "Makeup",           "Other"),
    ("skincare",            "Expense",   "Shopping",          "Personal Items",   "Other"),
    ("guardian",            "Expense",   "Shopping",          "Personal Items",   "Other"),
    ("watson",              "Expense",   "Shopping",          "Personal Items",   "Other"),
    ("ikea",                "Expense",   "Shopping",          "Houseware",        "Other"),
    ("ace hardware",        "Expense",   "Shopping",          "Houseware",        "Other"),

    # ── BILLS & UTILITIES ───────────────────────────────────────────────
    ("token listrik",       "Expense",   "Bills & Utilities", "Electricity Bill", "Other"),
    ("listrik",             "Expense",   "Bills & Utilities", "Electricity Bill", "Other"),
    ("pln",                 "Expense",   "Bills & Utilities", "Electricity Bill", "Other"),
    ("pdam",                "Expense",   "Bills & Utilities", "Water Bill",       "Other"),
    ("air",                 "Expense",   "Bills & Utilities", "Water Bill",       "Other"),
    ("indihome",            "Expense",   "Bills & Utilities", "Internet Bill",    "Other"),
    ("biznet",              "Expense",   "Bills & Utilities", "Internet Bill",    "Other"),
    ("myrepublic",          "Expense",   "Bills & Utilities", "Internet Bill",    "Other"),
    ("firstmedia",          "Expense",   "Bills & Utilities", "Internet Bill",    "Other"),
    ("internet",            "Expense",   "Bills & Utilities", "Internet Bill",    "Other"),
    ("wifi",                "Expense",   "Bills & Utilities", "Internet Bill",    "Other"),
    ("telkom",              "Expense",   "Bills & Utilities", "Phone Bill",       "Other"),
    ("telkomsel",           "Expense",   "Bills & Utilities", "Phone Bill",       "Other"),
    ("xl",                  "Expense",   "Bills & Utilities", "Phone Bill",       "Other"),
    ("indosat",             "Expense",   "Bills & Utilities", "Phone Bill",       "Other"),
    ("tri",                 "Expense",   "Bills & Utilities", "Phone Bill",       "Other"),
    ("smartfren",           "Expense",   "Bills & Utilities", "Phone Bill",       "Other"),
    ("pulsa",               "Expense",   "Bills & Utilities", "Phone Bill",       "Other"),
    ("paket data",          "Expense",   "Bills & Utilities", "Phone Bill",       "Other"),
    ("netflix",             "Expense",   "Bills & Utilities", "Television Bill",  "Other"),
    ("spotify",             "Expense",   "Entertainment",     "Streaming Service","Other"),
    ("youtube premium",     "Expense",   "Entertainment",     "Streaming Service","Other"),
    ("disney",              "Expense",   "Entertainment",     "Streaming Service","Other"),
    ("vidio",               "Expense",   "Entertainment",     "Streaming Service","Other"),
    ("rent",                "Expense",   "Bills & Utilities", "Rentals",          "Other"),
    ("sewa",                "Expense",   "Bills & Utilities", "Rentals",          "Other"),
    ("kos",                 "Expense",   "Bills & Utilities", "Rentals",          "Other"),
    ("kontrakan",           "Expense",   "Bills & Utilities", "Rentals",          "Other"),

    # ── HEALTH & FITNESS ────────────────────────────────────────────────
    ("apotek",              "Expense",   "Health & Fitness",  "Medicine",         "Other"),
    ("apotik",              "Expense",   "Health & Fitness",  "Medicine",         "Other"),
    ("kimia farma",         "Expense",   "Health & Fitness",  "Medicine",         "Other"),
    ("k24",                 "Expense",   "Health & Fitness",  "Medicine",         "Other"),
    ("obat",                "Expense",   "Health & Fitness",  "Medicine",         "Other"),
    ("halodoc",             "Expense",   "Health & Fitness",  "Medical Check-up", "Other"),
    ("alodokter",           "Expense",   "Health & Fitness",  "Medical Check-up", "Other"),
    ("rumah sakit",         "Expense",   "Health & Fitness",  "Hospital",         "Other"),
    ("klinik",              "Expense",   "Health & Fitness",  "Medical Check-up", "Other"),
    ("rs ",                 "Expense",   "Health & Fitness",  "Hospital",         "Other"),
    ("dokter",              "Expense",   "Health & Fitness",  "Medical Check-up", "Other"),
    ("gym",                 "Expense",   "Health & Fitness",  "Fitness",          "Other"),
    ("fitness",             "Expense",   "Health & Fitness",  "Fitness",          "Other"),

    # ── EDUCATION ───────────────────────────────────────────────────────
    ("udemy",               "Expense",   "Education",         "Course",           "Other"),
    ("coursera",            "Expense",   "Education",         "Course",           "Other"),
    ("ruangguru",           "Expense",   "Education",         "Course",           "Other"),
    ("zenius",              "Expense",   "Education",         "Course",           "Other"),
    ("buku",                "Expense",   "Education",         "Books & Supplies", "Other"),
    ("gramedia",            "Expense",   "Education",         "Books & Supplies", "Other"),
    ("spp",                 "Expense",   "Education",         "Tuition",          "Other"),
    ("ukt",                 "Expense",   "Education",         "Tuition",          "Other"),
    ("kuliah",              "Expense",   "Education",         "Tuition",          "Other"),
    ("sekolah",             "Expense",   "Education",         "Tuition",          "Other"),

    # ── ENTERTAINMENT ───────────────────────────────────────────────────
    ("bioskop",             "Expense",   "Entertainment",     "Movies",           "Other"),
    ("cgv",                 "Expense",   "Entertainment",     "Movies",           "Other"),
    ("xxi",                 "Expense",   "Entertainment",     "Movies",           "Other"),
    ("cinepolis",           "Expense",   "Entertainment",     "Movies",           "Other"),
    ("game",                "Expense",   "Entertainment",     "Games",            "Other"),
    ("steam",               "Expense",   "Entertainment",     "Games",            "Other"),
    ("playstation",         "Expense",   "Entertainment",     "Games",            "Other"),
    ("google play",         "Expense",   "Entertainment",     "Games",            "Other"),
    ("app store",           "Expense",   "Entertainment",     "Games",            "Other"),

    # ── GIFTS & DONATIONS ───────────────────────────────────────────────
    ("donasi",              "Expense",   "Gifts & Donations", "Donation",         "Other"),
    ("donation",            "Expense",   "Gifts & Donations", "Donation",         "Other"),
    ("sedekah",             "Expense",   "Gifts & Donations", "Charity",          "Other"),
    ("zakat",               "Expense",   "Gifts & Donations", "Charity",          "Other"),
    ("infaq",               "Expense",   "Gifts & Donations", "Charity",          "Other"),
    ("hadiah",              "Expense",   "Gifts & Donations", "Gift",             "Other"),
    ("kado",                "Expense",   "Gifts & Donations", "Gift",             "Other"),

    # ── FAMILY ──────────────────────────────────────────────────────────
    ("pampers",             "Expense",   "Family",            "Baby & Kids",      "Other"),
    ("susu",                "Expense",   "Family",            "Baby & Kids",      "Other"),
    ("popok",               "Expense",   "Family",            "Baby & Kids",      "Other"),
    ("pet",                 "Expense",   "Family",            "Pets",             "Other"),
    ("kucing",              "Expense",   "Family",            "Pets",             "Other"),
    ("anjing",              "Expense",   "Family",            "Pets",             "Other"),
    ("drh",                 "Expense",   "Family",            "Pets",             "Other"),
    ("renovasi",            "Expense",   "Family",            "Home Maintenance", "Other"),
    ("plumber",             "Expense",   "Family",            "Home Services",    "Other"),

    # ── INVESTMENT ──────────────────────────────────────────────────────
    ("bibit",               "Expense",   "Investment",        "Mutual Fund",      "Other"),
    ("bareksa",             "Expense",   "Investment",        "Mutual Fund",      "Other"),
    ("reksadana",           "Expense",   "Investment",        "Mutual Fund",      "Other"),
    ("saham",               "Expense",   "Investment",        "Stocks",           "Other"),
    ("crypto",              "Expense",   "Investment",        "Crypto",           "Other"),
    ("bitcoin",             "Expense",   "Investment",        "Crypto",           "Other"),
    ("emas",                "Expense",   "Investment",        "Gold",             "Other"),
    ("pegadaian",           "Expense",   "Investment",        "Gold",             "Other"),
    ("antam",               "Expense",   "Investment",        "Gold",             "Other"),

    # ── INSURANCE ───────────────────────────────────────────────────────
    ("asuransi",            "Expense",   "Insurance",         "Other Insurance",  "Other"),
    ("insurance",           "Expense",   "Insurance",         "Other Insurance",  "Other"),
    ("bpjs",                "Expense",   "Insurance",         "Health Insurance", "Other"),
    ("premi",               "Expense",   "Insurance",         "Other Insurance",  "Other"),

    # ── DEBT/LOAN ────────────────────────────────────────────────────────
    ("cicilan",             "Expense",   "Debt/Loan",         "Debt Repayment",   "Other"),
    ("angsuran",            "Expense",   "Debt/Loan",         "Debt Repayment",   "Other"),
    ("kredit",              "Expense",   "Debt/Loan",         "Debt Repayment",   "Other"),
    ("pinjam",              "Debt/Loan", "Borrow Money",      "From Friend",      "Cash"),
    ("hutang",              "Debt/Loan", "Borrow Money",      "Other Borrow",     "Cash"),
    ("bayar hutang",        "Debt/Loan", "Debt Repayment",    "Repay Friend",     "Cash"),
]

# Dompet mapping berdasarkan keyword sumber transaksi
_DOMPET_RULES: list[tuple[str, str]] = [
    ("seabank",  "Other"),
    ("gopay",    "GoPay"),
    ("ovo",      "OVO"),
    ("dana",     "Dana"),
    ("shopeepay","ShopeePay"),
    ("bca",      "Bank BCA"),
    ("mandiri",  "Bank Mandiri"),
    ("bri",      "Bank BRI"),
    ("bni",      "Bank BNI"),
    ("bsi",      "Bank BSI"),
    ("jago",     "Jago"),
    ("jenius",   "Jenius"),
    ("cash",     "Cash"),
    ("tunai",    "Cash"),
]


def _infer_dompet(sumber: str, deskripsi: str) -> str:
    """Tentukan dompet dari sumber/deskripsi transaksi."""
    teks = (sumber + " " + deskripsi).lower()
    for kw, dompet in _DOMPET_RULES:
        if kw in teks:
            return dompet
    return "Other"


def _kategorikan_rule_based(raw: dict) -> dict | None:
    """
    Kategorikan satu transaksi pakai rule-based.
    Return dict lengkap, atau None kalau tidak ada rule yang cocok.
    """
    desc    = raw.get("deskripsi", "").lower()
    sumber  = raw.get("sumber", "").lower()
    tipe_r  = raw.get("tipe_raw", "debit").lower()
    nominal = int(raw.get("nominal", 0))

    # Tentukan tipe dari tipe_raw dulu (default)
    if tipe_r in ("kredit", "masuk", "pengembalian", "cashback", "bunga"):
        default_tipe = "Income"
    else:
        default_tipe = "Expense"

    # Match rule keyword (prioritas: lebih panjang keyword = lebih spesifik)
    matched = None
    matched_len = 0
    for kw, tipe, kat, sub, dompet_kw in _RULES:
        if kw in desc and len(kw) > matched_len:
            matched = (tipe, kat, sub, dompet_kw)
            matched_len = len(kw)

    if not matched:
        return None

    tipe_final, kategori, sub_kategori, _ = matched

    # Override tipe kalau tipe_raw jelas income (misal "grab" bisa cashback)
    if default_tipe == "Income" and tipe_final == "Expense":
        if tipe_r in ("kredit", "masuk", "cashback", "pengembalian"):
            tipe_final  = "Income"
            kategori    = "Other Income"
            sub_kategori = "Refund"

    dompet = _infer_dompet(sumber, desc)

    # Buat catatan singkat dari deskripsi (max 5 kata)
    words  = raw.get("deskripsi", "-").split()
    catatan = " ".join(words[:5])

    return {
        "tanggal":     raw.get("tanggal"),
        "waktu":       raw.get("waktu", "00:00"),
        "tipe":        tipe_final,
        "kategori":    kategori,
        "sub_kategori":sub_kategori,
        "nominal":     nominal,
        "catatan":     catatan,
        "dompet":      dompet,
    }


def _kategorikan_rule_based_default(raw: dict) -> dict:
    """Fallback kalau tidak ada rule cocok: pakai tipe_raw untuk tipe, Other untuk kategori."""
    tipe_r  = raw.get("tipe_raw", "debit").lower()
    nominal = int(raw.get("nominal", 0))
    words   = raw.get("deskripsi", "-").split()
    catatan = " ".join(words[:5])
    dompet  = _infer_dompet(raw.get("sumber", ""), raw.get("deskripsi", ""))

    if tipe_r in ("kredit", "masuk", "pengembalian", "cashback", "bunga"):
        tipe, kat, sub = "Income", "Other Income", "Other Income"
    else:
        tipe, kat, sub = "Expense", "Other Expense", "Uncategorized"

    return {
        "tanggal":      raw.get("tanggal"),
        "waktu":        raw.get("waktu", "00:00"),
        "tipe":         tipe,
        "kategori":     kat,
        "sub_kategori": sub,
        "nominal":      nominal,
        "catatan":      catatan,
        "dompet":       dompet,
    }


def _kategorikan_batch_ai(batch: list, daftar_str: str, hari_ini: str, waktu_ini: str) -> list:
    """Kategorikan satu batch transaksi via AI (fallback kalau rule-based gagal)."""
    prompt = f"""Kamu adalah asisten kategorisasi transaksi keuangan pribadi.
Hari ini: {hari_ini}, pukul {waktu_ini}.

Daftar transaksi mentah:
{json.dumps(batch, ensure_ascii=False, indent=2)}

Untuk setiap transaksi, tentukan:
- tipe         : Expense / Income / Debt/Loan
- kategori     : pilih dari daftar
- sub_kategori : pilih sub yang paling cocok
- dompet       : pilih dari {json.dumps(DOMPET_LIST)}
- catatan      : ringkasan max 5 kata (dari deskripsi asli)
- tanggal      : pertahankan dari input, format YYYY-MM-DD
- waktu        : pertahankan dari input, format HH:MM
- nominal      : pertahankan dari input (angka positif)

Aturan tipe_raw -> tipe:
- "debit" / "pembelian" / "keluar" -> Expense
- "kredit" / "masuk" / "pengembalian" / "cashback" / "bunga" -> Income
- kalau tidak jelas -> Expense

DAFTAR KATEGORI & SUB:
{daftar_str}

Balas HANYA JSON array valid, tanpa markdown, tanpa backtick.
Jumlah item output HARUS sama dengan jumlah input ({len(batch)} item)."""

    resp = groq_client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.1,
        max_tokens=2048,
    )
    teks = resp.choices[0].message.content.strip()
    if "```" in teks:
        teks = teks.split("```")[1]
        if teks.startswith("json"):
            teks = teks[4:]
    return json.loads(teks.strip())


def _kategorikan(raw_list: list) -> list:
    """
    Kategorikan semua transaksi.
    Strategi hybrid:
      1. Rule-based dulu (0 token, instant) — cocok untuk ~90% transaksi umum
      2. Sisa yang tidak match rule → dikumpulkan dan dikirim ke AI dalam 1 batch
      3. Kalau AI juga gagal (quota habis) → fallback ke default (Other Expense/Income)
    """
    daftar_kat = []
    for tipe, cats in KATEGORI_MAP.items():
        for cat, subs in cats.items():
            daftar_kat.append(f"[{tipe}] {cat} -> sub: {', '.join(subs)}")
    daftar_str = "\n".join(daftar_kat)
    hari_ini   = datetime.now().strftime("%Y-%m-%d")
    waktu_ini  = datetime.now().strftime("%H:%M")

    # Pass 1: rule-based
    results      = []
    ai_needed    = []   # (original_index, raw_dict)
    rule_hit = rule_miss = 0

    for idx, raw in enumerate(raw_list):
        hasil = _kategorikan_rule_based(raw)
        if hasil:
            results.append((idx, hasil))
            rule_hit += 1
        else:
            results.append((idx, None))  # placeholder
            ai_needed.append((idx, raw))
            rule_miss += 1

    print(f"[KAT] Rule-based: {rule_hit} hit, {rule_miss} perlu AI")

    # Pass 2: AI untuk yang tidak match (batch 10)
    if ai_needed:
        BATCH_SIZE = 10
        batches = [ai_needed[i:i+BATCH_SIZE] for i in range(0, len(ai_needed), BATCH_SIZE)]
        for b_num, batch_items in enumerate(batches, 1):
            idxs  = [x[0] for x in batch_items]
            raws  = [x[1] for x in batch_items]
            print(f"[KAT] AI batch {b_num}/{len(batches)} ({len(raws)} transaksi)...")
            try:
                ai_hasil = _kategorikan_batch_ai(raws, daftar_str, hari_ini, waktu_ini)
                for idx, h in zip(idxs, ai_hasil):
                    results[idx] = (idx, h)
            except Exception as e:
                print(f"[KAT] AI batch {b_num} error: {e} — pakai default")
                for idx, raw in zip(idxs, raws):
                    results[idx] = (idx, _kategorikan_rule_based_default(raw))

    # Pastikan semua placeholder terisi (seharusnya tidak ada, tapi safety net)
    final = []
    for idx, h in results:
        if h is None:
            h = _kategorikan_rule_based_default(raw_list[idx])
        final.append(h)

    return final

# ==========================================
# 8. HELPER: PARSE PDF (batch per 3 halaman supaya tidak overload token)
# ==========================================
# Mapping bulan untuk parser SeaBank
_BULAN_MAP = {
    "JAN":"01","FEB":"02","MAR":"03","APR":"04","MEI":"05","MAY":"05",
    "JUN":"06","JUL":"07","AGU":"08","AUG":"08","SEP":"09","OKT":"10",
    "OCT":"10","NOV":"11","DES":"12","DEC":"12"
}
_DATE_NUM_RE = re.compile(
    r'^\(\d{2}\)\s+\(JAN|FEB|MAR|APR|MEI|MAY|JUN|JUL|AGU|AUG|SEP|OKT|OCT|NOV|DES|DEC\)'
    r'(?:\s+\(.+?\))?\s+\([\d.]+\)\s+\([\d.]+\)$',
    re.IGNORECASE
)

def _parse_seabank_regex(full_text: str, tahun: str = "2026") -> list:
    """
    Parser regex murni untuk e-statement SeaBank.
    Format per baris: "DD MMM [nama_extra] NOMINAL SALDO"
    Merchant ada di baris sebelumnya, tipe transaksi di baris sesudahnya.
    Tidak butuh AI, tidak ada limit token, akurat 100%.
    """
    import re as _re
    DATE_NUM = _re.compile(
        r'^(\d{2})\s+(JAN|FEB|MAR|APR|MEI|MAY|JUN|JUL|AGU|AUG|SEP|OKT|OCT|NOV|DES|DEC)'
        r'(?:\s+(.+?))?\s+([\d.]+)\s+([\d.]+)$',
        _re.IGNORECASE
    )
    SKIP_LINES = {
        "REKENING KORAN", "TANGGAL TRANSAKSI KELUAR (IDR) MASUK (IDR) SALDO AKHIR (IDR)",
        "TANGGAL TRANSAKSI", "KELUAR (IDR)", "MASUK (IDR)", "SALDO AKHIR (IDR)",
        "TABUNGAN", "REKENING", "TOTAL:"
    }
    TIPE_LINES = {"Pembayaran", "Transfer", "Bunga", "Top Up - eWallet",
                  "Bill - Token Listrik", "Bill"}

    lines = [l.strip() for l in full_text.splitlines() if l.strip()]
    results = []

    for idx, line in enumerate(lines):
        m = DATE_NUM.match(line)
        if not m:
            continue

        day     = m.group(1)
        mon_str = m.group(2).upper()
        extra   = (m.group(3) or "").strip()
        val1    = int(m.group(4).replace(".",""))
        # val2 = saldo akhir, diabaikan

        if mon_str not in _BULAN_MAP:
            continue
        tanggal = f"{tahun}-{_BULAN_MAP[mon_str]}-{day}"

        # Kumpulkan nama merchant dari baris-baris sebelumnya
        desc_parts = []
        for k in range(idx - 1, max(idx - 5, -1), -1):
            prev = lines[k]
            if DATE_NUM.match(prev):
                break
            if prev in SKIP_LINES or prev in TIPE_LINES:
                break
            if prev.startswith("halaman") or prev.startswith("S/N"):
                break
            desc_parts.insert(0, prev)

        # Tambah extra (nama di tengah baris tanggal)
        if extra and extra not in TIPE_LINES:
            desc_parts.append(extra)

        deskripsi = " ".join(desc_parts).strip() or "Transaksi"

        # Tipe dari baris sesudahnya
        next_line = lines[idx + 1] if idx + 1 < len(lines) else ""
        is_kredit = (
            next_line == "Bunga"
            or "Pengembalian Dana" in deskripsi
            or "Bunga Tabungan" in deskripsi
            or ("Transfer" in next_line and val1 > 500000
                and any(nama in deskripsi for nama in ["JEREMI","HARIS","GITA","M. UDE"]))
        )

        # Skip bunga harian kecil (< 1000) — noise
        if ("Bunga Tabungan" in deskripsi or next_line == "Bunga") and val1 < 1000:
            continue

        tipe_raw = "kredit" if is_kredit else "debit"

        results.append({
            "tanggal":   tanggal,
            "waktu":     "00:00",
            "deskripsi": deskripsi,
            "nominal":   val1,
            "tipe_raw":  tipe_raw,
            "sumber":    "SeaBank"
        })

    return results


def _parse_pdf(file_bytes: bytes) -> list:
    """
    Parse PDF e-statement bank.
    Strategi:
    1. Coba ekstrak teks via pdfplumber → parse regex (cepat, akurat, 0 token)
    2. Fallback: render tiap halaman jadi gambar → kirim ke vision model
       (butuh PyMuPDF atau poppler — install: pip install pymupdf)
    """
    all_results = []
    full_text   = ""
    tahun       = datetime.now().strftime("%Y")

    # --- Jalur 1: ekstrak teks ---
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    full_text += t + "\n"
    except Exception as e:
        print(f"[PDF] pdfplumber error: {e}")

    if full_text.strip():
        # Coba deteksi bank dari teks
        import re as _re
        m_tahun = _re.search(r'(20\d{2})', full_text)
        if m_tahun:
            tahun = m_tahun.group(1)

        if "seabank" in full_text.lower() or "sea bank" in full_text.lower():
            print(f"[PDF] Detected: SeaBank — pakai regex parser")
            all_results = _parse_seabank_regex(full_text, tahun)
            print(f"[PDF] Regex parser: {len(all_results)} transaksi")
            if all_results:
                return all_results

        # Bank lain: fallback ke AI chunk parser
        print(f"[PDF] Bank tidak dikenal, pakai AI chunk parser")
        hari_ini   = datetime.now().strftime("%Y-%m-%d")
        pages_text = [p.strip() for p in full_text.split("\n\n") if p.strip()]
        BATCH = 2
        for i in range(0, len(pages_text), BATCH):
            chunk = "\n\n".join(pages_text[i:i+BATCH])[:2000]
            hasil = _parse_pdf_chunk(chunk, hari_ini, "Bank")
            all_results.extend(hasil)
        if all_results:
            return all_results

    # --- Jalur 2: fallback gambar (butuh PyMuPDF atau poppler) ---
    print("[PDF] Teks kosong, coba render gambar...")
    waktu_kini = datetime.now().strftime("%H:%M")
    images = []

    try:
        import fitz
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        for page in doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(150/72, 150/72), alpha=False)
            images.append(pix.tobytes("png"))
        doc.close()
        print(f"[PDF] PyMuPDF: {len(images)} halaman")
    except ImportError:
        print("[PDF] PyMuPDF tidak ada. Install: pip install pymupdf")
    except Exception as e:
        print(f"[PDF] PyMuPDF error: {e}")

    if not images:
        try:
            from pdf2image import convert_from_bytes as _cfb
            for img in _cfb(file_bytes, dpi=150, fmt="png"):
                buf = io.BytesIO(); img.save(buf, format="PNG"); images.append(buf.getvalue())
            print(f"[PDF] pdf2image: {len(images)} halaman")
        except Exception as e:
            print(f"[PDF] pdf2image error: {e}")

    for i, img_bytes in enumerate(images):
        try:
            print(f"[PDF] Vision halaman {i+1}/{len(images)}...")
            raw = _gemini_vision(img_bytes, mime="image/png", waktu_kirim=waktu_kini)
            lst = json.loads(raw) if raw not in ("[]","[ ]","",None) else []
            all_results.extend(lst)
        except Exception as e:
            print(f"[PDF] Vision halaman {i+1} error: {e}")

    print(f"[PDF] Total transaksi: {len(all_results)}")
    return all_results

# ==========================================
# 9. HELPER: PARSE CSV
# ==========================================
def _parse_csv(file_bytes: bytes) -> list:
    teks    = file_bytes.decode("utf-8", errors="replace")
    lines   = teks.splitlines()[:301]
    sample  = "\n".join(lines)
    hari_ini = datetime.now().strftime("%Y-%m-%d")

    prompt = f"""Kamu adalah asisten parsing CSV mutasi keuangan Indonesia.
Hari ini: {hari_ini}.

Isi CSV:
\"\"\"
{sample}
\"\"\"

Ekstrak SEMUA transaksi. Balas HANYA JSON array valid, tanpa markdown:
[
  {{
    "tanggal": "YYYY-MM-DD",
    "waktu": "HH:MM",
    "deskripsi": "keterangan kolom mutasi",
    "nominal": 12345,
    "tipe_raw": "debit|kredit",
    "sumber": "nama bank/ewallet jika ada"
  }}
]
Kalau tidak ada transaksi, balas: []"""
    resp = groq_client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.1,
        max_tokens=4096,
    )
    teks_resp = resp.choices[0].message.content.strip()
    if "```" in teks_resp:
        teks_resp = teks_resp.split("```")[1]
        if teks_resp.startswith("json"):
            teks_resp = teks_resp[4:]
    return json.loads(teks_resp.strip())

# ==========================================
# 10. HELPER: SIMPAN KE SHEETS
# ==========================================
def _sheets_retry(fn, max_retries=5):
    """Jalankan fn() dengan exponential backoff kalau kena 429."""
    for attempt in range(max_retries):
        try:
            return fn()
        except Exception as e:
            err_str = str(e)
            if "429" in err_str or "Quota" in err_str or "quota" in err_str:
                wait = 2 ** attempt  # 1, 2, 4, 8, 16 detik
                print(f"[SHEETS] Rate limit, retry {attempt+1}/{max_retries} dalam {wait}s...")
                time.sleep(wait)
            else:
                raise
    raise Exception(f"Gagal setelah {max_retries} percobaan (rate limit terus-menerus)")


def _warnai_baris_batch(spreadsheet, ws, row_start: int, tipe_list: list):
    """
    Warnai banyak baris sekaligus dalam satu batch_update.
    row_start: indeks baris pertama (1-based, tapi repeatCell pakai 0-based).
    """
    WARNA_MAP = {
        "income":    {"red": 0.85, "green": 0.96, "blue": 0.85},
        "expense":   {"red": 1.00, "green": 0.90, "blue": 0.90},
        "debt/loan": {"red": 1.00, "green": 0.96, "blue": 0.80},
    }
    requests = []
    for i, tipe in enumerate(tipe_list):
        warna = WARNA_MAP.get(tipe.lower(), {"red": 1, "green": 1, "blue": 1})
        r = row_start + i - 1  # 0-based
        requests.append({"repeatCell": {
            "range": {"sheetId": ws.id,
                      "startRowIndex": r, "endRowIndex": r + 1,
                      "startColumnIndex": 0, "endColumnIndex": 8},
            "cell": {"userEnteredFormat": {"backgroundColor": warna}},
            "fields": "userEnteredFormat.backgroundColor"
        }})
    if requests:
        _sheets_retry(lambda: spreadsheet.batch_update({"requests": requests}))


def _simpan_ke_sheets(transactions: list) -> int:
    """
    Simpan semua transaksi ke Google Sheets secara batch per bulan.
    Menghindari 429 Rate Limit dengan:
    1. append_rows (satu call per bulan, bukan satu call per transaksi)
    2. Satu batch_update untuk semua pewarnaan sekaligus
    3. Exponential backoff kalau masih kena rate limit
    """
    if not spreadsheet_obj:
        return 0

    # Kelompokkan transaksi per bulan
    from collections import defaultdict
    bulan_groups: dict[str, list] = defaultdict(list)

    for data in transactions:
        tanggal      = data.get("tanggal",      datetime.now().strftime("%Y-%m-%d"))
        waktu        = data.get("waktu",         datetime.now().strftime("%H:%M"))
        tipe         = data.get("tipe",          "Expense")
        kategori     = data.get("kategori",      "Other Expense")
        sub_kategori = data.get("sub_kategori",  "Uncategorized")
        nominal      = int(data.get("nominal",   0))
        catatan      = data.get("catatan",       "-")
        dompet       = data.get("dompet",        "Cash")
        try:
            tgl_obj    = datetime.strptime(tanggal, "%Y-%m-%d")
            bulan_nama = tgl_obj.strftime("%B %Y")
        except Exception:
            bulan_nama = datetime.now().strftime("%B %Y")

        bulan_groups[bulan_nama].append({
            "row":  [tanggal, waktu, tipe, kategori, sub_kategori, nominal, catatan, dompet],
            "tipe": tipe
        })

    saved = 0
    for bulan_nama, items in bulan_groups.items():
        ws        = get_or_create_sheet(spreadsheet_obj, bulan_nama)
        rows      = [it["row"] for it in items]
        tipe_list = [it["tipe"] for it in items]

        # Ambil jumlah baris sebelum insert (untuk hitung posisi warna)
        existing  = _sheets_retry(lambda: ws.get_all_values())
        row_start = len(existing) + 1  # baris pertama yang akan ditulis (1-based)

        # Tulis semua baris sekaligus — 1 API call saja
        _sheets_retry(lambda r=rows: ws.append_rows(r, value_input_option="USER_ENTERED"))

        # Warnai semua baris baru sekaligus — 1 API call saja
        _warnai_baris_batch(spreadsheet_obj, ws, row_start, tipe_list)

        saved += len(rows)
        print(f"[SHEETS] {bulan_nama}: {len(rows)} transaksi disimpan (baris {row_start}–{row_start+len(rows)-1})")

    return saved

# ==========================================
# 11. HELPER: RINGKASAN & KEYBOARD
# ==========================================
def _buat_ringkasan(transactions: list, source_label: str) -> str:
    """
    <= 20 transaksi: detail per item.
    > 20 transaksi: ringkasan per kategori (hindari MESSAGE_TOO_LONG).
    """
    from collections import defaultdict
    total_expense = total_income = total_debt = 0
    for t in transactions:
        tipe    = t.get("tipe", "Expense").lower()
        nominal = int(t.get("nominal", 0))
        if tipe == "expense":      total_expense += nominal
        elif tipe == "income":     total_income  += nominal
        else:                      total_debt    += nominal

    header = f"\U0001f4cb *Hasil baca {source_label}* \u2014 {len(transactions)} transaksi\n"

    if len(transactions) <= 20:
        lines = [header]
        for i, t in enumerate(transactions, 1):
            tipe    = t.get("tipe", "Expense")
            emoji   = {"expense": "\U0001f534", "income": "\U0001f7e2", "debt/loan": "\U0001f7e1"}.get(tipe.lower(), "\u26aa")
            nominal = int(t.get("nominal", 0))
            lines.append(
                f"{i}. {emoji} *{t.get('catatan', '-')}*\n"
                f"   {t.get('tanggal','')} | {t.get('kategori','')} > {t.get('sub_kategori','')}\n"
                f"   Rp{nominal:,}  {t.get('dompet','')}"
            )
    else:
        kat_expense = defaultdict(int)
        kat_income  = defaultdict(int)
        tgl_min, tgl_max = "9999", "0000"
        for t in transactions:
            tipe    = t.get("tipe", "Expense").lower()
            nominal = int(t.get("nominal", 0))
            kat     = t.get("kategori", "Other")
            tgl     = t.get("tanggal", "")
            if tgl:
                if tgl < tgl_min: tgl_min = tgl
                if tgl > tgl_max: tgl_max = tgl
            if tipe == "expense":   kat_expense[kat] += nominal
            elif tipe == "income":  kat_income[kat]  += nominal

        lines = [header, f"\U0001f4c5 Periode: {tgl_min} s/d {tgl_max}\n"]
        if kat_expense:
            lines.append("\U0001f534 *Pengeluaran per kategori:*")
            for kat, total in sorted(kat_expense.items(), key=lambda x: -x[1]):
                lines.append(f"   - {kat}: Rp{total:,}")
        if kat_income:
            lines.append("\n\U0001f7e2 *Pemasukan per kategori:*")
            for kat, total in sorted(kat_income.items(), key=lambda x: -x[1]):
                lines.append(f"   - {kat}: Rp{total:,}")

    lines.append("")
    if total_expense: lines.append(f"\U0001f4ca Total Expense : *Rp{total_expense:,}*")
    if total_income:  lines.append(f"\U0001f4c8 Total Income  : *Rp{total_income:,}*")
    if total_debt:    lines.append(f"\U0001f7e1 Total Debt    : *Rp{total_debt:,}*")
    lines.append("\nPilih aksi:")

    msg = "\n".join(lines)
    if len(msg) > 4000:
        msg = msg[:3950] + "\n...(dipotong)\n\nPilih aksi:"
    return msg

def _keyboard_konfirmasi(chat_id: int) -> types.InlineKeyboardMarkup:
    kb = types.InlineKeyboardMarkup(row_width=3)
    kb.add(
        types.InlineKeyboardButton("✅ Konfirmasi", callback_data=f"konfirmasi:{chat_id}"),
        types.InlineKeyboardButton("✏️ Edit",       callback_data=f"edit:{chat_id}"),
        types.InlineKeyboardButton("❌ Batal",       callback_data=f"batal:{chat_id}"),
    )
    return kb

# ==========================================
# 12. FUNGSI CATAT TRANSAKSI (teks/voice)
# ==========================================
def catat_transaksi(message, teks_input, waktu_real=None):
    now       = datetime.now()
    hari_ini  = now.strftime("%Y-%m-%d")
    waktu_ini = waktu_real if waktu_real else now.strftime("%H:%M")

    daftar_kat = []
    for tipe, cats in KATEGORI_MAP.items():
        for cat, subs in cats.items():
            daftar_kat.append(f"[{tipe}] {cat} → sub: {', '.join(subs)}")
    daftar_str = "\n".join(daftar_kat)

    prompt = f"""
Kamu adalah asisten pencatat keuangan pribadi berbahasa Indonesia.
Hari ini: {hari_ini}, pukul tepat {waktu_ini} WIB.

PENTING: Untuk field "waktu", SELALU gunakan "{waktu_ini}" kecuali user menyebut waktu lain secara eksplisit.
JANGAN pernah menggunakan 00:00 atau mengosongkan waktu.

Dari kalimat berikut, ekstrak SEMUA transaksi keuangan yang disebutkan.
Balas HANYA dengan JSON array valid, tanpa markdown, tanpa backtick, tanpa penjelasan.
Jika 1 transaksi, tetap pakai array dengan 1 item.

Format per item:
{{
  "tanggal": "YYYY-MM-DD",
  "waktu": "{waktu_ini}",
  "tipe": "Expense|Income|Debt/Loan",
  "kategori": "pilih dari daftar",
  "sub_kategori": "pilih sub yang paling cocok",
  "nominal": 0,
  "catatan": "deskripsi singkat max 5 kata",
  "dompet": "pilih dari: {', '.join(DOMPET_LIST)}"
}}

DAFTAR KATEGORI & SUB:
{daftar_str}

Kalimat: "{teks_input}"
"""
    teks_respons = ""
    try:
        resp = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
        )
        teks_respons = resp.choices[0].message.content.strip()
        if "```" in teks_respons:
            teks_respons = teks_respons.split("```")[1]
            if teks_respons.startswith("json"):
                teks_respons = teks_respons[4:]
        teks_respons = teks_respons.strip()

        list_data = json.loads(teks_respons)
        if isinstance(list_data, dict):
            list_data = [list_data]

        total        = len(list_data)
        balasan_list = []

        for i, data in enumerate(list_data, 1):
            tanggal      = data.get("tanggal",      hari_ini)
            waktu        = data.get("waktu",         waktu_ini)
            tipe         = data.get("tipe",          "Expense")
            kategori     = data.get("kategori",      "Other Expense")
            sub_kategori = data.get("sub_kategori",  "Uncategorized")
            nominal      = int(data.get("nominal",   0))
            catatan      = data.get("catatan",       "-")
            dompet       = data.get("dompet",        "Cash")

            try:
                tgl_obj    = datetime.strptime(tanggal, "%Y-%m-%d")
                bulan_nama = tgl_obj.strftime("%B %Y")
            except Exception:
                bulan_nama = datetime.now().strftime("%B %Y")

            if spreadsheet_obj:
                ws = get_or_create_sheet(spreadsheet_obj, bulan_nama)
                ws.append_row(
                    [tanggal, waktu, tipe, kategori, sub_kategori, nominal, catatan, dompet],
                    value_input_option="USER_ENTERED"
                )
                all_rows  = ws.get_all_values()
                row_index = len(all_rows)
                warnai_baris(spreadsheet_obj, ws, row_index, tipe)

            emoji_tipe = {"expense": "🔴", "income": "🟢", "debt/loan": "🟡"}.get(tipe.lower(), "⚪")
            nomor      = f"*Transaksi {i}/{total}*\n" if total > 1 else ""
            balasan_list.append(
                f"{nomor}{emoji_tipe} *{tipe}*\n"
                f"━━━━━━━━━━━━━━━━\n"
                f"📅 {tanggal}  ⏰ {waktu}\n"
                f"🏷️ {kategori}\n"
                f"🔖 {sub_kategori}\n"
                f"💰 Rp{nominal:,}\n"
                f"📝 {catatan}\n"
                f"🏦 {dompet}"
            )

        status = "✅ *Tercatat!*\n\n" if spreadsheet_obj else "⚠️ *Sheets offline.*\n\n"
        bot.reply_to(message, status + "\n\n".join(balasan_list), parse_mode="Markdown")

    except json.JSONDecodeError:
        bot.reply_to(message, f"⚠️ Gagal parse AI:\n`{teks_respons}`", parse_mode="Markdown")
    except Exception as e:
        bot.reply_to(message, f"❌ Error: {e}")

# ==========================================
# 13. FUNGSI ANALISIS / QUERY KEUANGAN
# ==========================================
def ambil_semua_data():
    if not spreadsheet_obj:
        return []
    semua = []
    for ws in spreadsheet_obj.worksheets():
        if ws.title == "Sheet1":
            continue
        rows = ws.get_all_values()
        if len(rows) <= 1:
            continue
        for row in rows[1:]:
            if len(row) >= 6:
                semua.append(row)
    return semua

def analisis_keuangan(message, teks):
    semua_data = ambil_semua_data()
    if not semua_data:
        bot.reply_to(message, "📭 Belum ada data transaksi yang tercatat.")
        return

    sample   = semua_data[-200:] if len(semua_data) > 200 else semua_data
    data_str = "\n".join([
        f"{r[0]}|{r[2]}|{r[3]}|{r[4]}|{r[5]}|{r[6]}|{r[7]}"
        for r in sample if len(r) >= 7
    ])
    hari_ini = datetime.now().strftime("%Y-%m-%d")
    prompt   = f"""
Kamu adalah analis keuangan pribadi yang cerdas dan ringkas.
Hari ini: {hari_ini}.

Data transaksi user (format: Tanggal|Tipe|Kategori|Sub|Nominal|Catatan|Dompet):
{data_str}

Pertanyaan user: "{teks}"

ATURAN FORMAT OUTPUT (WAJIB DIIKUTI):
1. Semua angka rupiah ditulis bold: *Rp 150.000* (pakai titik pemisah ribuan)
2. Semua nama kategori/item ditulis bold: *Food & Beverage*
3. Struktur jawaban pakai format ini (sesuaikan dengan pertanyaan):

💸 *Pengeluaran terbesar:* *Nama Kategori* → *Rp X.XXX*
💰 *Pemasukan terbesar:* *Nama Kategori* → *Rp X.XXX*
🏆 *Paling sering:* *Nama Kategori* (X kali) → *Rp X.XXX*
📉 *Pengeluaran terkecil:* *Nama Kategori* → *Rp X.XXX*
🛍️ *Belanja terbesar:* *nama item* → *Rp X.XXX*
📊 *Total pengeluaran:* *Rp X.XXX*
📥 *Total pemasukan:* *Rp X.XXX*

4. Kalau rekap harian/mingguan/bulanan, tampilkan breakdown per kategori dengan format:
   • *Kategori* : *Rp X.XXX*
5. Jangan tulis kalimat panjang. Pakai bullet point pendek.
6. Kalau user tanya per hari/minggu/bulan, filter dulu berdasarkan tanggal sebelum jawab.
7. Akhiri dengan satu baris ringkasan total jika relevan.
8. Maksimal 20 baris.
"""
    try:
        resp = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
        )
        bot.reply_to(message, resp.choices[0].message.content.strip(), parse_mode="Markdown")
    except Exception as e:
        bot.reply_to(message, f"❌ Error analisis: {e}")

# ==========================================
# 14. CLASSIFIER INTENT
# ==========================================
def detect_intent(teks):
    prompt = f"""Kamu classifier intent untuk bot keuangan pribadi.
Klasifikasikan pesan berikut ke salah satu kategori:
- CATAT     : user ingin mencatat transaksi (beli, bayar, terima, pinjam, dll)
- ANALISIS  : user bertanya tentang data keuangannya (berapa total, paling boros, rekap, dll)
- OBROLAN   : percakapan umum yang tidak terkait keuangan

Balas HANYA satu kata: CATAT / ANALISIS / OBROLAN

Pesan: "{teks}"
"""
    try:
        resp = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.0,
            max_tokens=10,
        )
        return resp.choices[0].message.content.strip().upper()
    except Exception:
        return "CATAT"

# ==========================================
# 15. HANDLER VOICE
# ==========================================
@bot.message_handler(content_types=['voice'])
def handle_voice(message):
    bot.reply_to(message, "🎧 Dengerin bentar...")
    try:
        waktu_real      = datetime.now().strftime("%H:%M")
        file_info       = bot.get_file(message.voice.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        voice_path      = os.path.join(BASE_DIR, 'voice.ogg')
        with open(voice_path, 'wb') as f:
            f.write(downloaded_file)
        with open(voice_path, 'rb') as af:
            transcription = groq_client.audio.transcriptions.create(
                file=af,
                model="whisper-large-v3",
                language="id",
                prompt="Transaksi keuangan: QRIS, GoPay, OVO, Dana, ShopeePay, transfer, Alfamart, Indomaret, BCA, Mandiri, BRI, Tokopedia, Shopee, nominal rupiah"
            )
        teks = transcription.text
        bot.reply_to(message, f"🗣️ _\"{teks}\"_", parse_mode="Markdown")
        intent = detect_intent(teks)
        if intent == "ANALISIS":
            analisis_keuangan(message, teks)
        else:
            catat_transaksi(message, teks, waktu_real=waktu_real)
    except Exception as e:
        bot.reply_to(message, f"❌ Gagal transkrip: {e}")

# ==========================================
# 16. HANDLER FOTO (screenshot belanja, QRIS, dll)
# ==========================================
@bot.message_handler(content_types=['photo'])
def handle_photo(message):
    chat_id     = message.chat.id
    msg_loading = bot.reply_to(message, "🔍 Lagi baca gambarnya...")
    try:
        waktu_kirim = datetime.now().strftime("%H:%M")   # catat waktu saat user kirim
        file_id   = message.photo[-1].file_id   # resolusi tertinggi
        file_info = bot.get_file(file_id)
        img_bytes = bot.download_file(file_info.file_path)

        raw_json_str = _gemini_vision(img_bytes, waktu_kirim=waktu_kirim)
        raw_list     = json.loads(raw_json_str) if raw_json_str not in ("[]", "[ ]", "") else []

        if not raw_list:
            bot.edit_message_text(
                "⚠️ Tidak ada transaksi yang terdeteksi di gambar ini.",
                chat_id, msg_loading.message_id
            )
            return

        bot.edit_message_text("⚙️ Mengkategorikan transaksi...", chat_id, msg_loading.message_id)
        transactions = _kategorikan(raw_list)

        pending_transactions[chat_id] = {"transactions": transactions, "source": "screenshot"}
        ringkasan = _buat_ringkasan(transactions, "screenshot")
        bot.edit_message_text(
            ringkasan, chat_id, msg_loading.message_id,
            parse_mode="Markdown",
            reply_markup=_keyboard_konfirmasi(chat_id)
        )
    except Exception as e:
        bot.edit_message_text(f"❌ Gagal baca gambar: {e}", chat_id, msg_loading.message_id)

# ==========================================
# 17. HANDLER DOKUMEN (PDF e-statement, CSV, gambar sebagai file)
# ==========================================
@bot.message_handler(content_types=['document'])
def handle_document(message):
    chat_id  = message.chat.id
    doc      = message.document
    filename = (doc.file_name or "").lower()
    mime     = (doc.mime_type or "").lower()

    msg_loading = bot.reply_to(message, "📂 Lagi baca dokumennya...")
    try:
        waktu_kirim = datetime.now().strftime("%H:%M")
        file_info  = bot.get_file(doc.file_id)
        file_bytes = bot.download_file(file_info.file_path)

        if filename.endswith(".pdf") or "pdf" in mime:
            bot.edit_message_text("📄 Parsing PDF e-statement...", chat_id, msg_loading.message_id)
            raw_list     = _parse_pdf(file_bytes)
            source_label = "PDF e-statement"

        elif filename.endswith(".csv") or "csv" in mime or "text" in mime:
            bot.edit_message_text("📊 Parsing CSV mutasi...", chat_id, msg_loading.message_id)
            raw_list     = _parse_csv(file_bytes)
            source_label = "CSV mutasi"

        elif any(filename.endswith(ext) for ext in (".jpg", ".jpeg", ".png", ".webp")):
            detected_mime = (
                "image/jpeg" if filename.endswith((".jpg", ".jpeg")) else
                "image/png"  if filename.endswith(".png") else
                "image/webp"
            )
            bot.edit_message_text("🖼️ Membaca gambar...", chat_id, msg_loading.message_id)
            raw_json_str = _gemini_vision(file_bytes, detected_mime, waktu_kirim=waktu_kirim)
            raw_list     = json.loads(raw_json_str) if raw_json_str not in ("[]", "") else []
            source_label = "gambar"

        else:
            bot.edit_message_text(
                "⚠️ Format tidak didukung. Kirim PDF, CSV, atau gambar (JPG/PNG).",
                chat_id, msg_loading.message_id
            )
            return

        if not raw_list:
            bot.edit_message_text(
                "⚠️ Tidak ada transaksi yang berhasil diekstrak.",
                chat_id, msg_loading.message_id
            )
            return

        bot.edit_message_text("⚙️ Mengkategorikan transaksi...", chat_id, msg_loading.message_id)
        transactions = _kategorikan(raw_list)

        pending_transactions[chat_id] = {"transactions": transactions, "source": source_label}
        ringkasan = _buat_ringkasan(transactions, source_label)
        bot.edit_message_text(
            ringkasan, chat_id, msg_loading.message_id,
            parse_mode="Markdown",
            reply_markup=_keyboard_konfirmasi(chat_id)
        )
    except Exception as e:
        bot.edit_message_text(f"❌ Error: {e}", chat_id, msg_loading.message_id)

# ==========================================
# 18. HANDLER CALLBACK (tombol inline ✅ ✏️ ❌)
# ==========================================
@bot.callback_query_handler(func=lambda call: call.data.startswith(("konfirmasi:", "edit:", "batal:")))
def handle_callback(call):
    action, cid_str = call.data.split(":", 1)
    chat_id         = int(cid_str)

    if chat_id != call.message.chat.id:
        bot.answer_callback_query(call.id, "⛔ Bukan transaksi kamu!")
        return

    pending = pending_transactions.get(chat_id)

    if action == "konfirmasi":
        if not pending:
            bot.answer_callback_query(call.id, "Tidak ada data pending.")
            return
        bot.edit_message_reply_markup(call.message.chat.id, call.message.message_id, reply_markup=None)
        bot.answer_callback_query(call.id, "⏳ Menyimpan...")
        try:
            saved = _simpan_ke_sheets(pending["transactions"])
            del pending_transactions[chat_id]
            bot.send_message(
                chat_id,
                f"✅ *{saved} transaksi berhasil dicatat ke Google Sheets!*",
                parse_mode="Markdown"
            )
        except Exception as e:
            bot.send_message(chat_id, f"❌ Gagal simpan: {e}")

    elif action == "edit":
        waiting_edit[chat_id] = True
        bot.answer_callback_query(call.id)
        bot.send_message(
            chat_id,
            "✏️ *Mode Edit*\n\n"
            "Ketik koreksi kamu dalam format bebas. Contoh:\n"
            "• _transaksi 2 nominal 85000_\n"
            "• _transaksi 1 dompet GoPay_\n"
            "• _transaksi 3 kategori Food & Beverage sub Dining Out_\n"
            "• _hapus transaksi 4_",
            parse_mode="Markdown"
        )

    elif action == "batal":
        pending_transactions.pop(chat_id, None)
        waiting_edit.pop(chat_id, None)
        bot.answer_callback_query(call.id, "Dibatalkan.")
        bot.edit_message_reply_markup(call.message.chat.id, call.message.message_id, reply_markup=None)
        bot.send_message(chat_id, "❌ Transaksi dibatalkan, tidak ada yang disimpan.")

# ==========================================
# 19. HANDLER TEXT
# ==========================================
@bot.message_handler(commands=['start', 'help'])
def handle_start(message):
    bot.reply_to(message, (
        "👋 Halo! Aku *JayaMoney* 💰\n"
        "Asisten pencatat keuangan otomatis.\n\n"
        "*💾 Catat transaksi:*\n"
        "• _Beli makan siang 25rb_\n"
        "• _Terima gaji 5 juta_\n"
        "• _Bayar listrik 150rb pake GoPay_\n"
        "• _Pinjam ke Budi 200rb_\n\n"
        "*📸 Kirim gambar/dokumen:*\n"
        "• Screenshot order Shopee/Tokopedia\n"
        "• Bukti transfer / QRIS\n"
        "• PDF e-statement bank\n"
        "• CSV mutasi rekening\n\n"
        "*📊 Tanya data keuangan:*\n"
        "• _Pengeluaran minggu ini berapa?_\n"
        "• _Bulan ini paling boros di mana?_\n"
        "• _Rekap hari ini dong_\n\n"
        "Bisa teks atau *voice note* 🎙️\n"
        "Data otomatis masuk Google Sheets per bulan 📅"
    ), parse_mode="Markdown")

@bot.message_handler(commands=['rekap'])
def handle_rekap(message):
    analisis_keuangan(message, "berikan rekap lengkap semua transaksi bulan ini")

@bot.message_handler(func=lambda m: waiting_edit.get(m.chat.id))
def handle_edit_text(message):
    """Handler teks khusus saat user sedang mode edit pending transaksi."""
    chat_id = message.chat.id
    teks    = message.text.strip()
    pending = pending_transactions.get(chat_id)

    if not pending:
        waiting_edit.pop(chat_id, None)
        bot.reply_to(message, "⚠️ Tidak ada data pending untuk diedit.")
        return

    msg_loading = bot.reply_to(message, "⚙️ Memproses koreksi...")
    try:
        current_json = json.dumps(pending["transactions"], ensure_ascii=False, indent=2)
        prompt = f"""Kamu adalah asisten edit data transaksi keuangan.

Data transaksi saat ini:
{current_json}

Instruksi koreksi dari user:
"{teks}"

Terapkan koreksi dan kembalikan SELURUH array transaksi (termasuk yang tidak diubah).
Balas HANYA JSON array valid, tanpa markdown, tanpa backtick."""
        resp = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
        )
        teks_resp = resp.choices[0].message.content.strip()
        if "```" in teks_resp:
            teks_resp = teks_resp.split("```")[1]
            if teks_resp.startswith("json"):
                teks_resp = teks_resp[4:]

        new_transactions = json.loads(teks_resp.strip())
        pending_transactions[chat_id]["transactions"] = new_transactions
        waiting_edit.pop(chat_id, None)

        ringkasan = _buat_ringkasan(new_transactions, pending["source"] + " (edited)")
        bot.edit_message_text(
            ringkasan, chat_id, msg_loading.message_id,
            parse_mode="Markdown",
            reply_markup=_keyboard_konfirmasi(chat_id)
        )
    except Exception as e:
        waiting_edit.pop(chat_id, None)
        bot.edit_message_text(f"❌ Gagal proses koreksi: {e}", chat_id, msg_loading.message_id)

@bot.message_handler(func=lambda m: True)
def handle_text(message):
    teks   = message.text.strip()
    intent = detect_intent(teks)

    if intent == "CATAT":
        catat_transaksi(message, teks)
    elif intent == "ANALISIS":
        analisis_keuangan(message, teks)
    else:
        try:
            resp = groq_client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "user", "content": (
                    f"Kamu adalah JayaMoney, asisten keuangan pribadi yang ramah dan santai. "
                    f"Jawab singkat max 3 kalimat, bahasa sama dengan user.\n\nPesan: \"{teks}\""
                )}],
                temperature=0.7,
            )
            bot.reply_to(message, resp.choices[0].message.content.strip())
        except Exception as e:
            bot.reply_to(message, f"❌ Error: {e}")

# ==========================================
# 20. JALANKAN BOT
# ==========================================
if __name__ == "__main__":
    print("🤖 Bot JayaMoney aktif! Tekan Ctrl+C untuk berhenti.")
    bot.infinity_polling()