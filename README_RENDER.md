# Deploy JayaMoney ke Render

## File yang disiapkan
- `bot.py` — kode bot (API key sudah pakai env var)
- `requirements.txt` — Python dependencies
- `build.sh` — install sistem packages (poppler, ffmpeg, dll)
- `render.yaml` — konfigurasi Render

---

## Langkah Deploy

### 1. Push ke GitHub
Buat repo baru di GitHub, lalu push semua file ini:
```
bot.py
requirements.txt
build.sh
render.yaml
```
> ⚠️ **Jangan** push `credentials.json` ke GitHub. Tambahkan ke `.gitignore`.

---

### 2. Buat Service di Render

1. Buka [render.com](https://render.com) → **New** → **Background Worker**
2. Connect ke repo GitHub kamu
3. Render akan otomatis baca `render.yaml`

Kalau mau manual:
- **Build Command:** `bash build.sh`
- **Start Command:** `python bot.py`

---

### 3. Set Environment Variables

Di Render dashboard → **Environment** → tambahkan 3 variabel:

| Key | Value |
|-----|-------|
| `TELEGRAM_BOT_TOKEN` | Token bot kamu dari BotFather |
| `GROQ_API_KEY` | API key Groq kamu |
| `GOOGLE_CREDENTIALS_JSON` | *Lihat instruksi di bawah* |

#### Cara isi `GOOGLE_CREDENTIALS_JSON`:
Buka file `credentials.json`, copy **seluruh isinya** (dari `{` sampai `}`), paste sebagai value env var.

Contoh value-nya seperti ini (satu baris panjang atau multi-line, keduanya oke):
```json
{"type":"service_account","project_id":"keuangan-497217","private_key_id":"...","private_key":"-----BEGIN PRIVATE KEY-----\n..."}
```

---

### 4. Deploy

Klik **Deploy** → tunggu build selesai (~2-3 menit).

Log sukses akan terlihat seperti:
```
✅ Berhasil terhubung ke Google Sheets! Sheet aktif: May 2026
🤖 Bot JayaMoney aktif! Tekan Ctrl+C untuk berhenti.
```

---

## Troubleshooting

| Error | Solusi |
|-------|--------|
| `TELEGRAM_BOT_TOKEN` not set | Pastikan env var sudah diisi di Render |
| `FileNotFoundError: credentials.json` | Pastikan `GOOGLE_CREDENTIALS_JSON` sudah diisi |
| `poppler not found` | Build command harus `bash build.sh`, bukan `pip install -r requirements.txt` |
| Bot mati sendiri setelah beberapa jam | Normal untuk free tier; upgrade ke paid atau gunakan UptimeRobot |

---

## Catatan Free Tier Render
- Free worker **spin down setelah 15 menit idle** → bot bisa mati
- Solusi: upgrade ke **Starter ($7/bulan)** untuk always-on, atau deploy sebagai **Web Service** dengan endpoint `/health` dan ping pakai UptimeRobot tiap 5 menit
