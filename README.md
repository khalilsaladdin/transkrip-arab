# 📖 Transkripsi & Terjemahan Kitab Arab AI

Sistem AI filologi Islam untuk transkripsi otomatis naskah kitab Arab (gundul maupun berharakat), terjemahan ilmiah bahasa Indonesia, fihris (daftar isi otomatis), pemisah catatan kaki, dan audio pelafalan Qari (Text-to-Speech).

Didukung oleh model **Google Gemini 3.5 Flash** & **FastAPI**.

---

## 🚀 Menjalankan Server API

### Lokal (Development)
```bash
python -m venv venv
source venv/bin/activate  # atau venv\Scripts\activate di Windows
pip install -r requirements-api.txt
uvicorn api:app --host 0.0.0.0 --port 8000 --reload
```

### Docker
```bash
docker build -t transkripsi-arab .
docker run -p 8000:8000 -e GEMINI_API_KEY="YOUR_KEY" transkripsi-arab
```

---

## 📡 Endpoint Utama

- `GET /api/health` - Cek status server
- `POST /api/transcribe-canvas` - Transkripsi cepat dari tangkapan layar canvas pembaca buku
- `POST /api/transcribe-page` - Transkripsi 1 halaman dari PDF atau gambar
- `POST /api/transcribe-book` - Transkripsi batch rentang halaman dan pembuatan file Word (.docx)
- `GET /api/tts?text=...` - Streaming audio pelafalan bahasa Arab
