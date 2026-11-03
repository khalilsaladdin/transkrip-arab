"""
api.py - REST API Service untuk Transkripsi & Terjemahan Kitab Arab
Dapat diakses oleh aplikasi web apa pun (Misykah, React, Next.js, Vue, mobile apps)
Engine : Google Gemini 3.5 Flash & gTTS
"""
import base64
import hashlib
import io
import os
import re
import shutil
import time
import unicodedata
from pathlib import Path
from typing import Optional, List

from fastapi import FastAPI, File, UploadFile, Form, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, FileResponse, StreamingResponse
from pydantic import BaseModel
import fitz  # PyMuPDF
import requests
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

try:
    from gtts import gTTS
    GTTS_OK = True
except ImportError:
    GTTS_OK = False

app = FastAPI(
    title="API Transkripsi & Terjemahan Kitab Arab",
    description="Layanan AI filologi Islam untuk OCR kitab Arab gundul/berharakat, terjemahan ilmiah, fihris, dan audio qari.",
    version="1.0.0"
)

# Aktifkan CORS agar frontend (Misykah, localhost:3000, localhost:5500, dll.) dapat memanggil API ini
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).resolve().parent
TEMP_ROOT = BASE_DIR / "temp_output"
FINAL_DIR = BASE_DIR / "hasil"
DEFAULT_GEMINI_KEY = os.getenv("GEMINI_API_KEY", "")

TEMP_ROOT.mkdir(parents=True, exist_ok=True)
FINAL_DIR.mkdir(parents=True, exist_ok=True)


# ── Dynamic Model Discovery & Self-Healing ─────────────────────
ACTIVE_MODELS_CACHE = []
CACHE_TIMESTAMP = 0.0

def get_available_gemini_models(api_key: str = "", force_refresh: bool = False) -> list:
    """Mendeteksi secara dinamis model Gemini Vision terbaru yang aktif di akun Google user."""
    global ACTIVE_MODELS_CACHE, CACHE_TIMESTAMP
    now = time.time()
    if not force_refresh and ACTIVE_MODELS_CACHE and (now - CACHE_TIMESTAMP) < 14400:
        return ACTIVE_MODELS_CACHE

    target_key = api_key or DEFAULT_GEMINI_KEY
    discovered = []
    if target_key:
        try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models?key={target_key}"
            resp = requests.get(url, timeout=10)
            if resp.status_code == 200:
                models = resp.json().get("models", [])
                candidates = [
                    m["name"].replace("models/", "")
                    for m in models
                    if "flash" in m.get("name", "")
                    and "generateContent" in m.get("supportedGenerationMethods", [])
                    and "tts" not in m.get("name", "")
                ]
                def sort_key(name: str):
                    m_ver = re.search(r"gemini-(\d+(?:\.\d+)?)", name)
                    ver = float(m_ver.group(1)) if m_ver else 0.0
                    is_main = 1 if ("lite" not in name and "preview" not in name and "omni" not in name and "image" not in name) else 0
                    return (is_main, ver)

                candidates.sort(key=sort_key, reverse=True)
                if candidates:
                    discovered = candidates
        except Exception:
            pass

    fallback = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro", "gemini-3.8-flash", "gemini-3.5-flash"]
    result = [m for m in discovered] if discovered else fallback
    for m in fallback:
        if m not in result:
            result.append(m)

    ACTIVE_MODELS_CACHE = result
    CACHE_TIMESTAMP = now
    return result

# ---------------------------------------------------------------- Helper Functions
def is_heading_line(line: str) -> bool:
    s = line.strip()
    if not s or len(s) > 120:
        return False
    if s.startswith("#") or s.startswith("**[") or s.startswith("---"):
        return True
    keywords = ["البرهان", "فصل", "باب", "المبحث", "المطلب", "مسألة", "المقدمة", "الخاتمة", "فهرس", "قاعدة", "تنبيه"]
    for kw in keywords:
        if s.startswith(kw) or re.match(rf"^(?:فصل|باب|المبحث|المطلب|مسألة|البرهان|قاعدة)\s+[\w\d\u0660-\u0669]", s):
            return True
    return False


def is_footnote_line(line: str) -> bool:
    s = line.strip()
    if not s:
        return False
    return bool(re.match(r"^[\(\[]?[\d\u0660-\u0669]+[\)\]\.\:\-]", s))


def set_paragraph_rtl(p):
    pPr = p._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT


def parse_gemini_kitab_output(text: str):
    """Memisahkan respons AI menjadi blok teks Arab, terjemahan Indonesia, catatan kaki, dan judul bab."""
    arabic_lines = []
    translation_lines = []
    footnotes = []
    headings = []

    is_translating = False
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue

        # 1. Catatan Kaki / Footnotes selalu diprioritaskan
        if is_footnote_line(line):
            is_translating = False
            footnotes.append(line)
            continue

        # 2. Deteksi blok terjemahan bahasa Indonesia
        if line.startswith("[TERJEMAHAN") or line.startswith("[Terjemahan"):
            is_translating = True
            continue
        elif is_translating and any('\u0600' <= c <= '\u06FF' for c in line) and not re.search(r"[A-Za-z]", line):
            # Baris Arab murni tanpa huruf Latin -> kembali ke teks kitab
            is_translating = False

        if is_translating:
            clean_tr = re.sub(r"^Arti\s*:\s*", "", line)
            translation_lines.append(clean_tr)
        else:
            if is_heading_line(line):
                headings.append(re.sub(r"[*#_]", "", line).strip())
            arabic_lines.append(line)

    return {
        "arabic_text": "\n".join(arabic_lines),
        "translation": "\n".join(translation_lines),
        "footnotes": footnotes,
        "headings": headings,
        "raw_text": text
    }


def call_gemini_vision(img_bytes: bytes, include_translation: bool = True, api_key: str = DEFAULT_GEMINI_KEY) -> str:
    b64_img = base64.b64encode(img_bytes).decode("utf-8")

    if include_translation:
        prompt = """
        Ini adalah halaman buku/kitab berbahasa Arab.
        Tolong transkripsikan teks Arab dan berikan terjemahan bahasa Indonesia yang ilmiah:
        1. Salin seluruh teks Arab persis sesuai aslinya dalam format teks Arab Unicode (UTF-8) lengkap dengan harakat/tasykil jika ada.
        2. Tepat di bawah setiap ayat, hadits, atau paragraf teks Arab, sertakan terjemahan bahasa Indonesia yang akurat, ilmiah, dan luwes dengan format:
           [TERJEMAHAN]
           (terjemahan bahasa Indonesia di sini)
        3. Pisahkan antara judul bab, hadits, ayat Al-Qur'an, dan catatan kaki secara rapi.
        4. Dilarang memberikan komentar pengantar tambahan, keluarkan murni teks Arab dan terjemahannya saja.
        """
    else:
        prompt = """
        Ini adalah halaman buku/kitab berbahasa Arab.
        Tolong transkripsikan seluruh teks bahasa Arab pada gambar halaman ini dengan presisi tinggi dan sangat teliti:
        1. Salin seluruh teks persis sesuai aslinya dalam format teks Arab Unicode (UTF-8).
        2. Pertahankan harakat (tasykil) jika ada di naskah asli.
        3. Pisahkan antara judul, teks utama, hadits, dan catatan kaki secara rapi.
        4. Dilarang menerjemahkan atau memberikan komentar/mukadimah, keluarkan murni teks Arab hasil transkripsi saja.
        """

    payload = {
        "contents": [{
            "parts": [
                {"inlineData": {"mimeType": "image/jpeg", "data": b64_img}},
                {"text": prompt}
            ]
        }],
        "generationConfig": {
            "temperature": 0.1,
            "maxOutputTokens": 4096
        }
    }
    headers = {"Content-Type": "application/json"}
    models_to_try = get_available_gemini_models(api_key)
    last_err = "Gagal memproses setelah beberapa kali percobaan ke Gemini."

    for model_name in models_to_try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
        for attempt in range(2):
            try:
                resp = requests.post(url, headers=headers, json=payload, timeout=50)
                if resp.status_code == 200:
                    res_data = resp.json()
                    candidates = res_data.get("candidates", [])
                    if candidates:
                        parts = candidates[0].get("content", {}).get("parts", [])
                        if parts:
                            return parts[0].get("text", "").strip()
                elif resp.status_code in (429, 503):
                    time.sleep(2 * (attempt + 1))
                    continue
                elif resp.status_code == 404:
                    last_err = f"Gemini API Error (404): {resp.text[:180]}"
                    break
                else:
                    last_err = f"Gemini API Error ({resp.status_code}): {resp.text[:180]}"
                    break
            except requests.exceptions.RequestException as e:
                last_err = f"Koneksi ke Gemini gagal: {str(e)}"
                time.sleep(1)

    if "404" in last_err:
        get_available_gemini_models(api_key, force_refresh=True)

    raise HTTPException(status_code=500, detail=last_err)


def build_docx_file(pages_data: list, job_id: str, strip_tashkeel: bool = False) -> Path:
    out_path = FINAL_DIR / f"{job_id}.docx"
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # 1. Halaman Fihris / Daftar Isi
    p_title = doc.add_paragraph()
    set_paragraph_rtl(p_title)
    r_title = p_title.add_run("فهرس الموضوعات (Daftar Isi Kitab)")
    r_title.font.name = "Traditional Arabic"
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(20, 45, 95)

    p_info = doc.add_paragraph()
    set_paragraph_rtl(p_info)
    r_info = p_info.add_run("(Daftar isi otomatis hasil transkripsi AI)")
    r_info.font.name = "Calibri"
    r_info.font.size = Pt(10.5)
    r_info.font.italic = True
    r_info.font.color.rgb = RGBColor(120, 120, 120)

    for item in pages_data:
        p_num = item["page"]
        for heading in item["headings"]:
            p_item = doc.add_paragraph()
            set_paragraph_rtl(p_item)
            p_item.paragraph_format.space_after = Pt(2)
            r_t = p_item.add_run(f"• {heading}")
            r_t.font.name = "Traditional Arabic"
            r_t.font.size = Pt(14)
            r_t.font.bold = True
            r_t.font.color.rgb = RGBColor(25, 45, 85)

            r_dots = p_item.add_run("   ....................................................................   ")
            r_dots.font.name = "Traditional Arabic"
            r_dots.font.size = Pt(10)
            r_dots.font.color.rgb = RGBColor(180, 180, 180)

            r_p = p_item.add_run(f"صفحة {p_num}")
            r_p.font.name = "Traditional Arabic"
            r_p.font.size = Pt(13)
            r_p.font.bold = True
            r_p.font.color.rgb = RGBColor(90, 90, 90)

    doc.add_page_break()

    # 2. Tulis Isi Kitab
    for item in pages_data:
        p_num = item["page"]
        h = doc.add_paragraph()
        set_paragraph_rtl(h)
        run_h = h.add_run(f"--- [ صفحة {p_num} / Halaman {p_num} ] ---")
        run_h.font.name = "Traditional Arabic"
        run_h.font.size = Pt(13)
        run_h.font.bold = True
        run_h.font.color.rgb = RGBColor(120, 120, 120)

        raw_text = item["raw_text"]
        if strip_tashkeel:
            raw_text = re.sub(r"[\u064B-\u065F\u0670]", "", raw_text)

        is_translating = False
        first_trans_line = True
        footnote_seen = False

        for line in raw_text.splitlines():
            line = line.strip()
            if not line:
                continue

            if is_footnote_line(line):
                is_translating = False
                if not footnote_seen:
                    p_sep = doc.add_paragraph()
                    set_paragraph_rtl(p_sep)
                    p_sep.paragraph_format.space_before = Pt(8)
                    p_sep.paragraph_format.space_after = Pt(2)
                    r_sep = p_sep.add_run("ــــــــــــــــــــــــــــــــــــــــ")
                    r_sep.font.color.rgb = RGBColor(170, 170, 170)
                    footnote_seen = True

                p = doc.add_paragraph()
                set_paragraph_rtl(p)
                p.paragraph_format.right_indent = Inches(0.3)
                p.paragraph_format.space_after = Pt(2)
                run = p.add_run(line)
                run.font.name = "Traditional Arabic"
                run.font.size = Pt(12)
                run.font.color.rgb = RGBColor(85, 85, 85)
                continue

            if line.startswith("[TERJEMAHAN") or line.startswith("[Terjemahan"):
                is_translating = True
                first_trans_line = True
                continue
            elif is_translating and any('\u0600' <= c <= '\u06FF' for c in line) and not re.search(r"[A-Za-z]", line):
                is_translating = False

            if is_translating:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.left_indent = Inches(0.3)
                p.paragraph_format.space_after = Pt(3)
                prefix = "Terjemahan: " if first_trans_line else ""
                run = p.add_run(f"{prefix}{line}")
                first_trans_line = False
                run.font.name = "Calibri"
                run.font.size = Pt(11.5)
                run.font.italic = True
                run.font.color.rgb = RGBColor(30, 60, 95)
                continue

            p = doc.add_paragraph()
            set_paragraph_rtl(p)
            if is_heading_line(line):
                clean_t = re.sub(r"[*#_]", "", line).strip()
                p.style = "Heading 1"
                set_paragraph_rtl(p)
                p.paragraph_format.space_before = Pt(12)
                p.paragraph_format.space_after = Pt(6)
                run = p.add_run(clean_t)
                run.font.name = "Traditional Arabic"
                run.font.size = Pt(18)
                run.font.bold = True
                run.font.color.rgb = RGBColor(20, 45, 95)
            else:
                p.paragraph_format.space_after = Pt(4)
                p.paragraph_format.line_spacing = 1.4
                run = p.add_run(line)
                run.font.name = "Traditional Arabic"
                run.font.size = Pt(16)

    doc.save(str(out_path))
    return out_path


# ---------------------------------------------------------------- API Endpoints
@app.get("/api/health")
def health_check():
    return {
        "status": "online",
        "service": "Layanan Transkripsi & Terjemahan Kitab Arab AI",
        "engine": "Google Gemini 3.5 Flash",
        "tts_available": GTTS_OK
    }


@app.post("/api/transcribe-canvas")
async def transcribe_canvas(
    file: UploadFile = File(..., description="File gambar screenshot canvas halaman buku"),
    include_translation: bool = Form(True),
    strip_tashkeel: bool = Form(False),
    gemini_api_key: Optional[str] = Form(None)
):
    """Menerima kiriman tangkapan gambar canvas langsung dari pembaca PDF web (Misykah books.html).
    Sangat cepat karena tidak perlu mengunggah ulang file PDF puluhan MB."""
    key_to_use = (gemini_api_key or "").strip() or os.getenv("GEMINI_API_KEY", "").strip() or DEFAULT_GEMINI_KEY
    if not key_to_use:
        raise HTTPException(
            status_code=401,
            detail="Kunci API Google Gemini belum diatur di server."
        )
    img_bytes = await file.read()
    raw_res = call_gemini_vision(img_bytes, include_translation=include_translation, api_key=key_to_use)

    if strip_tashkeel:
        raw_res = re.sub(r"[\u064B-\u065F\u0670]", "", raw_res)

    parsed = parse_gemini_kitab_output(raw_res)
    return {
        "success": True,
        "data": parsed
    }


@app.post("/api/transcribe-page")
async def transcribe_page(
    file: UploadFile = File(..., description="File PDF atau gambar"),
    page_number: int = Form(1, description="Nomor halaman yang ingin ditranskripsi (1-indexed)"),
    include_translation: bool = Form(True),
    strip_tashkeel: bool = Form(False)
):
    """Mentranskripsikan satu halaman spesifik dari dokumen PDF atau gambar."""
    contents = await file.read()

    # Cek apakah file adalah PDF atau gambar
    if file.filename.lower().endswith(".pdf") or contents.startswith(b"%PDF"):
        try:
            doc = fitz.open(stream=contents, filetype="pdf")
            if page_number < 1 or page_number > doc.page_count:
                raise HTTPException(status_code=400, detail=f"Nomor halaman di luar jangkauan (1-{doc.page_count})")
            page = doc[page_number - 1]
            pix = page.get_pixmap(matrix=fitz.Matrix(200 / 72, 200 / 72), alpha=False)
            img_bytes = pix.tobytes("jpeg")
            del pix
            doc.close()
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Gagal merender halaman PDF: {str(e)}")
    else:
        img_bytes = contents

    raw_res = call_gemini_vision(img_bytes, include_translation=include_translation)
    if strip_tashkeel:
        raw_res = re.sub(r"[\u064B-\u065F\u0670]", "", raw_res)

    parsed = parse_gemini_kitab_output(raw_res)
    return {
        "success": True,
        "page": page_number,
        "data": parsed
    }


@app.post("/api/transcribe-book")
async def transcribe_book(
    file: UploadFile = File(..., description="File PDF lengkap"),
    start_page: int = Form(1),
    end_page: int = Form(5),
    include_translation: bool = Form(True),
    strip_tashkeel: bool = Form(False)
):
    """Mentranskripsikan rentang halaman PDF, dan menghasilkan file Word (.docx) & Teks (.txt) siap unduh."""
    contents = await file.read()
    try:
        doc = fitz.open(stream=contents, filetype="pdf")
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"File PDF tidak valid: {str(e)}")

    total_pages = doc.page_count
    start_p = max(1, min(start_page, total_pages))
    end_p = max(start_p, min(end_page, total_pages))

    job_id = f"job_{hashlib.md5(contents[:2000]).hexdigest()[:8]}_{int(time.time())}"
    pages_results = []

    for p_num in range(start_p, end_p + 1):
        page = doc[p_num - 1]
        pix = page.get_pixmap(matrix=fitz.Matrix(200 / 72, 200 / 72), alpha=False)
        img_bytes = pix.tobytes("jpeg")
        del pix

        raw_res = call_gemini_vision(img_bytes, include_translation=include_translation)
        parsed = parse_gemini_kitab_output(raw_res)
        parsed["page"] = p_num
        pages_results.append(parsed)
        time.sleep(3.2)  # Jeda aman RPM

    doc.close()

    # Susun file Word (.docx)
    docx_path = build_docx_file(pages_results, job_id, strip_tashkeel=strip_tashkeel)

    # Susun file Teks (.txt)
    txt_path = FINAL_DIR / f"{job_id}.txt"
    with open(txt_path, "w", encoding="utf-8") as f:
        for it in pages_results:
            f.write(f"\n--- [ صفحة {it['page']} / Halaman {it['page']} ] ---\n\n")
            f.write(it["raw_text"])
            f.write("\n\n")

    return {
        "success": True,
        "job_id": job_id,
        "total_processed": len(pages_results),
        "start_page": start_p,
        "end_page": end_p,
        "docx_download_url": f"/api/download/{job_id}/docx",
        "txt_download_url": f"/api/download/{job_id}/txt",
        "pages": pages_results
    }


@app.get("/api/download/{job_id}/docx")
def download_docx(job_id: str):
    p = FINAL_DIR / f"{job_id}.docx"
    if not p.exists():
        raise HTTPException(status_code=404, detail="File Word belum siap atau tidak ditemukan.")
    return FileResponse(
        str(p),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f"Transkripsi_Kitab_{job_id}.docx"
    )


@app.get("/api/download/{job_id}/txt")
def download_txt(job_id: str):
    p = FINAL_DIR / f"{job_id}.txt"
    if not p.exists():
        raise HTTPException(status_code=404, detail="File teks tidak ditemukan.")
    return FileResponse(
        str(p),
        media_type="text/plain; charset=utf-8",
        filename=f"Transkripsi_Kitab_{job_id}.txt"
    )


@app.get("/api/tts")
def get_arabic_audio(text: str = Query(..., description="Teks Arab yang ingin dilafalkan")):
    """Menghasilkan audio MP3 pelafalan bahasa Arab dari teks yang diberikan."""
    if not GTTS_OK:
        raise HTTPException(status_code=500, detail="Modul gTTS tidak tersedia.")

    # Filter teks Arab murni
    clean_ar = re.sub(r'\[TERJEMAHAN\].*?(?=(\n\n|\Z))', '', text, flags=re.DOTALL)
    clean_ar = re.sub(r'[*#_—\-–=]', '', clean_ar)
    clean_ar = re.sub(r'[A-Za-z0-9]', '', clean_ar)
    clean_ar = '\n'.join([
        l.strip() for l in clean_ar.splitlines() 
        if any('\u0600' <= c <= '\u06FF' for c in l) and not is_footnote_line(l)
    ])[:1000].strip()

    if not clean_ar:
        raise HTTPException(status_code=400, detail="Tidak ada teks Arab valid untuk diubah ke audio.")

    try:
        tts = gTTS(text=clean_ar, lang='ar', slow=False)
        buf = io.BytesIO()
        tts.write_to_fp(buf)
        buf.seek(0)
        return StreamingResponse(buf, media_type="audio/mpeg")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal membuat audio: {str(e)}")
