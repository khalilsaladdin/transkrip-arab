# Gunakan base image Python yang ringan
FROM python:3.10-slim

# Set working directory
WORKDIR /app

# Hindari buffering log python
ENV PYTHONUNBUFFERED=1

# Install dependency sistem untuk build PyMuPDF jika diperlukan
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Salin file requirements khusus API
COPY requirements-api.txt .

# Install paket Python
RUN pip install --no-cache-dir -r requirements-api.txt

# Salin source code
COPY api.py .

# Buat folder penyimpanan hasil & temp jika belum ada
RUN mkdir -p temp_output hasil

# Buka port 8000
EXPOSE 8000

# Perintah menjalankan server FastAPI produksi
CMD ["uvicorn", "api:app", "--host", "0.0.0.0", "--port", "8000"]
