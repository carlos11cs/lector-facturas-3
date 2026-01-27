FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    EASYOCR_MODEL_STORAGE_DIRECTORY=/opt/easyocr-models

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    gcc \
    libglib2.0-0 \
    libgl1 \
    libgomp1 \
  && rm -rf /var/lib/apt/lists/*

COPY requirements.txt ./

RUN pip install --no-cache-dir --upgrade pip \
  && pip install --no-cache-dir torch==2.2.2+cpu torchvision==0.17.2+cpu \
     --index-url https://download.pytorch.org/whl/cpu \
  && pip install --no-cache-dir -r requirements.txt

RUN mkdir -p "$EASYOCR_MODEL_STORAGE_DIRECTORY" \
  && python - <<'PY'
import easyocr
import os
os.environ["EASYOCR_MODEL_STORAGE_DIRECTORY"] = os.getenv("EASYOCR_MODEL_STORAGE_DIRECTORY", "/opt/easyocr-models")
# Preload models at build time so runtime doesn't download.
easyocr.Reader(
    ["es", "en"],
    gpu=False,
    model_storage_directory=os.environ["EASYOCR_MODEL_STORAGE_DIRECTORY"],
    download_enabled=True,
)
PY

COPY . .

EXPOSE 8000

CMD ["sh", "-c", "gunicorn app:app --timeout 180 --bind 0.0.0.0:${PORT:-8000}"]
