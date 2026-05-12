FROM python:3.11-slim

# =========================
# LIBREOFFICE + DEPENDENCIAS
# =========================

RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-calc \
    fonts-dejavu \
    fonts-liberation \
    default-jre \
    && rm -rf /var/lib/apt/lists/*

# =========================
# WORKDIR
# =========================

WORKDIR /app

# =========================
# REQUIREMENTS
# =========================

COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

# =========================
# APP
# =========================

COPY . .

# =========================
# PORT
# =========================

ENV PORT=10000

EXPOSE 10000

# =========================
# START
# =========================

CMD gunicorn app:app --bind 0.0.0.0:$PORT --timeout 300
