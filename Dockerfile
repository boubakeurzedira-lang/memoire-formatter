FROM python:3.12-slim
# LibreOffice (calcul des numéros de page du sommaire), poppler (lecture PDF) et polices arabes/latines
RUN apt-get update && apt-get install -y --no-install-recommends \
      libreoffice-writer poppler-utils fonts-liberation fonts-noto-core fonts-noto-naskh-arabic fonts-kacst \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app.py formatter.py ./
ENV PORT=8080
CMD ["sh", "-c", "gunicorn -b 0.0.0.0:${PORT} -w 2 --timeout 240 app:app"]
