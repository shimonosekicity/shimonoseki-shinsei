# PDF出力（LibreOffice変換）に対応したデプロイ用イメージ。
# Render では「Language: Docker」を選択するとこのファイルが使われる。
FROM python:3.12-slim

# LibreOffice（docx→PDF変換）と日本語フォント
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-writer \
        fonts-noto-cjk \
        fonts-ipafont-mincho \
        fonts-ipafont-gothic \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV PORT=10000
EXPOSE 10000

CMD gunicorn --bind 0.0.0.0:$PORT --timeout 180 app:app
