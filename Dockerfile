# Imagem do Extrator para hospedagem própria (Render, Fly.io, Railway, Hugging Face
# Spaces ou um servidor seu). Aqui a imagem base é nossa, então o Tesseract entra
# normalmente — sem o problema de repositório apt expirado da Streamlit Cloud.
FROM python:3.11-slim

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        tesseract-ocr tesseract-ocr-por curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY app_web.py ./

ENV STREAMLIT_SERVER_HEADLESS=true \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false \
    PORT=8501

EXPOSE 8501

HEALTHCHECK --interval=30s --timeout=5s --start-period=40s \
    CMD curl -f "http://localhost:${PORT}/healthz" || exit 1

# Forma shell para respeitar a variável PORT, que Render/Fly/Spaces definem sozinhos
CMD streamlit run app_web.py --server.port=${PORT} --server.address=0.0.0.0
