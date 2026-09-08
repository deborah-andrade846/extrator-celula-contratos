# Extrator de Relatórios – Apoena

App em Streamlit que extrai dados de PDFs (Notas Fiscais/NAI, Diárias de Hotel,
Exames Ocupacionais e Mapa de Refeições) e gera uma planilha Excel formatada.

## Como rodar localmente

```bash
pip install -r requirements.txt
streamlit run app_web.py
```

Só isso: todas as dependências de `requirements.txt` são wheels Python, sem
pacotes de sistema obrigatórios.

## OCR (opcional)

A leitura padrão usa a **camada de texto** do PDF — é mais rápida e mais precisa
que o OCR. O OCR só entra em PDFs digitalizados (imagem) ou como fallback, e
depende do binário do **Tesseract**, que é um pacote de sistema:

```bash
sudo xargs apt-get install -y < apt-packages.txt   # tesseract-ocr, tesseract-ocr-por
```

Sem o Tesseract instalado o app **continua funcionando**: a opção "OCR Alta
Precisão" aparece desabilitada e os arquivos que exigiriam OCR mostram um aviso
explicando o motivo, em vez de derrubar a execução.

## Deploy na Streamlit Community Cloud

Este repositório **não tem `packages.txt` de propósito**. Quando esse arquivo
existe, a Streamlit Cloud roda `apt-get update` antes de instalar, e a imagem
base tem uma entrada obsoleta de `bullseye-security` cujo `Release` está
expirado. O `apt-get update` retorna erro e o deploy é abortado:

```
E: Release file for http://deb.debian.org/debian-security/dists/bullseye-security/InRelease is expired
❗️ installer returned a non-zero exit code
❗️ Error during processing dependencies!
```

É uma falha da imagem base, não do app — e não há como corrigi-la pelo
repositório. A saída é não depender de apt no deploy:

- `pdf2image` foi removido (não era usado) → `poppler-utils` deixou de ser necessário;
- `pdfplumber>=0.11` rasteriza páginas com `pypdfium2` (wheel puro), sem poppler nem ImageMagick;
- `opencv-python-headless` não precisa de `libgl1`;
- o Tesseract virou opcional, conforme descrito acima.

A lista de pacotes apt continua versionada em `apt-packages.txt` para uso local,
no devcontainer e em Docker. Se quiser OCR também na nuvem, basta copiar esse
arquivo para `packages.txt` **depois** que a Streamlit corrigir a imagem base.
