# -*- coding: utf-8 -*-
"""Ponto de entrada do Extrator em modo desktop.

Sobe o servidor do Streamlit em segundo plano, no próprio processo, e abre o
navegador padrão apontando para ele. Para quem usa, é um programa comum: dois
cliques e a tela aparece — sem terminal, sem instalar Python, sem internet.

O Tesseract vai embutido no pacote; aqui só o colocamos no PATH antes de o app
subir, para o pytesseract encontrá-lo sozinho.
"""

import os
import socket
import sys
import threading
import time
import webbrowser
from pathlib import Path

PORTA_PREFERIDA = 8501


def raiz_recursos() -> Path:
    """Pasta com os arquivos empacotados (o PyInstaller a expõe em _MEIPASS)."""
    empacotado = getattr(sys, "_MEIPASS", None)
    return Path(empacotado) if empacotado else Path(__file__).resolve().parent.parent


def preparar_tesseract() -> None:
    """Deixa o Tesseract embutido visível para o pytesseract, se ele veio junto."""
    pasta = raiz_recursos() / "tesseract"
    binario = pasta / ("tesseract.exe" if os.name == "nt" else "tesseract")
    if not binario.exists():
        return  # sem OCR embutido: o app degrada sozinho e avisa na tela
    os.environ["PATH"] = f"{pasta}{os.pathsep}{os.environ.get('PATH', '')}"
    tessdata = pasta / "tessdata"
    if tessdata.is_dir():
        os.environ["TESSDATA_PREFIX"] = str(tessdata)


def escolher_porta(inicio: int = PORTA_PREFERIDA) -> int:
    """Primeira porta livre a partir de ``inicio`` (evita conflito com outro app)."""
    for porta in range(inicio, inicio + 50):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", porta)) != 0:
                return porta
    return inicio


def abrir_navegador(porta: int) -> None:
    """Espera o servidor responder e então abre o navegador (uma vez só)."""
    endereco = f"http://localhost:{porta}"
    limite = time.time() + 60
    while time.time() < limite:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", porta)) == 0:
                webbrowser.open(endereco)
                return
        time.sleep(0.3)


def main() -> int:
    preparar_tesseract()

    app = raiz_recursos() / "app_web.py"
    if not app.exists():  # execução a partir do código-fonte
        app = Path(__file__).resolve().parent.parent / "app_web.py"

    porta = escolher_porta()
    print(f"Extrator Apoena — abrindo em http://localhost:{porta}")
    print("Feche esta janela para encerrar o programa.")
    threading.Thread(target=abrir_navegador, args=(porta,), daemon=True).start()

    sys.argv = [
        "streamlit", "run", str(app),
        f"--server.port={porta}",
        "--server.address=localhost",
        "--server.headless=true",
        "--global.developmentMode=false",
        "--browser.gatherUsageStats=false",
        "--server.maxUploadSize=500",
    ]
    from streamlit.web import cli as stcli
    return stcli.main()


if __name__ == "__main__":
    sys.exit(main())
