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


def caminho_do_app() -> Path:
    """Onde está o app: dentro do pacote ou, ao rodar do código-fonte, ao lado."""
    app = raiz_recursos() / "app_web.py"
    if app.exists():
        return app
    return Path(__file__).resolve().parent.parent / "app_web.py"


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


def verificar_pacote() -> int:
    """Modo de diagnóstico (``--verificar``): confirma que o pacote está íntegro.

    Checa as três coisas que podem quebrar sem alarde num executável empacotado:
    as bibliotecas nativas (numpy, OpenCV e afins costumam perder arquivos na
    hora de empacotar), a execução do próprio app, e o OCR embutido. As três
    falham em silêncio quando testadas só pelo servidor no ar — o Streamlit
    responde em /healthz sem sequer executar o script do app.
    """
    falhas = []

    for modulo in ("numpy", "cv2", "pandas", "pdfplumber", "pypdfium2",
                   "openpyxl", "xlsxwriter", "pytesseract", "streamlit"):
        try:
            __import__(modulo)
        except Exception as erro:
            falhas.append(f"import de {modulo}: {erro}")
    if falhas:
        for falha in falhas:
            print(f"FALHOU {falha}")
        return 1
    print("Bibliotecas: todas importam.")

    app = caminho_do_app()
    try:
        from streamlit.testing.v1 import AppTest
        teste = AppTest.from_file(str(app), default_timeout=180).run()
        if teste.exception:
            print(f"FALHOU execução do app: {list(teste.exception)}")
            return 1
    except Exception as erro:
        print(f"FALHOU execução do app: {erro}")
        return 1
    print("App: executa sem erro.")

    preparar_tesseract()
    try:
        import pytesseract
        versao = pytesseract.get_tesseract_version()
        idiomas = sorted(pytesseract.get_languages())
    except Exception as erro:
        print(f"FALHOU OCR: {erro}")
        return 1
    if "por" not in idiomas:
        print(f"FALHOU OCR: falta o idioma português (encontrados: {idiomas})")
        return 1
    print(f"OCR: Tesseract {versao} | idiomas: {', '.join(idiomas)}")
    return 0


def main() -> int:
    if "--verificar" in sys.argv:
        return verificar_pacote()

    preparar_tesseract()
    app = caminho_do_app()
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
