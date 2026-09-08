# -*- mode: python ; coding: utf-8 -*-
"""Receita do PyInstaller para gerar o Extrator como programa de desktop.

Gera uma pasta (modo onedir) com ExtratorApoena.exe. Optamos por onedir em vez de
arquivo único porque o Streamlit carrega recursos do disco: o executável único
teria de descompactar centenas de MB a cada abertura, ficando lento e frágil.

Se a pasta ``desktop/tesseract`` existir no momento do build, ela entra no pacote
e o OCR funciona sem nenhuma instalação na máquina do usuário.
"""

from pathlib import Path

from PyInstaller.utils.hooks import collect_all, copy_metadata

RAIZ = Path(SPECPATH).parent

datas = [(str(RAIZ / "app_web.py"), ".")]
binaries = []
hiddenimports = ["streamlit.runtime.scriptrunner.magic_funcs"]

# Pacotes cujos dados/estáticos/binários precisam ir junto: o Streamlit serve seu próprio
# front-end, o pypdfium2 carrega biblioteca nativa, e numpy/OpenCV/pandas têm extensões em C
# cujos submódulos os hooks automáticos nem sempre alcançam — foi assim que um pacote saiu
# com o numpy quebrado ("No module named 'numpy._core._exceptions'").
for pacote in ("streamlit", "pdfplumber", "pypdfium2", "pytesseract", "altair",
               "numpy", "cv2", "pandas"):
    d, b, h = collect_all(pacote)
    datas += d
    binaries += b
    hiddenimports += h

# O Streamlit consulta a versão instalada de várias bibliotecas em tempo de execução;
# sem os metadados ele quebra ao iniciar.
for pacote in ("streamlit", "pandas", "numpy", "pdfplumber", "pypdfium2", "pytesseract",
               "openpyxl", "xlsxwriter", "opencv-python-headless", "altair", "pillow"):
    try:
        datas += copy_metadata(pacote)
    except Exception:  # pacote ausente no ambiente de build: segue sem ele
        pass

tesseract = RAIZ / "desktop" / "tesseract"
if tesseract.is_dir():
    datas.append((str(tesseract), "tesseract"))

analise = Analysis(
    [str(RAIZ / "desktop" / "launcher.py")],
    pathex=[str(RAIZ)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "PyQt5", "PySide2", "torch", "tensorflow"],
    noarchive=False,
)

pyz = PYZ(analise.pure)

exe = EXE(
    pyz,
    analise.scripts,
    [],
    exclude_binaries=True,
    name="ExtratorApoena",
    console=True,          # a janela do console mostra o endereço e mantém o servidor vivo
    icon=None,
    upx=False,
)

col = COLLECT(
    exe,
    analise.binaries,
    analise.datas,
    strip=False,
    upx=False,
    name="ExtratorApoena",
)
