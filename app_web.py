# -*- coding: utf-8 -*-
"""
Extrator de Relatórios - Apoena
Versão Final: Visão Computacional + Filtro de Ruído + Separador Inteligente
"""

import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pytesseract
import cv2
import numpy as np

# ==================== CONFIGURAÇÕES ====================
COLUNAS_CONFIG = {
    "hotel": ["Arquivo", "Data", "Informação adicional", "Qtde", "Unidade", "Total"],
    "exames": ["Arquivo", "Exame", "Valor"],
    "refeicoes": ["Arquivo", "Data", "Total"],
    "fiscal": [
        "Arquivo", "Tipo NAI", "Data", "Nº NF", "Chave da NF-e", "Fornecedor", 
        "UF", "NCM", "Descrição", "CFOP", "Valor NF", "BC ICMS", "ICMS", 
        "% Interna", "ICMS Origem", "VR DIFAL", "OBS"
    ]
}

def converter_para_numero(valor_str):
    if not valor_str: return None
    try: return float(valor_str.replace('.', '').replace(',', '.'))
    except: return valor_str

# ==================== REGEX FISCAL ====================
padrao_89701 = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(\d{44})\s+(.*?)\s+([A-Z]{2})\s+(.*?)\s+(\d{4})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$')
padrao_92284 = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(\d{44})\s+([A-Z]{2})\s+(\d{8})\s+(\d{4})\s+(.*?)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*(.*?)\s*$')
padrao_misto = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(\d{44})\s+(.*?)\s+([A-Z]{2})\s+(\d{8})\s+(.*?)\s+(\d{4})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$')
padrao_linha_fiscal_generica = re.compile(r'^\d{2}/\d{2}/\d{4}\s+\d+\s+\d{44}')

# ==================== OCR COM VISÃO COMPUTACIONAL ====================
def ler_texto_com_ocr(arquivo_pdf):
    texto_completo = ""
    barra_progresso = st.progress(0, text="A aplicar Visão Computacional e OCR...")
    with pdfplumber.open(arquivo_pdf) as pdf:
        total_paginas = len(pdf.pages)
        for i, pagina in enumerate(pdf.pages):
            # Alta resolução (300 DPI) para captar vírgulas
            img_pil = pagina.to_image(resolution=300).original
            img_cv = np.array(img_pil)
            if len(img_cv.shape) == 3: img_cv = cv2.cvtColor(img_cv, cv2.COLOR_RGB2BGR)
            
            # Filtros de Limpeza OpenCV
            img_cinza = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
            img_suave = cv2.GaussianBlur(img_cinza, (3, 3), 0)
            _, img_bin = cv2.threshold(img_suave, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
            
            texto_completo += pytesseract.image_to_string(img_bin, lang='por', config='--psm 6 --oem 3') + "\n"
            del img_pil, img_cv, img_cinza, img_suave, img_bin
            barra_progresso.progress((i + 1) / total_paginas, text=f"OCR Otimizado: Página {i+1} de {total_paginas}")
    barra_progresso.empty()
    return texto_completo

# ==================== MOTOR DE EXTRAÇÃO ====================
def extrair_linhas_fiscal(arquivo_pdf, usar_ocr):
    dados_locais, linhas_rejeitadas = [], []
    estatisticas = {"Arquivo": arquivo_pdf.name, "Total Fiscais Encontradas": 0, "Sucesso": 0, "Falhas": 0}
    
    if usar_ocr:
        linhas_brutas = ler_texto_com_ocr(arquivo_pdf).split('\n')
    else:
        linhas_brutas = []
        with pdfplumber.open(arquivo_pdf) as pdf:
            for pagina in pdf.pages:
                txt = pagina.extract_text(layout=True)
                if txt: linhas_brutas.extend(txt.split('\n'))

    # 🚀 Separador Inteligente e Filtro de Ruído
    linhas_limpas = []
    for linha in linhas_brutas:
        # Separa datas coladas (ex: 67,4418/02/2022)
        linha = re.sub(r'(?<!^)(?=\d{2}/\d{2}/\d{4}\s+\d+\s+\d{44})', '\n', linha)
        for sub_linha in linha.split('\n'):
            # Apaga barras verticais, colchetes e lixo de tabela
            sub_linha = re.sub(r'[|—\[\]]', '', sub_linha)
            # Corrige estados com ponto (ex: SP. -> SP)
            sub_linha = re.sub(r'\b([A-Z]{2})\.', r'\1', sub_linha)
            linhas_limpas.append(sub_linha.strip())

    for linha in linhas_limpas:
        if not linha or not padrao_linha_fiscal_generica.match(linha): continue
        estatisticas["Total Fiscais Encontradas"] += 1
        m1, m2, m3 = padrao_89701.match(linha), padrao_92284.match(linha), padrao_misto.match(linha)
        
        if m3:
            dados_locais.append({"Arquivo": arquivo_pdf.name, "Tipo NAI": "Misto", "Data": m3.group(1), "Nº NF": m3.group(2), "Chave da NF-e": m3.group(3), "Fornecedor": m3.group(4).strip(), "UF": m3.group(5), "NCM": m3.group(6), "Descrição": m3.group(7).strip(), "CFOP": m3.group(8), "Valor NF": converter_para_numero(m3.group(9)), "BC ICMS": converter_para_numero(m3.group(10)), "ICMS": converter_para_numero(m3.group(11)), "% Interna": converter_para_numero(m3.group(12)), "ICMS Origem": converter_para_numero(m3.group(13)), "VR DIFAL": converter_para_numero(m3.group(14)), "OBS": ""})
            estatisticas["Sucesso"] += 1
        elif m1:
            dados_locais.append({"Arquivo": arquivo_pdf.name, "Tipo NAI": "89701", "Data": m1.group(1), "Nº NF": m1.group(2), "Chave da NF-e": m1.group(3), "Fornecedor": m1.group(4).strip(), "UF": m1.group(5), "NCM": "", "Descrição": m1.group(6).strip(), "CFOP": m1.group(7), "Valor NF": converter_para_numero(m1.group(8)), "BC ICMS": None, "ICMS": None, "% Interna": None, "ICMS Origem": converter_para_numero(m1.group(9)), "VR DIFAL": converter_para_numero(m1.group(10)), "OBS": ""})
            estatisticas["Sucesso"] += 1
        elif m2:
            dados_locais.append({"Arquivo": arquivo_pdf.name, "Tipo NAI": "92284", "Data": m2.group(1), "Nº NF": m2.group(2), "Chave da NF-e": m2.group(3), "Fornecedor": "", "UF": m2.group(4), "NCM": m2.group(5), "Descrição": m2.group(7).strip(), "CFOP": m2.group(6), "Valor NF": converter_para_numero(m2.group(8)), "BC ICMS": converter_para_numero(m2.group(9)), "ICMS": converter_para_numero(m2.group(10)), "% Interna": converter_para_numero(m2.group(11)), "ICMS Origem": None, "VR DIFAL": converter_para_numero(m2.group(12)), "OBS": m2.group(13).strip()})
            estatisticas["Sucesso"] += 1
        else:
            estatisticas["Falhas"] += 1
            linhas_rejeitadas.append({"Arquivo": arquivo_pdf.name, "Linha Completa Não Processada": linha})
            
    return dados_locais, linhas_rejeitadas, estatisticas

# ==================== INTERFACE ====================
def gerar_excel_formatado_em_memoria(df, df_erros, tipo_relatorio, modo_abas):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    header_style = {"font": Font(name="Arial", size=10, bold=True, color="FFFFFF"), "fill": PatternFill(start_color="002060", end_color="002060", fill_type="solid")}
    
    def formatar_aba(ws, df_aba, nome):
        ws.title = re.sub(r'[\\/*?:\[\]]', '', nome)[:31]
        ws.append(list(df_aba.columns))
        for cell in ws[1]: cell.font, cell.fill, cell.alignment = header_style["font"], header_style["fill"], Alignment(horizontal="center")
        for r_idx, row in enumerate(df_aba.itertuples(index=False), 2):
            ws.append(list(row))
            for cell in ws[r_idx]: cell.font = Font(name="Arial", size=10)
        ws.freeze_panes = "A2"

    if modo_abas == "unica" or df.empty: formatar_aba(wb.create_sheet("Extração"), df, "Extração Completa")
    else:
        for arquivo in df['Arquivo'].unique(): formatar_aba(wb.create_sheet(), df[df['Arquivo'] == arquivo], arquivo[:30])
            
    if not df_erros.empty:
        ws_err = wb.create_sheet("⚠️ Linhas Rejeitadas")
        ws_err.append(["Arquivo", "Linha"])
        for cell in ws_err[1]: cell.font, cell.fill = header_style["font"], PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
        for row in df_erros.itertuples(index=False): ws_err.append(list(row))

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

st.set_page_config(page_title="Apoena Extrator", layout="wide")
st.title("📊 Extrator de Relatórios - Apoena")

tipo = st.radio("1. Tipo:", ["fiscal", "hotel", "exames", "refeicoes"])
modo = st.radio("2. Abas:", ["unica", "separadas"])
usar_ocr = st.checkbox("🔍 OCR Alta Precisão (Para PDFs com falhas)") if tipo == "fiscal" else False

files = st.file_uploader("3. PDFs:", type=['pdf'], accept_multiple_files=True)

if st.button("Extrair Dados", type="primary"):
    if files:
        finais, rejeitados, stats_lista = [], [], []
        with st.spinner("A processar..."):
            for f in files:
                if tipo == "fiscal":
                    res, rej, stats = extrair_linhas_fiscal(f, usar_ocr)
                    finais.extend(res); rejeitados.extend(rej); stats_lista.append(stats)
                else:
                    # Lógica simplificada para outros tipos (mantida do original)
                    with pdfplumber.open(f) as pdf:
                        txt = "\n".join([p.extract_text() or "" for p in pdf.pages])
                        if tipo == "refeicoes":
                            dt = re.search(r'\d{2}/\d{2}/\d{4}', txt).group(0) if re.search(r'\d{2}/\d{2}/\d{4}', txt) else "N/D"
                            tt = re.search(r'Total Geral\s*\|?\s*([\d.,]+)', txt).group(1) if re.search(r'Total Geral\s*\|?\s*([\d.,]+)', txt) else "N/D"
                            finais.append({"Arquivo": f.name, "Data": dt, "Total": tt})

        if tipo == "fiscal":
            for s in stats_lista:
                if s["Falhas"] == 0: st.success(f"✅ {s['Arquivo']} OK!")
                else: st.warning(f"⚠️ {s['Arquivo']}: {s['Falhas']} falhas.")

        if finais or rejeitados:
            df = pd.DataFrame(finais, columns=COLUNAS_CONFIG[tipo])
            df_e = pd.DataFrame(rejeitados)
            st.download_button("📥 Descarregar Excel", gerar_excel_formatado_em_memoria(df, df_e, tipo, modo), f"Extracao_{tipo}.xlsx")
