# -*- coding: utf-8 -*-
"""
Extrator de Relatórios - Apoena
Versão: Auditoria Avançada + MOTOR OCR NATIVO
"""

import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Novas bibliotecas para o OCR
import pytesseract
from pdf2image import convert_from_bytes

# ==================== CONFIGURAÇÕES DE COLUNAS ====================
COLUNAS_CONFIG = {
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

# ==================== FUNÇÃO DE OCR ====================
def ler_texto_com_ocr(arquivo_pdf):
    """Converte o PDF para imagens e usa o Tesseract OCR para ler o texto com espaços perfeitos"""
    # Recomeça a leitura do ficheiro do zero
    arquivo_pdf.seek(0)
    
    # Converte o PDF para imagens (DPI 300 garante alta qualidade para ler números pequenos)
    imagens = convert_from_bytes(arquivo_pdf.read(), dpi=300)
    texto_completo = ""
    
    # Mostra uma barra de progresso no Streamlit porque o OCR demora um pouco
    barra_progresso = st.progress(0, text="A aplicar OCR nas páginas...")
    total_paginas = len(imagens)
    
    for i, imagem in enumerate(imagens):
        # O psm 6 diz ao OCR para assumir que a imagem é um bloco de texto uniforme (ideal para tabelas)
        texto_pagina = pytesseract.image_to_string(imagem, lang='por', config='--psm 6')
        texto_completo += texto_pagina + "\n"
        barra_progresso.progress((i + 1) / total_paginas, text=f"OCR concluído na página {i+1} de {total_paginas}")
        
    barra_progresso.empty()
    return texto_completo

# ==================== MOTOR DE EXTRAÇÃO ====================
def extrair_linhas_fiscal(arquivo_pdf, usar_ocr):
    dados_locais, linhas_rejeitadas = [], []
    estatisticas = {"Arquivo": arquivo_pdf.name, "Total Fiscais Encontradas": 0, "Sucesso": 0, "Falhas": 0}
    
    # DECISÃO: Como ler o texto?
    if usar_ocr:
        # Usa a nossa nova função de OCR
        texto = ler_texto_com_ocr(arquivo_pdf)
        linhas_do_texto = texto.split('\n')
    else:
        # Usa a extração normal (rápida, mas falha se o PDF estiver aglutinado)
        linhas_do_texto = []
        with pdfplumber.open(arquivo_pdf) as pdf:
            for pagina in pdf.pages:
                txt = pagina.extract_text(layout=True)
                if txt: linhas_do_texto.extend(txt.split('\n'))

    # PROCESSAMENTO DAS LINHAS
    for linha in linhas_do_texto:
        linha = linha.strip()
        if not linha: continue
        
        if padrao_linha_fiscal_generica.match(linha):
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

# ==================== FORMATAÇÃO EXCEL ====================
def gerar_excel_formatado_em_memoria(df, df_erros, modo_abas):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    header_font, header_fill, data_font, alt_fill, border = Font(name="Arial", size=10, bold=True, color="FFFFFF"), PatternFill(start_color="002060", end_color="002060", fill_type="solid"), Font(name="Arial", size=10), PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"), Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    
    def formatar_aba(ws, df_aba, nome):
        ws.title = re.sub(r'[\\/*?:\[\]]', '', nome)[:31]
        ws.append(list(df_aba.columns))
        for cell in ws[1]: cell.font, cell.fill, cell.alignment, cell.border = header_font, header_fill, Alignment(horizontal="center", vertical="center"), border
        for r_idx, row in enumerate(df_aba.itertuples(index=False), 2):
            ws.append(list(row))
            for c_idx, cell in enumerate(ws[r_idx], 1):
                cell.font, cell.alignment, cell.border, cell.fill = data_font, Alignment(vertical="center"), border, (alt_fill if r_idx % 2 == 0 else None)
                if list(df_aba.columns)[c_idx-1] in ["Valor NF", "BC ICMS", "ICMS", "% Interna", "ICMS Origem", "VR DIFAL"] and cell.value is not None:
                    cell.number_format, cell.alignment = "#,##0.00", Alignment(horizontal="right", vertical="center")
        for i, w in enumerate([20, 10, 12, 10, 48, 35, 5, 10, 45, 7, 14, 14, 14, 12, 14, 14, 15], 1): ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes, ws.auto_filter.ref = "A2", ws.dimensions

    if modo_abas == "unica": formatar_aba(wb.create_sheet("Extração Completa"), df, "Extração Completa")
    else:
        for arquivo in df['Arquivo'].unique(): formatar_aba(wb.create_sheet(), df[df['Arquivo'] == arquivo], arquivo.replace(".pdf", ""))
            
    if not df_erros.empty:
        ws_err = wb.create_sheet("⚠️ Linhas Rejeitadas")
        ws_err.append(["Arquivo de Origem", "Linha Completa Não Processada"])
        for cell in ws_err[1]: cell.font, cell.fill, cell.alignment, cell.border = header_font, PatternFill(start_color="C00000", end_color="C00000", fill_type="solid"), Alignment(horizontal="center", vertical="center"), border
        for row in df_erros.itertuples(index=False): ws_err.append(list(row))
        ws_err.column_dimensions['A'].width, ws_err.column_dimensions['B'].width, ws_err.freeze_panes = 30, 150, "A2"
    
    buffer = io.BytesIO(); wb.save(buffer); buffer.seek(0)
    return buffer

# ==================== INTERFACE STREAMLIT ====================
st.set_page_config(page_title="Extrator de Relatórios - Apoena", page_icon="📊", layout="wide")
st.title("📊 Extrator de Relatórios - Apoena")

modo_abas = st.radio("1. Organização das abas:", options=["unica", "separadas"], format_func=lambda x: "📑 Todos os PDFs numa única Aba" if x == "unica" else "📁 Cada PDF numa Aba separada")

# NOVO BOTÃO DE OCR
usar_ocr = st.checkbox("🔍 Forçar Leitura por OCR (Marque apenas se o PDF tiver problemas. Pode demorar alguns minutos.)")

arquivos_selecionados = st.file_uploader("2. Selecione os ficheiros PDF", type=['pdf'], accept_multiple_files=True)

if st.button("Extrair Dados e Gerar Excel", type="primary"):
    if not arquivos_selecionados: st.warning("Selecione um PDF.")
    else:
        dados_finais, dados_rejeitados, todas_estatisticas = [], [], []
        with st.spinner("Processando (Se o OCR estiver ativo, aguarde...)..."):
            for arquivo_pdf in arquivos_selecionados:
                try:
                    res, rej, stats = extrair_linhas_fiscal(arquivo_pdf, usar_ocr)
                    dados_finais.extend(res); dados_rejeitados.extend(rej); todas_estatisticas.append(stats)
                except Exception as e: st.error(f"Erro em {arquivo_pdf.name}: {e}"); st.stop()

        st.markdown("### 🔍 Auditoria")
        for s in todas_estatisticas:
            if s["Falhas"] == 0: st.success(f"✅ **{s['Arquivo']}**: Todas as {s['Total Fiscais Encontradas']} linhas extraídas!")
            else: st.warning(f"⚠️ **{s['Arquivo']}**: {s['Sucesso']} sucesso, **{s['Falhas']} falhas**.")

        if dados_finais or dados_rejeitados:
            df = pd.DataFrame(dados_finais, columns=COLUNAS_CONFIG["fiscal"]) if dados_finais else pd.DataFrame(columns=COLUNAS_CONFIG["fiscal"])
            df_erros = pd.DataFrame(dados_rejeitados) if dados_rejeitados else pd.DataFrame()
            st.download_button("📥 Descarregar Excel Formatado", data=gerar_excel_formatado_em_memoria(df, df_erros, modo_abas), file_name=f"Extracao_Fiscal.xlsx")
