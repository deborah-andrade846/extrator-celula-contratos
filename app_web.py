# -*- coding: utf-8 -*-
"""
Extrator de Relatórios - Apoena
Versão Final: Fiscal Robusto (Abas, Auditoria, OCR) + Relatórios Originais Expandidos
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
from pdf2image import convert_from_bytes

# ==================== CONFIGURAÇÕES DE COLUNAS ====================
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

# ==================== FUNÇÕES AUXILIARES ====================
def converter_para_numero(valor_str):
    if not valor_str: return None
    try: return float(valor_str.replace('.', '').replace(',', '.'))
    except: return valor_str

def limpar_linha_hotel(linha, nome_hospede):
    padrao_data = r'^(\d{2}/\d{2}/\d{2})'
    match = re.search(padrao_data, linha.strip())
    if match:
        data = match.group(1)
        resto = linha[linha.find(data) + len(data):].strip()
        partes = resto.split()
        if len(partes) >= 7 and "," in partes[-1] and "," in partes[-6]:
            total = partes[-1]
            unidade = partes[-5]
            qtde = partes[-6]
            info = " ".join(partes[1:-6]).replace("|", "-").strip()
            info = info.split(" - Comanda")[0].strip()
            return {
                "Arquivo": nome_hospede,
                "Data": data,
                "Informação adicional": info,
                "Qtde": qtde,
                "Unidade": unidade,
                "Total": total
            }
    return None

# ==================== REGEX FISCAL ====================
padrao_89701 = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(\d{44})\s+(.*?)\s+([A-Z]{2})\s+(.*?)\s+(\d{4})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$')
padrao_92284 = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(\d{44})\s+([A-Z]{2})\s+(\d{8})\s+(\d{4})\s+(.*?)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*(.*?)\s*$')
padrao_misto = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(\d{44})\s+(.*?)\s+([A-Z]{2})\s+(\d{8})\s+(.*?)\s+(\d{4})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$')
padrao_linha_fiscal_generica = re.compile(r'^\d{2}/\d{2}/\d{4}\s+\d+\s+\d{44}')

# ==================== FUNÇÃO DE OCR ====================
def ler_texto_com_ocr(arquivo_pdf):
    arquivo_pdf.seek(0)
    imagens = convert_from_bytes(arquivo_pdf.read(), dpi=300)
    texto_completo = ""
    barra_progresso = st.progress(0, text="A aplicar OCR nas páginas...")
    total_paginas = len(imagens)
    for i, imagem in enumerate(imagens):
        texto_pagina = pytesseract.image_to_string(imagem, lang='por', config='--psm 6')
        texto_completo += texto_pagina + "\n"
        barra_progresso.progress((i + 1) / total_paginas, text=f"OCR concluído na página {i+1} de {total_paginas}")
    barra_progresso.empty()
    return texto_completo

# ==================== MOTOR DE EXTRAÇÃO FISCAL ====================
def extrair_linhas_fiscal(arquivo_pdf, usar_ocr):
    dados_locais, linhas_rejeitadas = [], []
    estatisticas = {"Arquivo": arquivo_pdf.name, "Total Fiscais Encontradas": 0, "Sucesso": 0, "Falhas": 0}
    
    if usar_ocr:
        texto = ler_texto_com_ocr(arquivo_pdf)
        linhas_do_texto = texto.split('\n')
    else:
        linhas_do_texto = []
        with pdfplumber.open(arquivo_pdf) as pdf:
            for pagina in pdf.pages:
                txt = pagina.extract_text(layout=True)
                if txt: linhas_do_texto.extend(txt.split('\n'))

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
def gerar_excel_formatado_em_memoria(df, df_erros, tipo_relatorio, modo_abas):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    data_font = Font(name="Arial", size=10)
    alt_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    
    colunas_moeda = ["Valor NF", "BC ICMS", "ICMS", "% Interna", "ICMS Origem", "VR DIFAL", "Total", "Valor"]

    def formatar_aba(ws, df_aba, nome):
        ws.title = re.sub(r'[\\/*?:\[\]]', '', nome)[:31]
        ws.append(list(df_aba.columns))
        for cell in ws[1]: 
            cell.font, cell.fill, cell.alignment, cell.border = header_font, header_fill, Alignment(horizontal="center", vertical="center"), border
        
        for r_idx, row in enumerate(df_aba.itertuples(index=False), 2):
            ws.append(list(row))
            for c_idx, cell in enumerate(ws[r_idx], 1):
                cell.font, cell.alignment, cell.border = data_font, Alignment(vertical="center"), border
                if r_idx % 2 == 0:
                    cell.fill = alt_fill
                nome_coluna = list(df_aba.columns)[c_idx-1]
                if nome_coluna in colunas_moeda and cell.value is not None:
                    try:
                        cell.number_format = "#,##0.00"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    except: pass
        
        widths = [20, 10, 12, 10, 48, 35, 5, 10, 45, 7, 14, 14, 14, 12, 14, 14, 15] if tipo_relatorio == "fiscal" else [20, 15, 30, 10, 10, 15]
        for i, w in enumerate(widths, 1): 
            if i <= len(list(df_aba.columns)):
                ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    if modo_abas == "unica": 
        formatar_aba(wb.create_sheet("Extração Completa"), df, "Extração Completa")
    else:
        for arquivo in df['Arquivo'].unique(): 
            formatar_aba(wb.create_sheet(), df[df['Arquivo'] == arquivo], arquivo.replace(".pdf", ""))
            
    if not df_erros.empty and tipo_relatorio == "fiscal":
        ws_err = wb.create_sheet("⚠️ Linhas Rejeitadas")
        ws_err.append(["Arquivo de Origem", "Linha Completa Não Processada"])
        for cell in ws_err[1]: 
            cell.font = header_font
            cell.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
        for row in df_erros.itertuples(index=False): 
            ws_err.append(list(row))
        ws_err.column_dimensions['A'].width = 30
        ws_err.column_dimensions['B'].width = 150
        ws_err.freeze_panes = "A2"
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ==================== INTERFACE STREAMLIT ====================
st.set_page_config(page_title="Extrator de Relatórios - Apoena", page_icon="📊", layout="wide")
st.title("📊 Extrator de Relatórios - Apoena")

st.markdown("### 1. O que deseja extrair?")
tipo_selecionado = st.radio(
    "Escolha o tipo de relatório:",
    options=["fiscal", "hotel", "exames", "refeicoes"],
    format_func=lambda x: {
        "hotel": "Diárias e Consumo (Plaza Hotel)",
        "exames": "Exames Ocupacionais (Biomed)",
        "refeicoes": "Mapa de Refeições",
        "fiscal": "Notas Fiscais (Autuação SEFAZ)"
    }[x],
    label_visibility="collapsed"
)

st.markdown("### 2. Configurações de Extração")
modo_abas = st.radio(
    "Organização das abas:", 
    options=["unica", "separadas"], 
    format_func=lambda x: "📑 Todos os PDFs numa única Aba" if x == "unica" else "📁 Cada PDF numa Aba separada"
)

usar_ocr = False
if tipo_selecionado == "fiscal":
    usar_ocr = st.checkbox("🔍 Forçar Leitura por OCR (Marque apenas se o PDF tiver problemas. Pode demorar alguns minutos.)")

st.markdown("### 3. Ficheiros PDF")
arquivos_selecionados = st.file_uploader("Arraste os ficheiros PDF aqui", type=['pdf'], accept_multiple_files=True)

if st.button("Extrair Dados e Gerar Excel", type="primary"):
    if not arquivos_selecionados: 
        st.warning("Por favor, selecione pelo menos um ficheiro PDF.")
    else:
        dados_finais = []
        dados_rejeitados = []
        todas_estatisticas = []
        
        with st.spinner("Processando (Se o OCR estiver ativo, por favor aguarde)..."):
            for arquivo_pdf in arquivos_selecionados:
                nome_arquivo = arquivo_pdf.name
                try:
                    if tipo_selecionado == "fiscal":
                        res, rej, stats = extrair_linhas_fiscal(arquivo_pdf, usar_ocr)
                        dados_finais.extend(res)
                        dados_rejeitados.extend(rej)
                        todas_estatisticas.append(stats)
                    
                    else:
                        # LÓGICA EXPANDIDA PARA OUTROS RELATÓRIOS
                        with pdfplumber.open(arquivo_pdf) as pdf:
                            texto_completo = "\n".join([p.extract_text() or "" for p in pdf.pages])
                            linhas = texto_completo.split('\n')
                            
                            # --- Lógica do Hotel ---
                            if tipo_selecionado == "hotel":
                                nome_hospede = "NÃO_IDENTIFICADO"
                                for linha in linhas:
                                    if "Hóspede principal:" in linha: 
                                        try:
                                            nome_cru = linha.split("Hóspede principal:")[1].split("|")[0].strip()
                                            nome_hospede = nome_cru.split()[0]
                                        except: pass
                                        continue
                                    
                                    if any(x in linha for x in ["PLAZA HOTEL", "Apartamento:", "Fechado", "Pagamentos", "Tarifário:"]): 
                                        continue
                                    
                                    item = limpar_linha_hotel(linha, nome_hospede)
                                    if item: 
                                        dados_finais.append(item)
                            
                            # --- Lógica de Exames ---
                            elif tipo_selecionado == "exames":
                                for linha in linhas:
                                    if "R$" in linha:
                                        partes = linha.split("R$")
                                        if len(partes) >= 2:
                                            nome_exame = partes[0].replace('"', '').replace(',', '').strip()
                                            valor_exame = "R$ " + partes[1].replace('"', '').replace(',', '').strip()
                                            if nome_exame:
                                                dados_finais.append({
                                                    "Arquivo": nome_arquivo, 
                                                    "Exame": nome_exame, 
                                                    "Valor": valor_exame
                                                })
                            
                            # --- Lógica de Refeições ---
                            elif tipo_selecionado == "refeicoes":
                                data_ref = "DATA_NAO_ENCONTRADA"
                                total_ref = "TOTAL_NAO_ENCONTRADO"
                                for linha in linhas:
                                    if "Período:" in linha: 
                                        match = re.search(r'\d{2}/\d{2}/\d{4}', linha)
                                        if match: 
                                            data_ref = match.group(0)
                                    if "Total Geral" in linha: 
                                        valor_limpo = linha.replace("Total Geral", "").replace("|", "").strip()
                                        if valor_limpo:
                                            total_ref = valor_limpo
                                            
                                if data_ref != "DATA_NAO_ENCONTRADA" or total_ref != "TOTAL_NAO_ENCONTRADO":
                                    dados_finais.append({
                                        "Arquivo": nome_arquivo, 
                                        "Data": data_ref, 
                                        "Total": total_ref
                                    })
                                    
                except Exception as e: 
                    st.error(f"Erro ao processar {nome_arquivo}: {e}")
                    st.stop()

        # Apresentação de Resultados
        if tipo_selecionado == "fiscal":
            st.markdown("### 🔍 Auditoria")
            for s in todas_estatisticas:
                if s["Falhas"] == 0: 
                    st.success(f"✅ **{s['Arquivo']}**: Todas as {s['Total Fiscais Encontradas']} linhas extraídas com sucesso!")
                else: 
                    st.warning(f"⚠️ **{s['Arquivo']}**: {s['Sucesso']} sucesso, **{s['Falhas']} falhas**.")

        if dados_finais or dados_rejeitados:
            df = pd.DataFrame(dados_finais, columns=COLUNAS_CONFIG[tipo_selecionado]) if dados_finais else pd.DataFrame(columns=COLUNAS_CONFIG[tipo_selecionado])
            df_erros = pd.DataFrame(dados_rejeitados) if dados_rejeitados else pd.DataFrame()
            
            st.download_button(
                "📥 Descarregar Excel Formatado", 
                data=gerar_excel_formatado_em_memoria(df, df_erros, tipo_selecionado, modo_abas), 
                file_name=f"Extracao_{tipo_selecionado}.xlsx"
            )
        else:
            st.info("Nenhuma linha válida encontrada nestes documentos para o tipo selecionado.")
