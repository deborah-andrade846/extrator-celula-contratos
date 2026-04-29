# -*- coding: utf-8 -*-
"""
Extrator de Relatórios - Apoena
Atualizado: Opção de Abas, Excel Formatado e Padrão Misto Integrado
"""

import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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

# ==================== FUNÇÕES AUXILIARES ====================
def converter_para_numero(valor_str):
    if not valor_str: return None
    try:
        return float(valor_str.replace('.', '').replace(',', '.'))
    except:
        return valor_str

# ==================== EXPRESSÕES REGULARES (REGEX) ====================
# Padrão Original 89701 (Tem Fornecedor, sem NCM, 3 valores finais)
padrao_89701 = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(\d{44})\s+(.*?)\s+([A-Z]{2})\s+(.*?)\s+(\d{4})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$')

# Padrão Original 92284 (Sem Fornecedor, com NCM, 5 valores finais)
padrao_92284 = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(\d{44})\s+([A-Z]{2})\s+(\d{8})\s+(\d{4})\s+(.*?)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*(.*?)\s*$')

# NOVO: Padrão Misto (Tem Fornecedor, tem NCM, 6 valores finais - Resolve os erros identificados)
padrao_misto = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(\d{44})\s+(.*?)\s+([A-Z]{2})\s+(\d{8})\s+(.*?)\s+(\d{4})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$')

# Padrão genérico apenas para detetar se a linha é ou não é uma linha de Nota Fiscal
padrao_linha_fiscal_generica = re.compile(r'^\d{2}/\d{2}/\d{4}\s+\d+\s+\d{44}')


# ==================== MOTOR DE EXTRAÇÃO ====================
def extrair_linhas_fiscal(arquivo_pdf):
    dados_locais = []
    estatisticas = {
        "Arquivo": arquivo_pdf.name,
        "Total Fiscais Encontradas": 0,
        "Sucesso": 0,
        "Falhas": 0,
        "Linhas com Erro": []
    }
    
    with pdfplumber.open(arquivo_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text(layout=True)
            if not texto: continue
            
            for linha in texto.split('\n'):
                linha = linha.strip()
                if not linha: continue
                
                if padrao_linha_fiscal_generica.match(linha):
                    estatisticas["Total Fiscais Encontradas"] += 1
                    
                    m1 = padrao_89701.match(linha)
                    m2 = padrao_92284.match(linha)
                    m3 = padrao_misto.match(linha)
                    
                    if m3: # 1º Teste: Formato Misto (Completo)
                        dados_locais.append({
                            "Arquivo": arquivo_pdf.name, "Tipo NAI": "Misto",
                            "Data": m3.group(1), "Nº NF": m3.group(2), "Chave da NF-e": m3.group(3),
                            "Fornecedor": m3.group(4).strip(), "UF": m3.group(5), "NCM": m3.group(6),
                            "Descrição": m3.group(7).strip(), "CFOP": m3.group(8),
                            "Valor NF": converter_para_numero(m3.group(9)), 
                            "BC ICMS": converter_para_numero(m3.group(10)),
                            "ICMS": converter_para_numero(m3.group(11)), 
                            "% Interna": converter_para_numero(m3.group(12)),
                            "ICMS Origem": converter_para_numero(m3.group(13)), 
                            "VR DIFAL": converter_para_numero(m3.group(14)), "OBS": ""
                        })
                        estatisticas["Sucesso"] += 1
                        
                    elif m1: # 2º Teste: Formato 89701
                        dados_locais.append({
                            "Arquivo": arquivo_pdf.name, "Tipo NAI": "89701",
                            "Data": m1.group(1), "Nº NF": m1.group(2), "Chave da NF-e": m1.group(3),
                            "Fornecedor": m1.group(4).strip(), "UF": m1.group(5), "NCM": "",
                            "Descrição": m1.group(6).strip(), "CFOP": m1.group(7),
                            "Valor NF": converter_para_numero(m1.group(8)), "BC ICMS": None,
                            "ICMS": None, "% Interna": None,
                            "ICMS Origem": converter_para_numero(m1.group(9)),
                            "VR DIFAL": converter_para_numero(m1.group(10)), "OBS": ""
                        })
                        estatisticas["Sucesso"] += 1
                        
                    elif m2: # 3º Teste: Formato 92284
                        dados_locais.append({
                            "Arquivo": arquivo_pdf.name, "Tipo NAI": "92284",
                            "Data": m2.group(1), "Nº NF": m2.group(2), "Chave da NF-e": m2.group(3),
                            "Fornecedor": "", "UF": m2.group(4), "NCM": m2.group(5),
                            "Descrição": m2.group(7).strip(), "CFOP": m2.group(6),
                            "Valor NF": converter_para_numero(m2.group(8)),
                            "BC ICMS": converter_para_numero(m2.group(9)),
                            "ICMS": converter_para_numero(m2.group(10)),
                            "% Interna": converter_para_numero(m2.group(11)),
                            "ICMS Origem": None, "VR DIFAL": converter_para_numero(m2.group(12)),
                            "OBS": m2.group(13).strip()
                        })
                        estatisticas["Sucesso"] += 1
                        
                    else: # Se não for nenhum dos 3, regista como falha para auditar
                        estatisticas["Falhas"] += 1
                        estatisticas["Linhas com Erro"].append(linha)
                        
    return dados_locais, estatisticas

# ==================== FORMATAÇÃO EXCEL VISUAL ====================
def gerar_excel_formatado_em_memoria(df, tipo, modo_abas):
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove a aba vazia inicial

    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    data_font = Font(name="Arial", size=10)
    alt_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    colunas_moeda = ["Valor NF", "BC ICMS", "ICMS", "% Interna", "ICMS Origem", "VR DIFAL", "Total", "Valor"]

    def preencher_e_formatar_aba(ws, df_aba, nome_aba):
        nome_seguro = re.sub(r'[\\/*?:\[\]]', '', nome_aba)[:31]
        ws.title = nome_seguro

        headers = list(df_aba.columns)
        ws.append(headers)
        for col_idx, cell in enumerate(ws[1], 1):
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
        ws.row_dimensions[1].height = 30

        for row_idx, row_data in enumerate(df_aba.itertuples(index=False), 2):
            ws.append(list(row_data))
            fill = alt_fill if row_idx % 2 == 0 else None
            
            for col_idx, cell in enumerate(ws[row_idx], 1):
                cell.font = data_font
                cell.alignment = Alignment(vertical="center")
                cell.border = border
                if fill: cell.fill = fill
                
                nome_coluna = headers[col_idx - 1]
                if nome_coluna in colunas_moeda and cell.value is not None:
                    try:
                        cell.number_format = "#,##0.00"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    except: pass

        if tipo == "fiscal":
            col_widths = [20, 10, 12, 10, 48, 35, 5, 10, 45, 7, 14, 14, 14, 12, 14, 14, 15]
        else:
            col_widths = [20, 15, 30, 10, 10, 15]
            
        for i, w in enumerate(col_widths, 1):
            if i <= len(headers):
                ws.column_dimensions[get_column_letter(i)].width = w

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    if modo_abas == "unica":
        ws_unica = wb.create_sheet("Extração Completa")
        preencher_e_formatar_aba(ws_unica, df, "Extração Completa")
    else:
        arquivos_unicos = df['Arquivo'].unique()
        for arquivo in arquivos_unicos:
            ws_nova = wb.create_sheet()
            df_filtrado = df[df['Arquivo'] == arquivo]
            titulo_aba = arquivo.replace(".pdf", "").replace(".PDF", "")
            preencher_e_formatar_aba(ws_nova, df_filtrado, titulo_aba)

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

st.markdown("### 2. Como organizar o Excel?")
modo_abas = st.radio(
    "Organização das abas:",
    options=["unica", "separadas"],
    format_func=lambda x: "📑 Todos os PDFs numa única Aba (Planilha)" if x == "unica" else "📁 Cada PDF numa Aba (Planilha) separada",
    label_visibility="collapsed"
)

st.markdown("### 3. Ficheiros PDF")
arquivos_selecionados = st.file_uploader("Arraste os ficheiros PDF aqui", type=['pdf'], accept_multiple_files=True)

if st.button("Extrair Dados e Gerar Excel", type="primary"):
    if not arquivos_selecionados:
        st.warning("Selecione pelo menos um ficheiro PDF.")
    else:
        dados_finais = []
        todas_estatisticas = []
        
        with st.spinner("A processar e a analisar ficheiros..."):
            for arquivo_pdf in arquivos_selecionados:
                try:
                    if tipo_selecionado == "fiscal":
                        linhas_extraidas, stats = extrair_linhas_fiscal(arquivo_pdf)
                        dados_finais.extend(linhas_extraidas)
                        todas_estatisticas.append(stats)
                    else:
                        st.warning("Por favor, implemente as funções de Hotel/Exames se necessitar delas.")

                except Exception as e:
                    st.error(f"Erro ao processar {arquivo_pdf.name}:\n{str(e)}")
                    st.stop()

        if tipo_selecionado == "fiscal":
            st.markdown("### 🔍 Auditoria da Extração")
            for stats in todas_estatisticas:
                if stats["Falhas"] == 0:
                    st.success(f"✅ **{stats['Arquivo']}**: Todas as {stats['Total Fiscais Encontradas']} linhas extraídas com sucesso!")
                else:
                    st.warning(f"⚠️ **{stats['Arquivo']}**: Extraídas {stats['Sucesso']} linhas, mas **{stats['Falhas']} falharam**.")
                    with st.expander("Ver as linhas que NÃO foram extraídas (Ajuste manual necessário):"):
                        for erro in stats["Linhas com Erro"]:
                            st.code(erro)

        if dados_finais:
            df = pd.DataFrame(dados_finais, columns=COLUNAS_CONFIG[tipo_selecionado])
            
            buffer_excel = gerar_excel_formatado_em_memoria(df, tipo_selecionado, modo_abas)
            
            st.download_button(
                label="📥 Descarregar Excel Formatado",
                data=buffer_excel,
                file_name=f"Extracao_{tipo_selecionado}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Não foram extraídos dados válidos.")
