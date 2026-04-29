# -*- coding: utf-8 -*-
"""
Extrator de Relatórios - Apoena
Atualizado: Retorno da Formatação Visual (OpenPyXL) e Auditoria de Linhas
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

# Ajuste fino nas Regex para ignorar espaços em branco no final da linha (\s*$)
padrao_89701 = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(\d{44})\s+(.*?)\s+([A-Z]{2})\s+(.*?)\s+(\d{4})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$')
padrao_92284 = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(\d{44})\s+([A-Z]{2})\s+(\d{8})\s+(\d{4})\s+(.*?)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*(.*?)\s*$')
padrao_linha_fiscal_generica = re.compile(r'^\d{2}/\d{2}/\d{4}\s+\d+\s+\d{44}')

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
                
                # Se a linha tem cara de ser uma linha fiscal, registamos no auditor
                if padrao_linha_fiscal_generica.match(linha):
                    estatisticas["Total Fiscais Encontradas"] += 1
                    
                    m1 = padrao_89701.match(linha)
                    m2 = padrao_92284.match(linha)
                    
                    if m1:
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
                    elif m2:
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
                    else:
                        estatisticas["Falhas"] += 1
                        estatisticas["Linhas com Erro"].append(linha)
                        
    return dados_locais, estatisticas

# ==================== FORMATAÇÃO EXCEL VISUAL ====================
def gerar_excel_formatado_em_memoria(df, tipo):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Extração"

    # Estilos exatos do seu código original
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    data_font = Font(name="Arial", size=10)
    alt_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Escrever Cabeçalho
    headers = list(df.columns)
    ws.append(headers)
    for col_idx, cell in enumerate(ws[1], 1):
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    ws.row_dimensions[1].height = 30

    # Determinar quais colunas levam formatação de moeda
    colunas_moeda = ["Valor NF", "BC ICMS", "ICMS", "% Interna", "ICMS Origem", "VR DIFAL", "Total", "Valor"]

    # Escrever Dados
    for row_idx, row_data in enumerate(df.itertuples(index=False), 2):
        ws.append(list(row_data))
        fill = alt_fill if row_idx % 2 == 0 else None
        
        for col_idx, cell in enumerate(ws[row_idx], 1):
            cell.font = data_font
            cell.alignment = Alignment(vertical="center")
            cell.border = border
            if fill:
                cell.fill = fill
            
            nome_coluna = headers[col_idx - 1]
            if nome_coluna in colunas_moeda and cell.value is not None:
                try:
                    cell.number_format = "#,##0.00"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                except:
                    pass

    # Larguras das colunas
    if tipo == "fiscal":
        col_widths = [20, 10, 12, 10, 48, 35, 5, 10, 45, 7, 14, 14, 14, 12, 14, 14, 15]
    else:
        col_widths = [20, 15, 30, 10, 10, 15]
        
    for i, w in enumerate(col_widths, 1):
        if i <= len(headers):
            ws.column_dimensions[get_column_letter(i)].width = w

    # Painel fixo e filtro
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Guardar em buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ==================== INTERFACE STREAMLIT ====================
st.set_page_config(page_title="Extrator de Relatórios - Apoena", page_icon="📊")
st.title("📊 Extrator de Relatórios - Apoena")

tipo_selecionado = st.radio(
    "Escolha o tipo de relatório:",
    options=["fiscal", "hotel", "exames", "refeicoes"],
    format_func=lambda x: {
        "hotel": "Diárias e Consumo (Plaza Hotel)",
        "exames": "Exames Ocupacionais (Biomed)",
        "refeicoes": "Mapa de Refeições",
        "fiscal": "Notas Fiscais (Autuação SEFAZ)"
    }[x]
)

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
                        # (Omitido aqui por brevidade, pode colar as lógicas antigas se precisar do hotel/exames)

                except Exception as e:
                    st.error(f"Erro ao processar {arquivo_pdf.name}:\n{str(e)}")
                    st.stop()

        # Apresentar o Resumo de Auditoria para dar Certeza
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
            
            # Chama a função de formatação para criar o Excel bonito
            buffer_excel = gerar_excel_formatado_em_memoria(df, tipo_selecionado)
            
            st.download_button(
                label="📥 Descarregar Excel Formatado e Colorido",
                data=buffer_excel,
                file_name=f"Extracao_{tipo_selecionado}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Não foram extraídos dados válidos.")
