# -*- coding: utf-8 -*-
"""
Extrator de Relatórios - Apoena
Versão Unificada: Fiscal (OCR + NAI) + Hotel + Exames + Refeições
"""

import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import pytesseract
import cv2
import numpy as np
from typing import List, Tuple, Dict, Optional
from dataclasses import dataclass, asdict
import logging
from contextlib import contextmanager
import gc

# ==================== CONFIGURAÇÕES GLOBAIS ====================
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

# ==================== LOGGING ====================
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# ==================== DATACLASS PARA ESTATÍSTICAS ====================
@dataclass
class EstatisticasProcessamento:
    arquivo: str
    total_linhas_encontradas: int = 0
    sucesso: int = 0
    falhas: int = 0

    def taxa_sucesso(self) -> float:
        if self.total_linhas_encontradas == 0:
            return 0.0
        return (self.sucesso / self.total_linhas_encontradas) * 100

# ==================== GERENCIADOR DE MEMÓRIA ====================
@contextmanager
def gerenciar_memoria(limite_mb: int = 500):
    try:
        yield
    finally:
        gc.collect()

# ==================== FUNÇÕES AUXILIARES ====================
def converter_para_numero(valor_str):
    if not valor_str or valor_str == "N/D":
        return None
    try:
        return float(valor_str.replace('.', '').replace(',', '.'))
    except:
        return valor_str

# ==================== OCR COM VISÃO COMPUTACIONAL (FISCAL) ====================
def ler_texto_com_ocr(arquivo_pdf):
    texto_completo = ""
    barra_progresso = st.progress(0, text="A aplicar Visão Computacional e OCR...")
    with pdfplumber.open(arquivo_pdf) as pdf:
        total_paginas = len(pdf.pages)
        for i, pagina in enumerate(pdf.pages):
            img_pil = pagina.to_image(resolution=300).original
            img_cv = np.array(img_pil)
            if len(img_cv.shape) == 3:
                img_cv = cv2.cvtColor(img_cv, cv2.COLOR_RGB2BGR)
            img_cinza = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
            img_suave = cv2.GaussianBlur(img_cinza, (3, 3), 0)
            _, img_bin = cv2.threshold(img_suave, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
            texto_completo += pytesseract.image_to_string(img_bin, lang='por', config='--psm 6 --oem 3') + "\n"
            barra_progresso.progress((i + 1) / total_paginas, text=f"OCR: Página {i+1} de {total_paginas}")
    barra_progresso.empty()
    return texto_completo

# ==================== PARSER FISCAL (NAI 89701, 92284, MISTO) ====================
class FiscalParser:
    PADROES = {
        '89701': re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(\d{44})\s+(.*?)\s+([A-Z]{2})\s+(.*?)\s+(\d{4})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$'),
        '92284': re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(\d{44})\s+([A-Z]{2})\s+(\d{8})\s+(\d{4})\s+(.*?)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*(.*?)\s*$'),
        'MISTO': re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(\d{44})\s+(.*?)\s+([A-Z]{2})\s+(\d{8})\s+(.*?)\s+(\d{4})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$')
    }

    @classmethod
    def parsear_linha(cls, linha: str, arquivo: str) -> Optional[Dict]:
        if not re.match(r'^\d{2}/\d{2}/\d{4}\s+\d+\s+\d{44}', linha):
            return None
        for tipo, padrao in cls.PADROES.items():
            m = padrao.match(linha)
            if m:
                return cls._montar_registro(m, tipo, arquivo)
        return None

    @staticmethod
    def _montar_registro(m, tipo, arquivo):
        if tipo == 'MISTO':
            return {
                "Arquivo": arquivo, "Tipo NAI": "Misto", "Data": m.group(1), "Nº NF": m.group(2),
                "Chave da NF-e": m.group(3), "Fornecedor": m.group(4).strip(), "UF": m.group(5),
                "NCM": m.group(6), "Descrição": m.group(7).strip(), "CFOP": m.group(8),
                "Valor NF": converter_para_numero(m.group(9)), "BC ICMS": converter_para_numero(m.group(10)),
                "ICMS": converter_para_numero(m.group(11)), "% Interna": converter_para_numero(m.group(12)),
                "ICMS Origem": converter_para_numero(m.group(13)), "VR DIFAL": converter_para_numero(m.group(14)), "OBS": ""
            }
        elif tipo == '89701':
            return {
                "Arquivo": arquivo, "Tipo NAI": "89701", "Data": m.group(1), "Nº NF": m.group(2),
                "Chave da NF-e": m.group(3), "Fornecedor": m.group(4).strip(), "UF": m.group(5),
                "NCM": "", "Descrição": m.group(6).strip(), "CFOP": m.group(7),
                "Valor NF": converter_para_numero(m.group(8)), "BC ICMS": None, "ICMS": None,
                "% Interna": None, "ICMS Origem": converter_para_numero(m.group(9)),
                "VR DIFAL": converter_para_numero(m.group(10)), "OBS": ""
            }
        else:  # 92284
            return {
                "Arquivo": arquivo, "Tipo NAI": "92284", "Data": m.group(1), "Nº NF": m.group(2),
                "Chave da NF-e": m.group(3), "Fornecedor": "", "UF": m.group(4),
                "NCM": m.group(5), "Descrição": m.group(7).strip(), "CFOP": m.group(6),
                "Valor NF": converter_para_numero(m.group(8)), "BC ICMS": converter_para_numero(m.group(9)),
                "ICMS": converter_para_numero(m.group(10)), "% Interna": converter_para_numero(m.group(11)),
                "ICMS Origem": None, "VR DIFAL": converter_para_numero(m.group(12)), "OBS": m.group(13).strip()
            }

# ==================== EXTRAÇÃO FISCAL COMPLETA ====================
def extrair_fiscal(arquivo_pdf, usar_ocr: bool):
    dados_locais, rejeitadas = [], []
    stats = EstatisticasProcessamento(arquivo=arquivo_pdf.name)

    # Extrai texto bruto
    if usar_ocr:
        texto = ler_texto_com_ocr(arquivo_pdf)
    else:
        texto = ""
        with pdfplumber.open(arquivo_pdf) as pdf:
            for pagina in pdf.pages:
                txt = pagina.extract_text(layout=True)
                if txt:
                    texto += txt + "\n"

    # Limpeza e separação inteligente
    linhas_brutas = texto.split('\n')
    linhas_limpas = []
    for linha in linhas_brutas:
        linha = re.sub(r'(?<!^)(?=\d{2}/\d{2}/\d{4}\s+\d+\s+\d{44})', '\n', linha)
        for sub in linha.split('\n'):
            sub = re.sub(r'[|—\[\]]', '', sub)
            sub = re.sub(r'\b([A-Z]{2})\.', r'\1', sub)
            linhas_limpas.append(sub.strip())

    for linha in linhas_limpas:
        if not linha or not re.match(r'^\d{2}/\d{2}/\d{4}\s+\d+\s+\d{44}', linha):
            continue
        stats.total_linhas_encontradas += 1
        registro = FiscalParser.parsear_linha(linha, arquivo_pdf.name)
        if registro:
            dados_locais.append(registro)
            stats.sucesso += 1
        else:
            stats.falhas += 1
            rejeitadas.append({"Arquivo": arquivo_pdf.name, "Linha Completa Não Processada": linha})

    return dados_locais, rejeitadas, stats

# ==================== EXTRAÇÃO HOTEL (código antigo funcional) ====================
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

def extrair_hotel(arquivo_pdf):
    dados = []
    nome_hospede = "NÃO_IDENTIFICADO"
    with pdfplumber.open(arquivo_pdf) as pdf:
        texto_completo = "\n".join([pagina.extract_text() or "" for pagina in pdf.pages])
        linhas = texto_completo.split('\n')
        for linha in linhas:
            if "Hóspede principal:" in linha:
                try:
                    nome_cru = linha.split("Hóspede principal:")[1].split("|")[0].strip()
                    nome_hospede = nome_cru.split()[0]
                except:
                    pass
                continue
            # Pula linhas irrelevantes
            if any(palavra in linha for palavra in ["PLAZA HOTEL", "Apartamento:", "Fechado", "Pagamentos", "Tarifário:"]):
                continue
            linha_extraida = limpar_linha_hotel(linha, nome_hospede)
            if linha_extraida:
                dados.append(linha_extraida)
    return dados

# ==================== EXTRAÇÃO EXAMES (código antigo funcional) ====================
def extrair_exames(arquivo_pdf):
    dados = []
    with pdfplumber.open(arquivo_pdf) as pdf:
        texto_completo = "\n".join([pagina.extract_text() or "" for pagina in pdf.pages])
        linhas = texto_completo.split('\n')
        for linha in linhas:
            if "R$" in linha:
                partes = linha.split("R$")
                if len(partes) >= 2:
                    nome_exame = partes[0].replace('"', '').replace(',', '').strip()
                    valor_exame = "R$ " + partes[1].replace('"', '').replace(',', '').strip()
                    if nome_exame:
                        dados.append({
                            "Arquivo": arquivo_pdf.name,
                            "Exame": nome_exame,
                            "Valor": valor_exame
                        })
    return dados

# ==================== EXTRAÇÃO REFEIÇÕES (código antigo funcional) ====================
def extrair_refeicoes(arquivo_pdf):
    dados = []
    data_refeicao = "DATA_NAO_ENCONTRADA"
    total_refeicoes = "TOTAL_NAO_ENCONTRADO"
    with pdfplumber.open(arquivo_pdf) as pdf:
        texto_completo = "\n".join([pagina.extract_text() or "" for pagina in pdf.pages])
        linhas = texto_completo.split('\n')
        for linha in linhas:
            if "Período:" in linha:
                match = re.search(r'\d{2}/\d{2}/\d{4}', linha)
                if match:
                    data_refeicao = match.group(0)
            if "Total Geral" in linha:
                valor_limpo = linha.replace("Total Geral", "").replace("|", "").strip()
                if valor_limpo:
                    total_refeicoes = valor_limpo
        if data_refeicao != "DATA_NAO_ENCONTRADA" or total_refeicoes != "TOTAL_NAO_ENCONTRADO":
            dados.append({
                "Arquivo": arquivo_pdf.name,
                "Data": data_refeicao,
                "Total": total_refeicoes
            })
    return dados

# ==================== EXPORTAÇÃO EXCEL FORMATADO ====================
def gerar_excel_formatado(df_principal, df_erros, tipo_relatorio, modo_abas):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    header_style = {
        "font": Font(name="Arial", size=10, bold=True, color="FFFFFF"),
        "fill": PatternFill(start_color="002060", end_color="002060", fill_type="solid"),
        "alignment": Alignment(horizontal="center")
    }

    def adicionar_aba(ws, df, nome):
        ws.title = re.sub(r'[\\/*?:\[\]]', '', nome)[:31]
        ws.append(list(df.columns))
        for cell in ws[1]:
            cell.font = header_style["font"]
            cell.fill = header_style["fill"]
            cell.alignment = header_style["alignment"]
        for r_idx, row in enumerate(df.itertuples(index=False), 2):
            ws.append(list(row))
            for cell in ws[r_idx]:
                cell.font = Font(name="Arial", size=10)
        ws.freeze_panes = "A2"

    if modo_abas == "unica" or df_principal.empty:
        adicionar_aba(wb.create_sheet("Extração"), df_principal, "Extração Completa")
    else:
        for arquivo in df_principal['Arquivo'].unique():
            adicionar_aba(wb.create_sheet(), df_principal[df_principal['Arquivo'] == arquivo], arquivo[:30])

    if not df_erros.empty:
        ws_err = wb.create_sheet("⚠️ Linhas Rejeitadas")
        ws_err.append(["Arquivo", "Linha"])
        for cell in ws_err[1]:
            cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
        for row in df_erros.itertuples(index=False):
            ws_err.append(list(row))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ==================== INTERFACE STREAMLIT ====================
st.set_page_config(page_title="Apoena Extrator", layout="wide")

# --- Estilos CSS profissionais ---
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #1A2B5E 0%, #F4614D 100%);
        padding: 1rem;
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    .stat-card {
        background: #F5F7FA;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #F4614D;
        margin: 0.5rem 0;
    }
    .success-badge {
        background: #22C55E;
        color: white;
        padding: 0.2rem 0.8rem;
        border-radius: 20px;
        display: inline-block;
    }
</style>
""", unsafe_allow_html=True)

# Cabeçalho
col1, col2, col3 = st.columns([1, 5, 1])
with col2:
    st.markdown('<div class="main-header">', unsafe_allow_html=True)
    st.title("🚀 Extrator Universal de Dados - Apoena")
    st.caption("""
        <span style="color:#F4614D;">●</span> Processamento Inteligente 
        <span style="color:#0D9488;">●</span> OCR de Alta Precisão 
        <span style="color:#F59E0B;">●</span> 100% Local e Seguro
    """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

# Sidebar com estatísticas (sessão)
with st.sidebar:
    st.markdown("### 📊 Estatísticas da Sessão")
    if 'total_processados' not in st.session_state:
        st.session_state.total_processados = 0
        st.session_state.total_linhas = 0
    st.metric("📄 PDFs Processados", st.session_state.total_processados)
    st.metric("📋 Linhas Extraídas", st.session_state.total_linhas)
    st.markdown("---")
    st.markdown("### ⚙️ Modo de Processamento")
    st.info("Seus dados **não saem** do seu computador. Conformidade LGPD garantida.")

# Opções principais
tipo = st.radio("1. Tipo de relatório:", 
                options=["fiscal", "hotel", "exames", "refeicoes"],
                format_func=lambda x: {
                    "fiscal": "Notas Fiscais (NAI 89701/92284/Misto)",
                    "hotel": "Diárias e Consumo (Plaza Hotel)",
                    "exames": "Exames Ocupacionais (Biomed)",
                    "refeicoes": "Mapa de Refeições"
                }[x])

modo_abas = st.radio("2. Organização do Excel:", ["unica", "separadas"], 
                     format_func=lambda x: "Uma única aba" if x == "unica" else "Uma aba por arquivo")

usar_ocr = False
if tipo == "fiscal":
    usar_ocr = st.checkbox("🔍 OCR Alta Precisão (recomendado para PDFs escaneados ou com falhas)")

arquivos = st.file_uploader("3. Selecione os PDFs:", type=['pdf'], accept_multiple_files=True)

# Botão de extração
if st.button("🚀 Extrair Dados", type="primary"):
    if not arquivos:
        st.warning("⚠️ Selecione pelo menos um arquivo PDF.")
    else:
        dados_totais = []
        rejeitados_totais = []
        stats_lista = []

        with st.spinner("Processando..."):
            for arquivo in arquivos:
                with gerenciar_memoria():
                    if tipo == "fiscal":
                        dados, rejeitados, stats = extrair_fiscal(arquivo, usar_ocr)
                        dados_totais.extend(dados)
                        rejeitados_totais.extend(rejeitados)
                        stats_lista.append(stats)
                        if stats.falhas == 0:
                            st.success(f"✅ {arquivo.name}: {stats.sucesso} linhas fiscais extraídas.")
                        else:
                            st.warning(f"⚠️ {arquivo.name}: {stats.sucesso} OK, {stats.falhas} falhas.")
                    
                    elif tipo == "hotel":
                        dados = extrair_hotel(arquivo)
                        dados_totais.extend(dados)
                        st.success(f"✅ {arquivo.name}: {len(dados)} itens de hotel extraídos.")
                        stats_lista.append(EstatisticasProcessamento(arquivo=arquivo.name, sucesso=len(dados)))
                    
                    elif tipo == "exames":
                        dados = extrair_exames(arquivo)
                        dados_totais.extend(dados)
                        st.success(f"✅ {arquivo.name}: {len(dados)} exames extraídos.")
                        stats_lista.append(EstatisticasProcessamento(arquivo=arquivo.name, sucesso=len(dados)))
                    
                    elif tipo == "refeicoes":
                        dados = extrair_refeicoes(arquivo)
                        dados_totais.extend(dados)
                        st.success(f"✅ {arquivo.name}: {len(dados)} registros de refeição extraídos.")
                        stats_lista.append(EstatisticasProcessamento(arquivo=arquivo.name, sucesso=len(dados)))

        # Atualiza estatísticas da sessão
        st.session_state.total_processados += len(arquivos)
        st.session_state.total_linhas += len(dados_totais)

        # Exibe resumo visual
        if stats_lista:
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("📄 Arquivos processados", len(arquivos))
            with col2:
                st.metric("📋 Linhas extraídas", len(dados_totais))
            with col3:
                taxa = sum(s.sucesso for s in stats_lista) / max(1, sum(s.total_linhas_encontradas for s in stats_lista)) * 100
                st.metric("✅ Taxa de sucesso", f"{taxa:.1f}%")

        if dados_totais:
            df = pd.DataFrame(dados_totais, columns=COLUNAS_CONFIG[tipo])
            df_erros = pd.DataFrame(rejeitados_totais) if rejeitados_totais else pd.DataFrame()
            excel_buffer = gerar_excel_formatado(df, df_erros, tipo, modo_abas)
            st.download_button(
                label="📥 Descarregar Excel",
                data=excel_buffer,
                file_name=f"Extracao_{tipo}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("ℹ️ Nenhum dado foi encontrado nos arquivos enviados para o tipo selecionado.")
