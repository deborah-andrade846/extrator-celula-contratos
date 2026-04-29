# -*- coding: utf-8 -*-
"""
Criado em Mon Apr  6 11:22:01 2026
@author: deborah.goncalves
Atualizado com extração de Notas Fiscais (Excel/CSV)
"""

import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from io import BytesIO

# ==================== CONFIGURAÇÕES ====================
COLUNAS_CONFIG = {
    "hotel": ["Arquivo", "Data", "Informação adicional", "Qtde", "Unidade", "Total"],
    "exames": ["Arquivo", "Exame", "Valor"],
    "refeicoes": ["Arquivo", "Data", "Total"],
    "fiscal": [
        "data_emissao", "numero_nf", "chave_nfe", "fornecedor", "uf",
        "descricao", "cfop", "valor_total", "icms_origem", "valor_difal",
        "ncm", "bc_icms", "percentual_interna", "observacao"
    ]
}

# Mapeamento padrão para colunas fiscais (chaves em minúsculo)
MAPA_COLUNAS_PADRAO = {
    'data': 'data_emissao',
    'data emissão': 'data_emissao',
    'nº n.f': 'numero_nf',
    'n. fiscal': 'numero_nf',
    'nº nota': 'numero_nf',
    'nf': 'numero_nf',
    'chave da nota fiscal eletrônica': 'chave_nfe',
    'chave da nota fiscal eletrónica': 'chave_nfe',
    'chave nfe': 'chave_nfe',
    'fornecedor': 'fornecedor',
    'uf': 'uf',
    'ncm': 'ncm',
    'cfop': 'cfop',
    'descrição da mercadoria': 'descricao',
    'descrição da mercadoria/serviço': 'descricao',
    'descrição do documento': 'descricao',
    'descricao': 'descricao',
    'valor dos itens': 'valor_total',
    'valor nf.': 'valor_total',
    'valor total': 'valor_total',
    'bc icms': 'bc_icms',
    'base icms': 'bc_icms',
    'icms origem': 'icms_origem',
    'icms': 'icms_origem',
    '% interna': 'percentual_interna',
    'alíquota interna': 'percentual_interna',
    'vr difal': 'valor_difal',
    'difal': 'valor_difal',
    'obs': 'observacao',
}

COLUNAS_OBRIGATORIAS = ['data_emissao', 'numero_nf', 'valor_total']

# ==================== FUNÇÕES DE EXTRAÇÃO ====================

# --- Hotel ---
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

# --- Fiscal (nova) ---
def extrair_dados_fiscais(arquivo, mapeamento=None, colunas_obrigatorias=None):
    """
    Extrai dados fiscais de um arquivo Excel ou CSV, normalizando as colunas.
    """
    if mapeamento is None:
        mapeamento = MAPA_COLUNAS_PADRAO
    if colunas_obrigatorias is None:
        colunas_obrigatorias = COLUNAS_OBRIGATORIAS

    # 1. Carregar o arquivo conforme extensão
    if isinstance(arquivo, str):
        nome_arquivo = arquivo
        if nome_arquivo.lower().endswith('.csv'):
            df_raw = _ler_csv_com_delimitador(nome_arquivo)
        else:
            df_raw = pd.read_excel(nome_arquivo, header=None, dtype=str)
    else:
        nome_arquivo = getattr(arquivo, 'name', 'arquivo.csv')
        ext = nome_arquivo.split('.')[-1].lower() if '.' in nome_arquivo else ''
        if ext == 'csv':
            content = arquivo.read()
            df_raw = _ler_csv_com_delimitador(BytesIO(content))
        else:
            df_raw = pd.read_excel(arquivo, header=None, dtype=str)

    # Se múltiplas abas, tentar cada uma
    if isinstance(df_raw, dict):
        for sheet_name, df_sheet in df_raw.items():
            try:
                return _processar_sheet(df_sheet, mapeamento, colunas_obrigatorias)
            except ValueError:
                continue
        raise ValueError("Nenhuma aba contém as colunas obrigatórias.")

    return _processar_sheet(df_raw, mapeamento, colunas_obrigatorias)

def _ler_csv_com_delimitador(arquivo):
    """Tenta ler CSV detectando delimitador."""
    try:
        if isinstance(arquivo, BytesIO):
            raw_bytes = arquivo.read()
            arquivo.seek(0)
            sample = raw_bytes[:4096].decode('utf-8', errors='ignore')
        else:
            with open(arquivo, 'rb') as f:
                sample = f.read(4096).decode('utf-8', errors='ignore')

        delimitadores = [',', ';', '\t', '|']
        contagens = {d: sample.count(d) for d in delimitadores}
        melhor = max(contagens, key=contagens.get)
        if contagens[melhor] < 2:
            melhor = ','
        if isinstance(arquivo, BytesIO):
            arquivo.seek(0)
        return pd.read_csv(arquivo, sep=melhor, dtype=str, header=None, encoding='utf-8')
    except Exception:
        if isinstance(arquivo, BytesIO):
            arquivo.seek(0)
        return pd.read_csv(arquivo, sep=',', dtype=str, header=None, encoding='latin1')

def _processar_sheet(df_raw, mapeamento, colunas_obrigatorias):
    """Encontra cabeçalho e extrai dados de um DataFrame bruto."""
    df = df_raw.astype(str).applymap(lambda x: x.strip())
    df.replace(['nan', 'None', '', ' '], pd.NA, inplace=True)

    idx_cabecalho = None
    chaves_normalizadas = {k.lower(): v for k, v in mapeamento.items()}

    for i, row in df.iterrows():
        valores_linha = [str(v).lower().strip() for v in row if pd.notna(v)]
        matches = sum(1 for v in valores_linha if v in chaves_normalizadas)
        if matches >= max(2, len(colunas_obrigatorias) // 2):
            idx_cabecalho = i
            cabecalho_original = row.tolist()
            break

    if idx_cabecalho is None:
        raise ValueError("Não foi possível localizar a linha de cabeçalho. Verifique o arquivo.")

    mapa_colunas_pos = {}
    for idx_col, nome_original in enumerate(cabecalho_original):
        if pd.isna(nome_original):
            continue
        chave = str(nome_original).lower().strip()
        if chave in chaves_normalizadas:
            mapa_colunas_pos[idx_col] = chaves_normalizadas[chave]

    padronizadas_presentes = set(mapa_colunas_pos.values())
    faltantes = set(colunas_obrigatorias) - padronizadas_presentes
    if faltantes:
        raise ValueError(f"Colunas obrigatórias não encontradas: {faltantes}. "
                         f"Cabeçalhos detectados: {list(cabecalho_original)}")

    dados = df.iloc[idx_cabecalho + 1:].copy()
    dados = dados.dropna(how='all')
    mascara_titulo = dados.apply(lambda r: r.astype(str).str.match(r'^(Total|Subtotal|Pagina|Página)\b').any(), axis=1)
    dados = dados[~mascara_titulo]

    if dados.empty:
        raise ValueError("Nenhuma linha de dados encontrada após o cabeçalho.")

    colunas_utilizadas = sorted(mapa_colunas_pos.keys())
    df_final = dados.iloc[:, colunas_utilizadas].copy()
    df_final.columns = [mapa_colunas_pos[c] for c in colunas_utilizadas]

    # Conversão de tipos
    if 'data_emissao' in df_final.columns:
        df_final['data_emissao'] = pd.to_datetime(df_final['data_emissao'], dayfirst=True, errors='coerce')

    colunas_numericas = ['valor_total', 'bc_icms', 'icms_origem', 'percentual_interna', 'valor_difal']
    for col in colunas_numericas:
        if col in df_final.columns:
            df_final[col] = df_final[col].astype(str).str.replace('.', '', regex=False)
            df_final[col] = df_final[col].str.replace(',', '.', regex=False)
            df_final[col] = pd.to_numeric(df_final[col], errors='coerce')

    return df_final.reset_index(drop=True)

# ==================== INTERFACE STREAMLIT ====================
st.set_page_config(page_title="Extrator de Relatórios - Apoena", page_icon="📊")
st.title("📊 Extrator de Relatórios - Apoena")

st.markdown("### 1. O que deseja extrair?")
tipo_selecionado = st.radio(
    "Escolha o tipo de relatório:",
    options=["hotel", "exames", "refeicoes", "fiscal"],
    format_func=lambda x: {
        "hotel": "Diárias e Consumo (Plaza Hotel)",
        "exames": "Exames Ocupacionais (Biomed)",
        "refeicoes": "Mapa de Refeições",
        "fiscal": "Notas Fiscais (Excel/CSV)"
    }[x]
)

st.markdown("### 2. Selecione os ficheiros:")
if tipo_selecionado == "fiscal":
    arquivos_selecionados = st.file_uploader(
        "Arraste e solte ficheiros Excel ou CSV",
        type=['xlsx', 'xls', 'csv'],
        accept_multiple_files=True
    )
else:
    arquivos_selecionados = st.file_uploader(
        "Arraste e solte ficheiros PDF",
        type=['pdf'],
        accept_multiple_files=True
    )

# ==================== BOTÃO DE EXTRAÇÃO ====================
if st.button("Extrair Dados e Gerar Excel", type="primary"):
    if not arquivos_selecionados:
        st.warning("Por favor, selecione pelo menos um ficheiro.")
    else:
        dados_finais = []
        
        with st.spinner("A processar ficheiros..."):
            for arquivo_pdf in arquivos_selecionados:
                nome_arquivo = arquivo_pdf.name
                
                try:
                    if tipo_selecionado == "hotel":
                        with pdfplumber.open(arquivo_pdf) as pdf:
                            texto_completo = "\n".join([pagina.extract_text() or "" for pagina in pdf.pages])
                            linhas = texto_completo.split('\n')
                            nome_hospede = "NÃO_IDENTIFICADO"
                            for linha in linhas:
                                if "Hóspede principal:" in linha:
                                    try:
                                        nome_cru = linha.split("Hóspede principal:")[1].split("|")[0].strip()
                                        nome_hospede = nome_cru.split()[0]
                                    except:
                                        pass
                                    continue
                                if "PLAZA HOTEL" in linha or "Apartamento:" in linha or "Fechado" in linha or "Pagamentos" in linha or "Tarifário:" in linha:
                                    continue
                                linha_extraida = limpar_linha_hotel(linha, nome_hospede)
                                if linha_extraida:
                                    dados_finais.append(linha_extraida)

                    elif tipo_selecionado == "exames":
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
                                            dados_finais.append({
                                                "Arquivo": nome_arquivo,
                                                "Exame": nome_exame,
                                                "Valor": valor_exame
                                            })

                    elif tipo_selecionado == "refeicoes":
                        with pdfplumber.open(arquivo_pdf) as pdf:
                            texto_completo = "\n".join([pagina.extract_text() or "" for pagina in pdf.pages])
                            linhas = texto_completo.split('\n')
                            data_refeicao = "DATA_NAO_ENCONTRADA"
                            total_refeicoes = "TOTAL_NAO_ENCONTRADO"
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
                                dados_finais.append({
                                    "Arquivo": nome_arquivo,
                                    "Data": data_refeicao,
                                    "Total": total_refeicoes
                                })

                    elif tipo_selecionado == "fiscal":
                        df_extraido = extrair_dados_fiscais(arquivo_pdf)
                        # Converte o DataFrame para lista de dicionários e adiciona coluna do arquivo
                        registros = df_extraido.to_dict('records')
                        for registro in registros:
                            registro['Arquivo'] = nome_arquivo
                            dados_finais.append(registro)

                except Exception as e:
                    st.error(f"Erro ao processar o ficheiro {nome_arquivo}:\n{str(e)}")
                    st.stop()

        # ==================== GERAR EXCEL ====================
        if dados_finais:
            st.success(f"Extração concluída com sucesso! Foram encontradas {len(dados_finais)} linhas.")
            
            # Definir ordem das colunas conforme o tipo
            colunas_excel = COLUNAS_CONFIG[tipo_selecionado]
            # Para fiscal, o DataFrame pode ter colunas extras que não estão em colunas_excel
            # Vamos filtrar apenas as que existem nos dados
            df = pd.DataFrame(dados_finais)
            colunas_presentes = [c for c in colunas_excel if c in df.columns]
            df = df[colunas_presentes]
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Descarregar Tabela em Excel",
                data=buffer.getvalue(),
                file_name=f"Extracao_{tipo_selecionado}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("A extração não encontrou dados válidos para o tipo selecionado nestes ficheiros.")
