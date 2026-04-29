# -*- coding: utf-8 -*-
"""
Extrator de Relatórios - Apoena
Inclui: Hotel, Exames, Refeições e Notas Fiscais (PDF)
"""

import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# ==================== CONFIGURAÇÕES ====================
COLUNAS_CONFIG = {
    "hotel": ["Arquivo", "Data", "Informação adicional", "Qtde", "Unidade", "Total"],
    "exames": ["Arquivo", "Exame", "Valor"],
    "refeicoes": ["Arquivo", "Data", "Total"],
    "fiscal": [
        "data_emissao", "numero_nf", "chave_nfe", "fornecedor", "uf",
        "descricao", "cfop", "valor_total", "icms_origem", "valor_difal"
    ]
}

# Mapeamento de cabeçalhos frequentes para colunas normalizadas
MAPA_COLUNAS_FISCAIS = {
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
    'icms origem': 'icms_origem',
    'icms': 'icms_origem',
    'vr difal': 'valor_difal',
    'difal': 'valor_difal',
}

# ==================== FUNÇÕES DE EXTRAÇÃO ====================

# --- Hotel (mantida) ---
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

# --- Fiscal (nova, para PDF) ---
def extrair_dados_fiscais_pdf(arquivo_pdf):
    """
    Extrai a tabela de notas fiscais de um PDF e retorna um DataFrame normalizado.
    """
    with pdfplumber.open(arquivo_pdf) as pdf:
        # Tentar extração automática da tabela (funciona para a maioria dos PDFs)
        tabelas = []
        for pagina in pdf.pages:
            tabela = pagina.extract_table()
            if tabela:
                tabelas.extend(tabela)

        if tabelas:
            return _processar_tabela_extraida(tabelas)
        else:
            # Fallback: extrair texto e parse linha a linha
            texto_completo = "\n".join([pagina.extract_text() or "" for pagina in pdf.pages])
            return _parse_texto_fiscal(texto_completo)

def _processar_tabela_extraida(tabela):
    """
    Recebe uma tabela extraída (lista de listas) e normaliza.
    """
    # A primeira linha é o cabeçalho
    cabecalho = tabela[0]
    # Mapear índices
    mapa_idx = {}
    for i, col_name in enumerate(cabecalho):
        if col_name is None:
            continue
        chave = str(col_name).strip().lower()
        if chave in MAPA_COLUNAS_FISCAIS:
            mapa_idx[i] = MAPA_COLUNAS_FISCAIS[chave]

    if not mapa_idx:
        raise ValueError("Não foi possível identificar as colunas no PDF.")

    # Processar linhas de dados
    dados = []
    for linha in tabela[1:]:
        registro = {}
        for idx, col_padrao in mapa_idx.items():
            valor = linha[idx] if idx < len(linha) else None
            registro[col_padrao] = valor.strip() if valor else None
        # Filtrar linhas completamente vazias
        if any(v for v in registro.values()):
            dados.append(registro)

    df = pd.DataFrame(dados)
    return _converter_tipos(df)

def _parse_texto_fiscal(texto):
    """
    Parse alternativo para quando a tabela não é detectada.
    Cobre os formatos dos seus dois prints.
    """
    linhas = texto.split('\n')
    # Identificar linha de cabeçalho (contém palavras-chave)
    cabecalho_idx = None
    for i, linha in enumerate(linhas):
        if any(palavra in linha.lower() for palavra in ['data', 'nº n.f', 'chave']):
            cabecalho_idx = i
            break
    if cabecalho_idx is None:
        raise ValueError("Cabeçalho da tabela fiscal não encontrado no PDF.")

    # Extrair nomes das colunas (podem estar concatenados, ex: "DataNFISCAL...")
    cabecalho_linha = linhas[cabecalho_idx]
    # Tentar separar por espaços múltiplos ou por palavras-chave conhecidas
    # Vamos usar uma abordagem mais robusta: procurar a posição de cada coluna pelo padrão
    colunas_ordenadas = [
        ('data', r'\bData\b'),
        ('numero_nf', r'(N[º°]\s*N\.?\s*F|N\.?\s*Fiscal)'),
        ('chave_nfe', r'Chave\s*(da\s*)?Nota\s*Fiscal\s*Eletr[ôó]nica'),
        ('fornecedor', r'Fornecedor'),
        ('uf', r'\bUF\b'),
        ('ncm', r'\bNCM\b'),
        ('cfop', r'\bCFOP\b'),
        ('descricao', r'Descri[çc][ãa]o\s*(da\s*Mercadoria|do\s*Documento)'),
        ('valor_total', r'Valor\s*(dos\s*Itens|NF\.?)'),
        ('icms_origem', r'ICMS\s*Origem'),
        ('valor_difal', r'(VR\s*DIFAL|DIFAL)'),
    ]
    posicoes = {}
    cabecalho_lower = cabecalho_linha.lower()
    for nome_padrao, regex in colunas_ordenadas:
        match = re.search(regex, cabecalho_lower)
        if match:
            posicoes[nome_padrao] = match.start()

    if not posicoes:
        raise ValueError("Não foi possível interpretar o cabeçalho fiscal.")

    # Ordenar colunas pela posição
    colunas_ordenadas = sorted(posicoes.keys(), key=lambda x: posicoes[x])

    # Agora processar as linhas de dados (a partir da linha seguinte)
    dados = []
    for linha in linhas[cabecalho_idx+1:]:
        linha = linha.strip()
        if not linha or linha.startswith('Página') or linha.startswith('Total'):
            continue
        # Tentar particionar a linha pelas posições das colunas
        registro = {}
        for i, col in enumerate(colunas_ordenadas):
            inicio = posicoes[col]
            fim = posicoes[colunas_ordenadas[i+1]] if i+1 < len(colunas_ordenadas) else len(linha)
            valor = linha[inicio:fim].strip()
            if valor:
                registro[col] = valor
        if registro:
            dados.append(registro)

    df = pd.DataFrame(dados)
    return _converter_tipos(df)

def _converter_tipos(df):
    """Converte colunas de data e numéricas (formato brasileiro)"""
    if 'data_emissao' in df.columns:
        df['data_emissao'] = pd.to_datetime(df['data_emissao'], dayfirst=True, errors='coerce')

    colunas_numericas = ['valor_total', 'icms_origem', 'valor_difal']
    for col in colunas_numericas:
        if col in df.columns:
            # Remover pontos de milhar e trocar vírgula decimal
            df[col] = df[col].astype(str).str.replace('.', '', regex=False)
            df[col] = df[col].str.replace(',', '.', regex=False)
            df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

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
        "fiscal": "Notas Fiscais (PDF)"
    }[x]
)

st.markdown("### 2. Selecione os ficheiros PDF:")
arquivos_selecionados = st.file_uploader(
    "Arraste e solte ou clique para procurar",
    type=['pdf'],
    accept_multiple_files=True
)

# ==================== BOTÃO DE EXTRAÇÃO ====================
if st.button("Extrair Dados e Gerar Excel", type="primary"):
    if not arquivos_selecionados:
        st.warning("Por favor, selecione pelo menos um ficheiro PDF.")
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
                                if any(p in linha for p in ["PLAZA HOTEL", "Apartamento:", "Fechado", "Pagamentos", "Tarifário:"]):
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
                        # Nova extração fiscal a partir de PDF
                        df_fiscal = extrair_dados_fiscais_pdf(arquivo_pdf)
                        # Adicionar nome do arquivo como coluna informativa
                        df_fiscal['Arquivo'] = nome_arquivo
                        dados_finais.extend(df_fiscal.to_dict('records'))

                except Exception as e:
                    st.error(f"Erro ao processar o ficheiro {nome_arquivo}:\n{str(e)}")
                    st.stop()

        # ==================== GERAR EXCEL ====================
        if dados_finais:
            st.success(f"Extração concluída! {len(dados_finais)} linhas extraídas.")
            
            df = pd.DataFrame(dados_finais)
            # Reordenar colunas conforme CONFIG, mantendo apenas as que existem
            colunas_esperadas = COLUNAS_CONFIG[tipo_selecionado]
            colunas_presentes = [c for c in colunas_esperadas if c in df.columns]
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
