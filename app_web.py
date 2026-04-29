# -*- coding: utf-8 -*-
"""
Criado em Mon Apr  6 11:22:01 2026
@author: deborah.goncalves
Atualizado com extração de Notas Fiscais (PDF)
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
        "Arquivo", "data_emissao", "numero_nf", "chave_nfe", "fornecedor", "uf",
        "descricao", "cfop", "valor_total", "icms_origem", "valor_difal"
    ]
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
    Extrai dados fiscais de um PDF de notas fiscais (Anexo I) e retorna um DataFrame.
    Analisa o texto plano de todas as páginas.
    """
    # Padrão regex para uma linha de dados completa
    padrao_linha = re.compile(
        r'^'
        r'(\d{2}/\d{2}/\d{4})\s+'          # data (grupo 1)
        r'(\d+)\s+'                        # nº nf (grupo 2)
        r'(\d{44})\s+'                     # chave NFe (grupo 3)
        r'(.+?)\s+'                        # fornecedor (grupo 4) – captura mínima até a UF
        r'([A-Z]{2})\s+'                   # UF (grupo 5)
        r'(.+?)\s+'                        # descrição (grupo 6) – termina antes do CFOP
        r'(\d{4})\s+'                      # CFOP (grupo 7)
        r'([\d.]+,\d{2})\s+'              # valor total (grupo 8)
        r'([\d.]+,\d{2})\s+'              # ICMS origem (grupo 9)
        r'([\d.]+,\d{2})'                 # VR DIFAL (grupo 10)
        r'\s*$'
    )

    # Padrão para identificar linhas que devem ser ignoradas
    padrao_ignorar = re.compile(
        r'^\s*(Data|TOTAIS\s+DO\s+MÊS|TOTAL\s+DO\s+MÊS|Página|TERMO DE CIÊNCIA|ANEXO|CONTRIBUINTE|DEMONSTRATIVO|Código da Infração|OBS:)',
        re.IGNORECASE
    )

    dados = []
    with pdfplumber.open(arquivo_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if not texto:
                continue
            linhas = texto.split('\n')
            for linha in linhas:
                linha = linha.strip()
                if not linha or padrao_ignorar.match(linha):
                    continue
                match = padrao_linha.match(linha)
                if match:
                    dados.append({
                        'data_emissao': match.group(1),
                        'numero_nf': match.group(2),
                        'chave_nfe': match.group(3),
                        'fornecedor': match.group(4).strip(),
                        'uf': match.group(5),
                        'descricao': match.group(6).strip(),
                        'cfop': match.group(7),
                        'valor_total': match.group(8),
                        'icms_origem': match.group(9),
                        'valor_difal': match.group(10),
                    })

    if not dados:
        raise ValueError("Nenhuma linha de dados fiscais foi encontrada no PDF. Verifique se o arquivo contém o Anexo I com as colunas Data, Nº N.F, Chave, etc.")

    df = pd.DataFrame(dados)

    # Conversão de tipos
    df['data_emissao'] = pd.to_datetime(df['data_emissao'], dayfirst=True, errors='coerce')

    # Converte valores monetários de formato brasileiro (1.234,56) para float
    def converter_moeda(valor):
        if isinstance(valor, str):
            return float(valor.replace('.', '').replace(',', '.'))
        return valor

    for col in ['valor_total', 'icms_origem', 'valor_difal']:
        df[col] = df[col].apply(converter_moeda)

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
