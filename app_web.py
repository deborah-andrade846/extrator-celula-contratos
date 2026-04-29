# -*- coding: utf-8 -*-
"""
Criado em Mon Apr  6 11:22:01 2026
@author: deborah.goncalves
Atualizado com extração de Notas Fiscais (PDF) – dois formatos suportados
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
        "Arquivo", "data_emissao", "numero_nf", "chave_nfe", "fornecedor", "uf", "ncm",
        "descricao", "cfop", "valor_total", "bc_icms", "icms", "percentual_interna",
        "icms_origem", "valor_difal"
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


# --- Fiscal (duas estratégias separadas) ---
def _extrair_formato_a(arquivo_pdf):
    """
    Formato com fornecedor (ex.: AnexoSNE89701).
    Estratégia: após data, NF e chave (44 dígitos), identifica a UF
    (duas letras maiúsculas) e, em seguida, localiza o bloco CFOP (4 dígitos)
    seguido de três valores monetários.
    """
    dados = []
    with pdfplumber.open(arquivo_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if not texto:
                continue
            for linha in texto.split('\n'):
                linha = linha.strip()
                if not linha or linha.startswith(('Data','TOTAIS','Página','ANEXO')):
                    continue

                # 1. Extrai data, número NF e chave de 44 dígitos
                match_inicio = re.match(
                    r'(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(\d{44})\s+(.+)', linha
                )
                if not match_inicio:
                    continue

                data = match_inicio.group(1)
                nf = match_inicio.group(2)
                chave = match_inicio.group(3)
                restante = match_inicio.group(4)

                # 2. Localiza a UF – duas letras maiúsculas "soltas"
                uf_match = re.search(r'\b([A-Z]{2})\b', restante)
                if not uf_match:
                    continue
                uf = uf_match.group(1)
                fornecedor = restante[:uf_match.start()].strip()
                after_uf = restante[uf_match.end():].strip()

                # 3. Encontra o CFOP (4 dígitos) seguido dos três valores monetários
                valores_match = re.search(
                    r'(\d{4})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)', after_uf
                )
                if not valores_match:
                    continue

                cfop = valores_match.group(1)
                valor_total = valores_match.group(2)
                icms_origem = valores_match.group(3)
                vr_difal = valores_match.group(4)

                # Descrição é tudo após UF e antes do CFOP
                descricao = after_uf[:valores_match.start()].strip()

                dados.append({
                    'data_emissao': data,
                    'numero_nf': nf,
                    'chave_nfe': chave,
                    'fornecedor': fornecedor,
                    'uf': uf,
                    'ncm': None,
                    'descricao': descricao,
                    'cfop': cfop,
                    'valor_total': valor_total,
                    'bc_icms': None,
                    'icms': None,
                    'percentual_interna': None,
                    'icms_origem': icms_origem,
                    'valor_difal': vr_difal,
                })
    return dados


def _extrair_formato_b(arquivo_pdf):
    """Formato B: Data | NF | Chave | UF | NCM | CFOP | Descrição | Valor NF | BC ICMS | ICMS | % Interna | DIFAL"""
    dados = []
    padrao = re.compile(
        r'(\d{2}/\d{2}/\d{4})\s+'
        r'(\d+)\s+'
        r'(\d{44})\s+'
        r'([A-Z]{2})\s+'
        r'(\d{4,12})\s+'
        r'(\d{4})\s+'
        r'(.+?)\s+'              # descrição
        r'([\d.,]+)\s+'          # valor NF
        r'([\d.,]+)\s+'          # BC ICMS
        r'([\d.,]+)\s+'          # ICMS
        r'([\d.,]+)\s+'          # % Interna
        r'([\d.,]+)'             # DIFAL
    )
    
    with pdfplumber.open(arquivo_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if not texto:
                continue
            for linha in texto.split('\n'):
                linha = linha.strip()
                if not linha or linha.startswith(('Data','TOTAIS','Página','ANEXO','OBS:')):
                    continue
                m = padrao.search(linha)
                if m:
                    dados.append({
                        'data_emissao': m.group(1),
                        'numero_nf': m.group(2),
                        'chave_nfe': m.group(3),
                        'fornecedor': None,
                        'uf': m.group(4),
                        'ncm': m.group(5),
                        'descricao': m.group(7).strip(),
                        'cfop': m.group(6),
                        'valor_total': m.group(8),
                        'bc_icms': m.group(9),
                        'icms': m.group(10),
                        'percentual_interna': m.group(11),
                        'icms_origem': None,
                        'valor_difal': m.group(12),
                    })
    return dados


def extrair_dados_fiscais_pdf(arquivo_pdf):
    """Tenta os dois formatos e retorna o que gerar mais dados."""
    dados_a = _extrair_formato_a(arquivo_pdf)
    dados_b = _extrair_formato_b(arquivo_pdf)
    
    # Escolhe o que tiver mais linhas
    dados = dados_a if len(dados_a) >= len(dados_b) else dados_b
    
    if not dados:
        raise ValueError("Nenhuma linha de dados fiscais foi encontrada no PDF.")
    
    df = pd.DataFrame(dados)
    df['data_emissao'] = pd.to_datetime(df['data_emissao'], dayfirst=True, errors='coerce')
    
    # Converte valores monetários
    def converter(v):
        if isinstance(v, str):
            return float(v.replace('.', '').replace(',', '.'))
        return v
    
    for col in ['valor_total','bc_icms','icms','percentual_interna','icms_origem','valor_difal']:
        if col in df.columns:
            df[col] = df[col].apply(converter)
    
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
                        df_fiscal = extrair_dados_fiscais_pdf(arquivo_pdf)
                        df_fiscal['Arquivo'] = nome_arquivo
                        dados_finais.extend(df_fiscal.to_dict('records'))

                except Exception as e:
                    st.error(f"Erro ao processar o ficheiro {nome_arquivo}:\n{str(e)}")
                    st.stop()

        # ==================== GERAR EXCEL ====================
        if dados_finais:
            st.success(f"Extração concluída! {len(dados_finais)} linhas extraídas.")

            df = pd.DataFrame(dados_finais)

            # Reordenar colunas: "Arquivo" primeiro, depois todas as outras
            colunas = [c for c in df.columns if c != 'Arquivo']
            colunas_presentes = ['Arquivo'] + [c for c in colunas if c in df.columns]
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
