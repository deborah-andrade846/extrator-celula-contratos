# -*- coding: utf-8 -*-
"""
Criado em Mon Apr  6 11:22:01 2026
@author: deborah.goncalves
Atualizado: extração fiscal robusta a textos concatenados (formato A)
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

UFS_VALIDAS = {
    'AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO',
    'MA', 'MT', 'MS', 'MG', 'PA', 'PB', 'PR', 'PE', 'PI',
    'RJ', 'RN', 'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'
}

# ==================== FUNÇÕES DE HOTEL, EXAMES, REFEIÇÕES (mantidas) ====================

def limpar_linha_hotel(linha, nome_hospede):
    # ... (inalterado)
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


def _extrair_formato_a(arquivo_pdf):
    """
    Formato A (com fornecedor). Trata textos concatenados (sem espaços).
    """
    dados = []
    total_fiscais = 0
    nao_capturadas = 0

    # Padrão para UF: duas letras maiúsculas isoladas, que sejam uma UF válida.
    uf_pattern = re.compile(r'(?<![A-Z])(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)(?![A-Z])')

    # Padrão flexível para o bloco final: CFOP (4 dígitos) seguido de três valores monetários (com vírgula, 2 decimais)
    # Aceita espaços ou pontos entre os grupos. NÃO exige fim de linha – captura a última ocorrência.
    bloco_pattern = re.compile(r'(\d{4})\s*([\d.]+,\d{2})\s*([\d.]+,\d{2})\s*([\d.]+,\d{2})')

    with pdfplumber.open(arquivo_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if not texto:
                continue
            for linha in texto.split('\n'):
                linha = linha.strip()
                if not linha or linha.startswith(('Data','TOTAIS','Página','ANEXO')):
                    continue

                # Pré‑filtro obrigatório: data + NF + chave de 44 dígitos
                inicio = re.match(r'(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(\d{44})\s*(.+)', linha)
                if not inicio:
                    continue

                total_fiscais += 1
                data = inicio.group(1)
                nf = inicio.group(2)
                chave = inicio.group(3)
                restante = inicio.group(4).strip()

                # ---- Pré-processamento: separar dígitos de letras (ex: "21012.713,58" → "2101 2.713,58")
                restante = re.sub(r'(?<=[A-Za-z])(?=\d)|(?<=\d)(?=[A-Za-z])', ' ', restante)

                # ---- Localizar UF e fornecedor ----
                uf_match = uf_pattern.search(restante)
                if uf_match:
                    uf = uf_match.group(1)
                    fornecedor = restante[:uf_match.start()].strip()
                    after_uf = restante[uf_match.end():].strip()
                else:
                    uf = ""
                    fornecedor = ""
                    after_uf = restante   # tudo após a chave

                # ---- Extrair bloco CFOP + 3 valores (última ocorrência) ----
                # Vamos procurar todas as ocorrências e pegar a última
                matches = list(bloco_pattern.finditer(after_uf))
                if matches:
                    bloco = matches[-1]  # último conjunto CFOP+valores
                    cfop = bloco.group(1)
                    valor_total = bloco.group(2)
                    icms_origem = bloco.group(3)
                    vr_difal = bloco.group(4)
                    descricao = after_uf[:bloco.start()].strip()
                else:
                    # ---- Fallback melhorado: captura números com vírgula (2 casas decimais) e tenta localizar o CFOP antes deles
                    tokens_vals = re.findall(r'[\d.]+\d,\d{2}', after_uf)
                    if len(tokens_vals) >= 3:
                        # Pega os três últimos valores
                        vr_difal = tokens_vals[-1]
                        icms_origem = tokens_vals[-2]
                        valor_total = tokens_vals[-3]
                        # Procura um CFOP imediatamente antes do primeiro valor
                        idx_val = after_uf.rfind(tokens_vals[-3])
                        antes_vals = after_uf[:idx_val]
                        cfop_match = re.search(r'(\d{4})\s*$', antes_vals)
                        cfop = cfop_match.group(1) if cfop_match else ""
                        descricao = antes_vals[:cfop_match.start()].strip() if cfop_match else antes_vals.strip()
                    else:
                        # Fallback extremo: converte números com mais de 2 casas decimais para 2 casas (arredondando)
                        tokens_brutos = re.findall(r'[\d.]+\d,\d+', after_uf)
                        tokens_arredondados = []
                        for t in tokens_brutos:
                            try:
                                num = float(t.replace('.', '').replace(',', '.'))
                                tokens_arredondados.append(f"{num:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                            except:
                                tokens_arredondados.append(t)
                        if len(tokens_arredondados) >= 3:
                            vr_difal = tokens_arredondados[-1]
                            icms_origem = tokens_arredondados[-2]
                            valor_total = tokens_arredondados[-3]
                            idx_val = after_uf.rfind(tokens_brutos[-3])
                            antes_vals = after_uf[:idx_val]
                            cfop_match = re.search(r'(\d{4})\s*$', antes_vals)
                            cfop = cfop_match.group(1) if cfop_match else ""
                            descricao = antes_vals[:cfop_match.start()].strip() if cfop_match else antes_vals.strip()
                        else:
                            cfop = ""
                            valor_total = ""
                            icms_origem = ""
                            vr_difal = ""
                            descricao = after_uf

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
    return dados, total_fiscais, nao_capturadas


def _extrair_formato_b(arquivo_pdf):
    """Formato B (com NCM). Mantido como antes."""
    dados = []
    total_fiscais = 0
    nao_capturadas = 0
    padrao = re.compile(
        r'(\d{2}/\d{2}/\d{4})\s+'
        r'(\d+)\s+'
        r'(\d{44})\s+'
        r'([A-Z]{2})\s+'
        r'(\d{4,12})\s+'
        r'(\d{4})\s+'
        r'(.+?)\s+'
        r'([\d.,]+)\s+'
        r'([\d.,]+)\s+'
        r'([\d.,]+)\s+'
        r'([\d.,]+)\s*'
        r'([\d.,]+)'
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
                if not re.match(r'\d{2}/\d{2}/\d{4}\s+\d+\s+\d{44}', linha):
                    continue
                total_fiscais += 1
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
                else:
                    nao_capturadas += 1
    return dados, total_fiscais, nao_capturadas


def extrair_dados_fiscais_pdf(arquivo_pdf):
    dados_a, tot_a, nao_a = _extrair_formato_a(arquivo_pdf)
    dados_b, tot_b, nao_b = _extrair_formato_b(arquivo_pdf)

    if len(dados_a) >= len(dados_b):
        dados, total_fiscais, nao_capturadas = dados_a, tot_a, nao_a
        formato = "A (com fornecedor)"
    else:
        dados, total_fiscais, nao_capturadas = dados_b, tot_b, nao_b
        formato = "B (com NCM)"

    if not dados:
        raise ValueError("Nenhuma linha fiscal encontrada.")

    df = pd.DataFrame(dados)
    df['data_emissao'] = pd.to_datetime(df['data_emissao'], dayfirst=True, errors='coerce')

    def converter(v):
        if isinstance(v, str) and v.strip():
            return float(v.replace('.', '').replace(',', '.'))
        return None

    for col in ['valor_total','bc_icms','icms','percentual_interna','icms_origem','valor_difal']:
        if col in df.columns:
            df[col] = df[col].apply(converter)

    estatisticas = {
        'Formato': formato,
        'Linhas fiscais': total_fiscais,
        'Capturadas': len(dados),
        'Não capturadas': nao_capturadas,
        'Percentual': f"{(len(dados)/total_fiscais*100):.1f}%" if total_fiscais else "N/A"
    }
    return df, estatisticas


# ==================== INTERFACE (mantida) ====================
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
    "Arraste e solte ou clique para procurar", type=['pdf'], accept_multiple_files=True
)

if st.button("Extrair Dados e Gerar Excel", type="primary"):
    if not arquivos_selecionados:
        st.warning("Selecione pelo menos um ficheiro.")
    else:
        dados_finais = []
        estatisticas_gerais = []

        with st.spinner("Processando..."):
            for arquivo_pdf in arquivos_selecionados:
                nome_arquivo = arquivo_pdf.name
                try:
                    if tipo_selecionado == "hotel":
                        # ... (código hotel mantido)
                        with pdfplumber.open(arquivo_pdf) as pdf:
                            texto_completo = "\n".join([p.extract_text() or "" for p in pdf.pages])
                            linhas = texto_completo.split('\n')
                            nome_hospede = "NÃO_IDENTIFICADO"
                            for linha in linhas:
                                if "Hóspede principal:" in linha:
                                    try:
                                        nome_cru = linha.split("Hóspede principal:")[1].split("|")[0].strip()
                                        nome_hospede = nome_cru.split()[0]
                                    except: pass
                                    continue
                                if any(p in linha for p in ["PLAZA HOTEL", "Apartamento:", "Fechado", "Pagamentos", "Tarifário:"]):
                                    continue
                                li = limpar_linha_hotel(linha, nome_hospede)
                                if li: dados_finais.append(li)

                    elif tipo_selecionado == "exames":
                        with pdfplumber.open(arquivo_pdf) as pdf:
                            texto = "\n".join([p.extract_text() or "" for p in pdf.pages])
                            for linha in texto.split('\n'):
                                if "R$" in linha:
                                    partes = linha.split("R$")
                                    if len(partes) >= 2:
                                        nome = partes[0].replace('"','').replace(',','').strip()
                                        valor = "R$ " + partes[1].replace('"','').replace(',','').strip()
                                        if nome:
                                            dados_finais.append({"Arquivo": nome_arquivo, "Exame": nome, "Valor": valor})

                    elif tipo_selecionado == "refeicoes":
                        with pdfplumber.open(arquivo_pdf) as pdf:
                            texto = "\n".join([p.extract_text() or "" for p in pdf.pages])
                            linhas = texto.split('\n')
                            data_ref = "DATA_NAO_ENCONTRADA"
                            total_ref = "TOTAL_NAO_ENCONTRADO"
                            for linha in linhas:
                                if "Período:" in linha:
                                    m = re.search(r'\d{2}/\d{2}/\d{4}', linha)
                                    if m: data_ref = m.group(0)
                                if "Total Geral" in linha:
                                    v = linha.replace("Total Geral","").replace("|","").strip()
                                    if v: total_ref = v
                            if data_ref != "DATA_NAO_ENCONTRADA" or total_ref != "TOTAL_NAO_ENCONTRADO":
                                dados_finais.append({"Arquivo": nome_arquivo, "Data": data_ref, "Total": total_ref})

                    elif tipo_selecionado == "fiscal":
                        df_fiscal, stats = extrair_dados_fiscais_pdf(arquivo_pdf)
                        df_fiscal['Arquivo'] = nome_arquivo
                        dados_finais.extend(df_fiscal.to_dict('records'))
                        stats['Arquivo'] = nome_arquivo
                        estatisticas_gerais.append(stats)

                except Exception as e:
                    st.error(f"Erro no ficheiro {nome_arquivo}:\n{str(e)}")
                    st.stop()

        if tipo_selecionado == "fiscal" and estatisticas_gerais:
            st.markdown("### 📊 Resumo da extração")
            for s in estatisticas_gerais:
                with st.expander(f"📄 {s['Arquivo']} ({s['Formato']})"):
                    st.write(f"✅ Capturadas: **{s['Capturadas']}**")
                    st.write(f"📄 Linhas fiscais: **{s['Linhas fiscais']}**")
                    st.write(f"⚠️ Não capturadas: **{s['Não capturadas']}**")
                    st.write(f"📈 Percentual: **{s['Percentual']}**")

        if dados_finais:
            st.success(f"Extração concluída! {len(dados_finais)} linhas.")
            df = pd.DataFrame(dados_finais)
            colunas = ['Arquivo'] + [c for c in df.columns if c != 'Arquivo']
            df = df[colunas]
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            st.download_button("📥 Descarregar Excel", data=buffer.getvalue(),
                               file_name=f"Extracao_{tipo_selecionado}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("Nenhum dado válido encontrado.")
