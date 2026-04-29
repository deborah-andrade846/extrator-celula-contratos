# -*- coding: utf-8 -*-
"""
Extrator de Relatórios - Apoena
Atualizado com o módulo Fiscal usando Regex Espacial
"""

import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# 1. Configuração direta das colunas
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

# 2. Funções Auxiliares
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

def converter_para_numero(valor_str):
    """Converte os valores com vírgula de Portugal/Brasil para decimais matemáticos (floats)"""
    if not valor_str:
        return None
    try:
        return float(valor_str.replace('.', '').replace(',', '.'))
    except:
        return valor_str

def extrair_linhas_fiscal(arquivo_pdf):
    """Extrai os dados fiscais utilizando layout espacial e as Regex precisas"""
    dados_locais = []
    
    # Regex adaptadas dos seus ficheiros para ler o espaçamento exato
    padrao_89701 = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(\d{44})\s+(.*?)\s+([A-Z]{2})\s+(.*?)\s+(\d{4})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)$')
    padrao_92284 = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(\d{44})\s+([A-Z]{2})\s+(\d{8})\s+(\d{4})\s+(.*?)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*(.*)$')

    with pdfplumber.open(arquivo_pdf) as pdf:
        for pagina in pdf.pages:
            # O layout=True imita o funcionamento do pdftotext
            texto = pagina.extract_text(layout=True)
            if not texto: continue
            
            for linha in texto.split('\n'):
                linha = linha.strip()
                if not linha: continue
                
                # Tenta encaixar no formato 89701
                m1 = padrao_89701.match(linha)
                if m1:
                    dados_locais.append({
                        "Arquivo": arquivo_pdf.name,
                        "Tipo NAI": "89701",
                        "Data": m1.group(1),
                        "Nº NF": m1.group(2),
                        "Chave da NF-e": m1.group(3),
                        "Fornecedor": m1.group(4).strip(),
                        "UF": m1.group(5),
                        "NCM": "",
                        "Descrição": m1.group(6).strip(),
                        "CFOP": m1.group(7),
                        "Valor NF": converter_para_numero(m1.group(8)),
                        "BC ICMS": None,
                        "ICMS": None,
                        "% Interna": None,
                        "ICMS Origem": converter_para_numero(m1.group(9)),
                        "VR DIFAL": converter_para_numero(m1.group(10)),
                        "OBS": ""
                    })
                    continue
                    
                # Tenta encaixar no formato 92284
                m2 = padrao_92284.match(linha)
                if m2:
                    dados_locais.append({
                        "Arquivo": arquivo_pdf.name,
                        "Tipo NAI": "92284",
                        "Data": m2.group(1),
                        "Nº NF": m2.group(2),
                        "Chave da NF-e": m2.group(3),
                        "Fornecedor": "",
                        "UF": m2.group(4),
                        "NCM": m2.group(5),
                        "Descrição": m2.group(7).strip(),
                        "CFOP": m2.group(6),
                        "Valor NF": converter_para_numero(m2.group(8)),
                        "BC ICMS": converter_para_numero(m2.group(9)),
                        "ICMS": converter_para_numero(m2.group(10)),
                        "% Interna": converter_para_numero(m2.group(11)),
                        "ICMS Origem": None,
                        "VR DIFAL": converter_para_numero(m2.group(12)),
                        "OBS": m2.group(13).strip()
                    })
                    
    return dados_locais

# 3. Interface Visual do Streamlit
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
        "fiscal": "Notas Fiscais (Autuação SEFAZ)"
    }[x]
)

st.markdown("### 2. Selecione os ficheiros PDF:")
arquivos_selecionados = st.file_uploader("Arraste e solte ou clique para procurar", type=['pdf'], accept_multiple_files=True)

# 4. Botão de Extração
if st.button("Extrair Dados e Gerar Excel", type="primary"):
    if not arquivos_selecionados:
        st.warning("Por favor, selecione pelo menos um ficheiro PDF.")
    else:
        dados_finais = []
        
        with st.spinner("A processar ficheiros..."):
            for arquivo_pdf in arquivos_selecionados:
                nome_arquivo = arquivo_pdf.name
                
                try:
                    # Ramo Fiscal usa a nova função especializada
                    if tipo_selecionado == "fiscal":
                        dados_finais.extend(extrair_linhas_fiscal(arquivo_pdf))
                        continue
                    
                    # Ramos antigos continuam como estavam
                    with pdfplumber.open(arquivo_pdf) as pdf:
                        texto_completo = "\n".join([pagina.extract_text() or "" for pagina in pdf.pages])
                        linhas = texto_completo.split('\n')
                        
                        if tipo_selecionado == "hotel":
                            nome_hospede = "NÃO_IDENTIFICADO"
                            for linha in linhas:
                                if "Hóspede principal:" in linha:
                                    try:
                                        nome_cru = linha.split("Hóspede principal:")[1].split("|")[0].strip()
                                        nome_hospede = nome_cru.split()[0]
                                    except: pass
                                    continue
                                if "PLAZA HOTEL" in linha or "Apartamento:" in linha or "Fechado" in linha or "Pagamentos" in linha or "Tarifário:" in linha:
                                    continue
                                linha_extraida = limpar_linha_hotel(linha, nome_hospede)
                                if linha_extraida:
                                    dados_finais.append(linha_extraida)

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

                        elif tipo_selecionado == "refeicoes":
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

                except Exception as e:
                    st.error(f"Erro ao processar o ficheiro {nome_arquivo}:\n{str(e)}")
                    st.stop()

        # 5. Gerar e oferecer o Excel para download
        if dados_finais:
            st.success(f"Extração concluída com sucesso! Foram encontradas {len(dados_finais)} linhas.")
            
            df = pd.DataFrame(dados_finais, columns=COLUNAS_CONFIG[tipo_selecionado])
            buffer = io.BytesIO()
            
            # Exportação inteligente com XlsxWriter
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Extração')
                
                # Se for fiscal, formata automaticamente as colunas como moeda contábil
                if tipo_selecionado == "fiscal":
                    workbook = writer.book
                    worksheet = writer.sheets['Extração']
                    formato_moeda = workbook.add_format({'num_format': '#,##0.00'})
                    
                    # Colunas do Excel (A=0, B=1, etc). As colunas monetárias estão no índice 10 ao 15.
                    worksheet.set_column('A:Q', 15) # Alarga todas um pouco
                    worksheet.set_column('I:I', 45) # Alarga a Descrição
                    worksheet.set_column('K:P', 14, formato_moeda) # Aplica formato financeiro

            st.download_button(
                label="📥 Descarregar Tabela em Excel",
                data=buffer.getvalue(),
                file_name=f"Extracao_{tipo_selecionado}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("A extração não encontrou dados válidos para o tipo selecionado nestes ficheiros.")
