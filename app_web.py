# -*- coding: utf-8 -*-
"""
Created on Mon Apr  6 11:22:01 2026

@author: deborah.goncalves
"""

import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# 1. Configuração direta das colunas (elimina o config.json)
COLUNAS_CONFIG = {
    "hotel": ["Arquivo", "Data", "Informação adicional", "Qtde", "Unidade", "Total"],
    "exames": ["Arquivo", "Exame", "Valor"],
    "refeicoes": ["Arquivo", "Data", "Total"],
    "notas_fiscais": ["Data", "NF", "Chave_NFe", "Fornecedor", "UF", "Descricao", "CFOP", "Valor_Itens", "ICMS_Origem", "VR_DIFAL"]
}

# 2. Função de limpeza para os relatórios do hotel
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

# 3. Função para extrair dados de notas fiscais
def extrair_notas_fiscais(texto_completo):
    """Extrai dados estruturados de demonstrativos de notas fiscais com ICMS DIFAL"""
    
    linhas = texto_completo.split('\n')
    dados_extraidos = []
    
    # Localizar o cabeçalho da tabela
    cabecalho_encontrado = False
    colunas_esperadas = ['Data', 'Nº N.F', 'Chave da Nota Fiscal Eletrônica', 'Fornecedor', 
                         'UF', 'Descrição da Mercadoria', 'CFOP', 'Valor dos Itens', 
                         'ICMS Origem', 'VR DIFAL']
    
    for linha in linhas:
        linha_limpa = linha.strip()
        
        # Verificar se é o cabeçalho
        if not cabecalho_encontrado:
            if all(col in linha_limpa for col in ['Data', 'Nº N.F', 'Fornecedor', 'UF']):
                cabecalho_encontrado = True
                continue
        
        # Se já encontrou o cabeçalho, processar as linhas
        if cabecalho_encontrado:
            # Pular linhas vazias
            if not linha_limpa:
                continue
                
            # Pular títulos e totais
            if 'DEMONSTRATIVO' in linha_limpa.upper() or 'TOTAL' in linha_limpa.upper():
                continue
                
            # Tentar extrair os dados da linha
            # Padrão regex para identificar uma linha de produto
            # Formato: Data | NF | Chave NFe | Fornecedor | UF | Descricao | CFOP | Valor Itens | ICMS Origem | VR DIFAL
            
            # Identificar se a linha começa com uma data (formato DD/MM/AAAA)
            padrao_data = r'^(\d{2}/\d{2}/\d{4})'
            if not re.match(padrao_data, linha_limpa):
                # Verificar se é continuação de descrição da linha anterior
                if dados_extraidos and not re.match(r'^\d', linha_limpa):
                    continue  # Ignorar linhas que não são novos registros
                continue
            
            # Tentar extrair usando split por espaços múltiplos
            # Primeiro, extrair a data
            match_data = re.match(r'^(\d{2}/\d{2}/\d{4})\s+', linha_limpa)
            if not match_data:
                continue
                
            data = match_data.group(1)
            resto_linha = linha_limpa[match_data.end():]
            
            # Tentar capturar o número da NF (primeiro número após a data)
            match_nf = re.match(r'(\d+)\s+', resto_linha)
            if not match_nf:
                continue
                
            nf = match_nf.group(1)
            resto_linha = resto_linha[match_nf.end():]
            
            # Capturar a chave NFe (sequência longa de números)
            match_chave = re.match(r'(\d{44})\s+', resto_linha)
            if not match_chave:
                # Pode ser continuação de linha anterior ou formato diferente
                continue
                
            chave_nfe = match_chave.group(1)
            resto_linha = resto_linha[match_chave.end():]
            
            # Extrair Fornecedor (texto até encontrar UF)
            match_fornecedor_uf = re.match(r'(.+?)\s+([A-Z]{2})\s+(.+)', resto_linha)
            if not match_fornecedor_uf:
                continue
                
            fornecedor = match_fornecedor_uf.group(1).strip()
            uf = match_fornecedor_uf.group(2)
            resto_final = match_fornecedor_uf.group(3)
            
            # O resto deve conter: Descricao CFOP Valor_Itens ICMS_Origem VR_DIFAL
            # Procurar pelo padrão: texto CFOP números números números
            match_dados = re.match(r'(.+?)\s+(\d{4})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)', resto_final)
            if not match_dados:
                continue
                
            descricao = match_dados.group(1).strip()
            cfop = match_dados.group(2)
            valor_itens = match_dados.group(3).replace('.', '').replace(',', '.')
            icms_origem = match_dados.group(4).replace('.', '').replace(',', '.')
            vr_difal = match_dados.group(5).replace('.', '').replace(',', '.')
            
            # Converter para float
            try:
                valor_itens = float(valor_itens)
                icms_origem = float(icms_origem)
                vr_difal = float(vr_difal)
            except ValueError:
                continue
            
            dados_extraidos.append({
                "Data": data,
                "NF": nf,
                "Chave_NFe": chave_nfe,
                "Fornecedor": fornecedor,
                "UF": uf,
                "Descricao": descricao,
                "CFOP": cfop,
                "Valor_Itens": valor_itens,
                "ICMS_Origem": icms_origem,
                "VR_DIFAL": vr_difal
            })
    
    return dados_extraidos

# 4. Interface Visual do Streamlit
st.set_page_config(page_title="Extrator de Relatórios - Apoena", page_icon="📊")
st.title("📊 Extrator de Relatórios - Apoena")

st.markdown("### 1. O que deseja extrair?")
tipo_selecionado = st.radio(
    "Escolha o tipo de relatório:",
    options=["hotel", "exames", "refeicoes", "notas_fiscais"],
    format_func=lambda x: {
        "hotel": "Diárias e Consumo (Plaza Hotel)",
        "exames": "Exames Ocupacionais (Biomed)",
        "refeicoes": "Mapa de Refeições",
        "notas_fiscais": "Notas Fiscais com Produtos (ICMS DIFAL)"
    }[x]
)

st.markdown("### 2. Selecione os ficheiros PDF:")
arquivos_selecionados = st.file_uploader("Arraste e solte ou clique para procurar", type=['pdf'], accept_multiple_files=True)

# 5. Botão de Extração
if st.button("Extrair Dados e Gerar Excel", type="primary"):
    if not arquivos_selecionados:
        st.warning("Por favor, selecione pelo menos um ficheiro PDF.")
    else:
        dados_finais = []
        
        with st.spinner("A processar ficheiros..."):
            for arquivo_pdf in arquivos_selecionados:
                nome_arquivo = arquivo_pdf.name
                
                try:
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
                                    # CORREÇÃO: Aceita linhas que tenham o R$ uma ou mais vezes
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
                                
                        elif tipo_selecionado == "notas_fiscais":
                            dados_notas = extrair_notas_fiscais(texto_completo)
                            dados_finais.extend(dados_notas)

                except Exception as e:
                    st.error(f"Erro ao processar o ficheiro {nome_arquivo}:\n{str(e)}")
                    st.stop()

        # 6. Gerar e oferecer o Excel para download
        if dados_finais:
            st.success(f"Extração concluída com sucesso! Foram encontradas {len(dados_finais)} linhas.")
            
            df = pd.DataFrame(dados_finais, columns=COLUNAS_CONFIG[tipo_selecionado])
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
