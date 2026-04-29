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
    metricas = {
        'total_linhas_arquivo': len(linhas),
        'linhas_processadas': 0,
        'linhas_extraidas': 0,
        'linhas_ignoradas': 0,
        'total_paginas': 1,
        'nf_unicas': set(),
        'meses_encontrados': set(),
        'valor_total_itens': 0.0,
        'valor_total_difal': 0.0
    }
    
    # Localizar o cabeçalho da tabela
    cabecalho_encontrado = False
    colunas_esperadas = ['Data', 'Nº N.F', 'Chave da Nota Fiscal Eletrônica', 'Fornecedor', 
                         'UF', 'Descrição da Mercadoria', 'CFOP', 'Valor dos Itens', 
                         'ICMS Origem', 'VR DIFAL']
    
    for i, linha in enumerate(linhas):
        linha_limpa = linha.strip()
        
        # Verificar se é o cabeçalho
        if not cabecalho_encontrado:
            if all(col in linha_limpa for col in ['Data', 'Nº N.F', 'Fornecedor', 'UF']):
                cabecalho_encontrado = True
                continue
        
        # Se já encontrou o cabeçalho, processar as linhas
        if cabecalho_encontrado:
            metricas['linhas_processadas'] += 1
            
            # Pular linhas vazias
            if not linha_limpa:
                metricas['linhas_ignoradas'] += 1
                continue
                
            # Pular títulos e totais
            if 'DEMONSTRATIVO' in linha_limpa.upper() or 'TOTAIS DO MÊS' in linha_limpa.upper() or 'Total Geral' in linha_limpa.upper():
                metricas['linhas_ignoradas'] += 1
                continue
                
            # Tentar extrair os dados da linha
            # Identificar se a linha começa com uma data (formato DD/MM/AAAA)
            padrao_data = r'^(\d{2}/\d{2}/\d{4})'
            if not re.match(padrao_data, linha_limpa):
                metricas['linhas_ignoradas'] += 1
                continue
            
            # Tentar extrair usando split por espaços múltiplos
            match_data = re.match(r'^(\d{2}/\d{2}/\d{4})\s+', linha_limpa)
            if not match_data:
                metricas['linhas_ignoradas'] += 1
                continue
                
            data = match_data.group(1)
            
            # Extrair mês para métricas
            mes = data[3:5] + '/' + data[6:10]
            metricas['meses_encontrados'].add(mes)
            
            resto_linha = linha_limpa[match_data.end():]
            
            # Tentar capturar o número da NF (primeiro número após a data)
            match_nf = re.match(r'(\d+)\s+', resto_linha)
            if not match_nf:
                metricas['linhas_ignoradas'] += 1
                continue
                
            nf = match_nf.group(1)
            metricas['nf_unicas'].add(nf)
            resto_linha = resto_linha[match_nf.end():]
            
            # Capturar a chave NFe (sequência longa de números)
            match_chave = re.match(r'(\d{44})\s+', resto_linha)
            if not match_chave:
                metricas['linhas_ignoradas'] += 1
                continue
                
            chave_nfe = match_chave.group(1)
            resto_linha = resto_linha[match_chave.end():]
            
            # Extrair Fornecedor (texto até encontrar UF)
            match_fornecedor_uf = re.match(r'(.+?)\s+([A-Z]{2})\s+(.+)', resto_linha)
            if not match_fornecedor_uf:
                metricas['linhas_ignoradas'] += 1
                continue
                
            fornecedor = match_fornecedor_uf.group(1).strip()
            uf = match_fornecedor_uf.group(2)
            resto_final = match_fornecedor_uf.group(3)
            
            # O resto deve conter: Descricao CFOP Valor_Itens ICMS_Origem VR_DIFAL
            match_dados = re.match(r'(.+?)\s+(\d{4})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)', resto_final)
            if not match_dados:
                metricas['linhas_ignoradas'] += 1
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
                metricas['linhas_ignoradas'] += 1
                continue
            
            metricas['valor_total_itens'] += valor_itens
            metricas['valor_total_difal'] += vr_difal
            metricas['linhas_extraidas'] += 1
            
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
    
    # Converter set para contagem
    metricas['nf_unicas_count'] = len(metricas['nf_unicas'])
    metricas['meses_count'] = len(metricas['meses_encontrados'])
    metricas['meses_encontrados'] = sorted(list(metricas['meses_encontrados']))
    
    return dados_extraidos, metricas

# 4. Função para validar totais do PDF
def validar_totais_pdf(texto_completo, metricas, tipo):
    """Compara os totais extraídos com os declarados no PDF"""
    validacao = {}
    
    if tipo == "notas_fiscais":
        # Procurar linha de total no PDF
        linhas = texto_completo.split('\n')
        total_itens_pdf = None
        total_difal_pdf = None
        
        for linha in linhas:
            if 'Total Geral' in linha or 'TOTAIS' in linha:
                # Procurar valores na próxima linha
                continue
            if 'Valor Original do Tributo' in linha or 'Total do Crédito Tributário' in linha:
                # Tentar extrair valores
                numeros = re.findall(r'[\d.]+,\d{2}', linha)
                if len(numeros) >= 2:
                    try:
                        total_difal_pdf = float(numeros[-1].replace('.', '').replace(',', '.'))
                    except:
                        pass
        
        if total_difal_pdf:
            validacao['VR_DIFAL_Extraido'] = metricas['valor_total_difal']
            validacao['VR_DIFAL_PDF'] = total_difal_pdf
            validacao['Diferenca'] = abs(metricas['valor_total_difal'] - total_difal_pdf)
            validacao['Status'] = '✓ OK' if validacao['Diferenca'] < 0.01 else '✗ Divergência'
    
    return validacao

# 5. Interface Visual do Streamlit
st.set_page_config(page_title="Extrator de Relatórios - Apoena", page_icon="📊", layout="wide")
st.title("📊 Extrator de Relatórios - Apoena")

# Sidebar para configurações
with st.sidebar:
    st.markdown("## ⚙️ Configurações")
    st.markdown("---")
    
    # Checkbox para ativar métricas detalhadas
    mostrar_metricas = st.checkbox("📈 Mostrar painel de auditoria", value=True)
    mostrar_preview = st.checkbox("👁️ Mostrar preview dos dados", value=True)
    mostrar_validacao = st.checkbox("✅ Validar totais com PDF", value=True)

st.markdown("### 1. O que deseja extrair?")
tipo_selecionado = st.radio(
    "Escolha o tipo de relatório:",
    options=["hotel", "exames", "refeicoes", "notas_fiscais"],
    format_func=lambda x: {
        "hotel": "🏨 Diárias e Consumo (Plaza Hotel)",
        "exames": "🩺 Exames Ocupacionais (Biomed)",
        "refeicoes": "🍽️ Mapa de Refeições",
        "notas_fiscais": "📋 Notas Fiscais com Produtos (ICMS DIFAL)"
    }[x],
    horizontal=True
)

st.markdown("### 2. Selecione os ficheiros PDF:")
arquivos_selecionados = st.file_uploader(
    "Arraste e solte ou clique para procurar", 
    type=['pdf'], 
    accept_multiple_files=True,
    help="Selecione um ou mais arquivos PDF para processar"
)

# 6. Botão de Extração
col1, col2, col3 = st.columns([2, 1, 2])
with col2:
    botao_extrair = st.button("🚀 Extrair Dados e Gerar Excel", type="primary", use_container_width=True)

if botao_extrair:
    if not arquivos_selecionados:
        st.warning("⚠️ Por favor, selecione pelo menos um ficheiro PDF.")
    else:
        dados_finais = []
        metricas_gerais = {
            'arquivos_processados': 0,
            'arquivos_com_erro': 0,
            'total_linhas_extraidas': 0,
            'tempo_processamento': 0
        }
        
        import time
        inicio = time.time()
        
        # Criar containers para organizar a saída
        progress_container = st.container()
        metrics_container = st.container()
        preview_container = st.container()
        download_container = st.container()
        
        with progress_container:
            progress_bar = st.progress(0)
            status_text = st.empty()
        
        for idx, arquivo_pdf in enumerate(arquivos_selecionados):
            nome_arquivo = arquivo_pdf.name
            status_text.text(f"Processando: {nome_arquivo} ({idx+1}/{len(arquivos_selecionados)})")
            
            try:
                with pdfplumber.open(arquivo_pdf) as pdf:
                    texto_completo = "\n".join([pagina.extract_text() or "" for pagina in pdf.pages])
                    num_paginas = len(pdf.pages)
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
                        metricas_gerais['total_linhas_extraidas'] += len(dados_finais)

                    elif tipo_selecionado == "exames":
                        linhas_exames = 0
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
                                        linhas_exames += 1
                        metricas_gerais['total_linhas_extraidas'] += linhas_exames

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
                            metricas_gerais['total_linhas_extraidas'] += 1
                            
                    elif tipo_selecionado == "notas_fiscais":
                        dados_notas, metricas = extrair_notas_fiscais(texto_completo)
                        metricas['num_paginas'] = num_paginas
                        metricas['nome_arquivo'] = nome_arquivo
                        metricas['texto_completo'] = texto_completo
                        dados_finais.extend(dados_notas)
                        metricas_gerais['metricas_notas'] = metricas
                        metricas_gerais['total_linhas_extraidas'] += len(dados_notas)
                    
                    metricas_gerais['arquivos_processados'] += 1
                    
            except Exception as e:
                metricas_gerais['arquivos_com_erro'] += 1
                st.error(f"❌ Erro ao processar {nome_arquivo}:\n{str(e)}")
            
            # Atualizar barra de progresso
            progress_bar.progress((idx + 1) / len(arquivos_selecionados))
        
        fim = time.time()
        metricas_gerais['tempo_processamento'] = round(fim - inicio, 2)
        status_text.text("✅ Processamento concluído!")
        
        # 7. Painel de Métricas e Auditoria
        if mostrar_metricas and dados_finais:
            with metrics_container:
                st.markdown("---")
                st.markdown("## 📊 Painel de Auditoria da Extração")
                
                # Métricas gerais em cards
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("📁 Arquivos Processados", metricas_gerais['arquivos_processados'])
                with col2:
                    st.metric("📝 Total de Registros Extraídos", metricas_gerais['total_linhas_extraidas'])
                with col3:
                    st.metric("⏱️ Tempo de Processamento", f"{metricas_gerais['tempo_processamento']}s")
                with col4:
                    taxa_sucesso = (metricas_gerais['arquivos_processados'] - metricas_gerais['arquivos_com_erro']) / len(arquivos_selecionados) * 100
                    st.metric("✅ Taxa de Sucesso", f"{taxa_sucesso:.0f}%")
                
                # Métricas específicas para notas fiscais
                if tipo_selecionado == "notas_fiscais" and 'metricas_notas' in metricas_gerais:
                    m = metricas_gerais['metricas_notas']
                    
                    st.markdown("### 📋 Detalhes da Extração - Notas Fiscais")
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("📄 Páginas Processadas", m['num_paginas'])
                    with col2:
                        st.metric("🧾 Notas Fiscais Únicas", m['nf_unicas_count'])
                    with col3:
                        st.metric("📅 Meses Encontrados", m['meses_count'])
                    with col4:
                        st.metric("📊 Linhas de Produto", m['linhas_extraidas'])
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("💰 Valor Total Itens (R$)", f"{m['valor_total_itens']:,.2f}")
                    with col2:
                        st.metric("💵 Total DIFAL (R$)", f"{m['valor_total_difal']:,.2f}")
                    with col3:
                        st.metric("📈 Média DIFAL/Item (R$)", f"{m['valor_total_difal']/m['linhas_extraidas']:,.2f}" if m['linhas_extraidas'] > 0 else "0,00")
                    
                    # Distribuição por mês
                    if m['meses_encontrados']:
                        st.markdown("**🗓️ Período encontrado:** " + " | ".join(m['meses_encontrados']))
                    
                    # Taxa de aproveitamento
                    if m['linhas_processadas'] > 0:
                        taxa_aproveitamento = (m['linhas_extraidas'] / m['linhas_processadas']) * 100
                        st.progress(taxa_aproveitamento / 100, text=f"Taxa de Aproveitamento: {taxa_aproveitamento:.1f}% ({m['linhas_extraidas']} de {m['linhas_processadas']} linhas)")
                    
                    # Validação com totais do PDF
                    if mostrar_validacao and m.get('texto_completo'):
                        st.markdown("### 🔍 Validação com Totais do PDF")
                        
                        # Extrair totais do PDF para comparação
                        texto = m['texto_completo']
                        totais_pdf = {}
                        
                        # Procurar linha de total geral no PDF
                        linhas_pdf = texto.split('\n')
                        for linha in linhas_pdf:
                            if 'Total Geral' in linha:
                                # Próxima linha geralmente contém os valores
                                idx_linha = linhas_pdf.index(linha)
                                if idx_linha + 1 < len(linhas_pdf):
                                    prox_linha = linhas_pdf[idx_linha + 1]
                                    numeros = re.findall(r'[\d.]+,\d{2}', prox_linha)
                                    if numeros:
                                        totais_pdf['itens'] = float(numeros[0].replace('.', '').replace(',', '.')) if len(numeros) > 0 else None
                                        totais_pdf['difal'] = float(numeros[-1].replace('.', '').replace(',', '.')) if len(numeros) > 1 else None
                        
                        if totais_pdf:
                            col1, col2 = st.columns(2)
                            with col1:
                                if totais_pdf.get('itens'):
                                    dif_itens = abs(m['valor_total_itens'] - totais_pdf['itens'])
                                    status = "✅" if dif_itens < 0.01 else "⚠️"
                                    st.metric(
                                        f"{status} Valor Total Itens",
                                        f"Extraído: R$ {m['valor_total_itens']:,.2f}",
                                        f"PDF: R$ {totais_pdf['itens']:,.2f} | Dif: R$ {dif_itens:,.2f}"
                                    )
                            
                            with col2:
                                if totais_pdf.get('difal'):
                                    dif_difal = abs(m['valor_total_difal'] - totais_pdf['difal'])
                                    status = "✅" if dif_difal < 0.01 else "⚠️"
                                    st.metric(
                                        f"{status} Total DIFAL",
                                        f"Extraído: R$ {m['valor_total_difal']:,.2f}",
                                        f"PDF: R$ {totais_pdf['difal']:,.2f} | Dif: R$ {dif_difal:,.2f}"
                                    )
                        else:
                            st.info("ℹ️ Não foi possível localizar os totais no PDF para validação automática.")
        
        # 8. Preview dos dados extraídos
        if mostrar_preview and dados_finais:
            with preview_container:
                st.markdown("---")
                st.markdown("## 👁️ Preview dos Dados Extraídos")
                
                df = pd.DataFrame(dados_finais, columns=COLUNAS_CONFIG[tipo_selecionado])
                
                # Mostrar primeiras e últimas linhas
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**Primeiros 5 registros:**")
                    st.dataframe(df.head(), use_container_width=True)
                with col2:
                    st.markdown("**Últimos 5 registros:**")
                    st.dataframe(df.tail(), use_container_width=True)
                
                # Estatísticas rápidas
                with st.expander("📊 Ver estatísticas dos dados"):
                    st.dataframe(df.describe(), use_container_width=True)
        
        # 9. Download
        if dados_finais:
            with download_container:
                st.markdown("---")
                
                df = pd.DataFrame(dados_finais, columns=COLUNAS_CONFIG[tipo_selecionado])
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Dados Extraídos')
                    
                    # Adicionar aba de métricas se disponível
                    if mostrar_metricas and tipo_selecionado == "notas_fiscais" and 'metricas_notas' in metricas_gerais:
                        m = metricas_gerais['metricas_notas']
                        df_metricas = pd.DataFrame([
                            ['Arquivo', m.get('nome_arquivo', 'N/A')],
                            ['Páginas', m.get('num_paginas', 0)],
                            ['Linhas Extraídas', m.get('linhas_extraidas', 0)],
                            ['Linhas Processadas', m.get('linhas_processadas', 0)],
                            ['Linhas Ignoradas', m.get('linhas_ignoradas', 0)],
                            ['Notas Fiscais Únicas', m.get('nf_unicas_count', 0)],
                            ['Meses Encontrados', m.get('meses_count', 0)],
                            ['Valor Total Itens (R$)', f"{m.get('valor_total_itens', 0):,.2f}"],
                            ['Total DIFAL (R$)', f"{m.get('valor_total_difal', 0):,.2f}"],
                            ['Taxa Aproveitamento', f"{(m.get('linhas_extraidas', 0)/m.get('linhas_processadas', 1))*100:.1f}%"],
                            ['Tempo Processamento', f"{metricas_gerais.get('tempo_processamento', 0)}s"]
                        ], columns=['Métrica', 'Valor'])
                        df_metricas.to_excel(writer, index=False, sheet_name='Métricas de Extração')
                
                st.success(f"✅ Extração concluída! {len(dados_finais)} registros processados.")
                
                st.download_button(
                    label="📥 Descarregar Tabela em Excel",
                    data=buffer.getvalue(),
                    file_name=f"Extracao_{tipo_selecionado}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.info("ℹ️ A extração não encontrou dados válidos para o tipo selecionado nestes ficheiros.")
