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

# 1. Configuração das colunas de saída (padronizadas)
COLUNAS_CONFIG = {
    "hotel": ["Arquivo", "Data", "Informação adicional", "Qtde", "Unidade", "Total"],
    "exames": ["Arquivo", "Exame", "Valor"],
    "refeicoes": ["Arquivo", "Data", "Total"],
    "notas_fiscais": ["Data", "NF", "Chave_NFe", "Fornecedor", "UF", "Descricao", "CFOP", "NCM", "Valor_Itens", "BC_ICMS", "ICMS_Origem", "VR_DIFAL", "OBS"]
}

# 2. Lista de UFs brasileiras
UFS_BRASIL = [
    'AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 
    'MT', 'MS', 'MG', 'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 
    'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'
]

# 3. Mapeamento de nomes de colunas para nomes padronizados
MAPA_COLUNAS = {
    # Data
    'DATA': 'Data',
    'DT': 'Data',
    'DT.': 'Data',
    'DATA EMISSÃO': 'Data',
    'DATA EMISSAO': 'Data',
    'EMISSÃO': 'Data',
    'EMISSAO': 'Data',
    
    # NF
    'N. FISCAL': 'NF',
    'Nº N.F': 'NF',
    'N.F': 'NF',
    'NF': 'NF',
    'NÚMERO NF': 'NF',
    'NUMERO NF': 'NF',
    'NR NF': 'NF',
    'NOTA FISCAL': 'NF',
    
    # Chave NFe
    'CHAVE DA NOTA FISCAL ELETRÔNICA': 'Chave_NFe',
    'CHAVE DA NOTA FISCAL ELETRONICA': 'Chave_NFe',
    'CHAVE NFE': 'Chave_NFe',
    'CHAVE NF-E': 'Chave_NFe',
    'CHAVE DE ACESSO': 'Chave_NFe',
    'CHAVE': 'Chave_NFe',
    
    # Fornecedor
    'FORNECEDOR': 'Fornecedor',
    'EMITENTE': 'Fornecedor',
    'RAZÃO SOCIAL': 'Fornecedor',
    'RAZAO SOCIAL': 'Fornecedor',
    'NOME': 'Fornecedor',
    
    # UF
    'UF': 'UF',
    'ESTADO': 'UF',
    
    # Descrição
    'DESCRIÇÃO DA MERCADORIA': 'Descricao',
    'DESCRIÇÃO DA MERCADORIA/SERVIÇO': 'Descricao',
    'DESCRICAO DA MERCADORIA': 'Descricao',
    'DESCRICAO DA MERCADORIA/SERVICO': 'Descricao',
    'DESCRIÇÃO': 'Descricao',
    'DESCRICAO': 'Descricao',
    'MERCADORIA': 'Descricao',
    'PRODUTO': 'Descricao',
    'ITEM': 'Descricao',
    
    # CFOP
    'CFOP': 'CFOP',
    'CÓDIGO CFOP': 'CFOP',
    'CODIGO CFOP': 'CFOP',
    'CFOP CÓD': 'CFOP',
    
    # NCM
    'NCM': 'NCM',
    'NCM/SH': 'NCM',
    'CÓDIGO NCM': 'NCM',
    'CODIGO NCM': 'NCM',
    
    # Valor Itens
    'VALOR DOS ITENS': 'Valor_Itens',
    'VALOR NF.': 'Valor_Itens',
    'VALOR NF': 'Valor_Itens',
    'VALOR TOTAL': 'Valor_Itens',
    'VALOR DA NOTA': 'Valor_Itens',
    'VL. TOTAL': 'Valor_Itens',
    'VLR TOTAL': 'Valor_Itens',
    'TOTAL NF': 'Valor_Itens',
    
    # BC ICMS
    'BC ICMS': 'BC_ICMS',
    'BASE DE CÁLCULO': 'BC_ICMS',
    'BASE DE CALCULO': 'BC_ICMS',
    'BASE CÁLCULO': 'BC_ICMS',
    'BASE CALCULO': 'BC_ICMS',
    'BC': 'BC_ICMS',
    
    # ICMS Origem
    'ICMS ORIGEM': 'ICMS_Origem',
    'ICMS': 'ICMS_Origem',
    'ICMS DESTACADO': 'ICMS_Origem',
    'VL. ICMS': 'ICMS_Origem',
    'VLR ICMS': 'ICMS_Origem',
    'ICMS ORIG': 'ICMS_Origem',
    
    # VR DIFAL
    'VR DIFAL': 'VR_DIFAL',
    'DIFAL': 'VR_DIFAL',
    'ICMS DIFAL': 'VR_DIFAL',
    'DIFERENCIAL': 'VR_DIFAL',
    'DIFERENCIAL DE ALÍQUOTA': 'VR_DIFAL',
    'DIFERENCIAL DE ALIQUOTA': 'VR_DIFAL',
    'DIF. ALÍQUOTA': 'VR_DIFAL',
    'DIF. ALIQUOTA': 'VR_DIFAL',
    
    # OBS
    'OBS': 'OBS',
    'OBS.': 'OBS',
    'OBSERVAÇÃO': 'OBS',
    'OBSERVACAO': 'OBS',
    'OBSERVAÇÕES': 'OBS',
    'OBSERVACOES': 'OBS',
    'COMPLEMENTO': 'OBS',
    
    # % Interna (pode ser ignorada ou armazenada)
    '% INTERNA': 'Pct_Interna',
    'ALÍQUOTA INTERNA': 'Pct_Interna',
    'ALIQUOTA INTERNA': 'Pct_Interna',
    '% INT': 'Pct_Interna',
    'ALQ INTERNA': 'Pct_Interna',
}


def normalizar_nome_coluna(nome):
    """Converte nome de coluna para o nome padronizado"""
    nome_upper = nome.upper().strip()
    
    # Remover caracteres especiais e espaços extras
    nome_upper = re.sub(r'\s+', ' ', nome_upper)
    
    # Buscar no mapeamento
    if nome_upper in MAPA_COLUNAS:
        return MAPA_COLUNAS[nome_upper]
    
    # Busca parcial (contém)
    for chave, valor in MAPA_COLUNAS.items():
        if chave in nome_upper or nome_upper in chave:
            return valor
    
    # Se não encontrou, retorna o nome original
    return nome


def mapear_cabecalho(cabecalho):
    """
    Recebe a linha de cabeçalho e retorna um dicionário:
    {indice_coluna: nome_padronizado}
    """
    mapeamento = {}
    
    for idx, coluna in enumerate(cabecalho):
        if coluna and str(coluna).strip():
            nome_original = str(coluna).strip()
            nome_padronizado = normalizar_nome_coluna(nome_original)
            mapeamento[idx] = nome_padronizado
    
    return mapeamento


def limpar_valor(valor_str):
    """Converte string de valor para float de forma robusta"""
    if not valor_str:
        return 0.0
    
    valor_str = str(valor_str).strip()
    
    # Se já é número, retorna
    try:
        return float(valor_str)
    except ValueError:
        pass
    
    # Remove caracteres não numéricos (exceto vírgula e ponto)
    valor_limpo = re.sub(r'[^\d.,\-]', '', valor_str)
    
    if not valor_limpo:
        return 0.0
    
    # Detecta formato brasileiro (vírgula como decimal)
    if ',' in valor_limpo:
        # Remove pontos de milhar
        valor_limpo = valor_limpo.replace('.', '')
        # Substitui vírgula por ponto
        valor_limpo = valor_limpo.replace(',', '.')
    
    try:
        return float(valor_limpo)
    except ValueError:
        return 0.0


def extrair_notas_fiscais_universal(arquivo_pdf):
    """
    Extrator UNIVERSAL de notas fiscais.
    Detecta automaticamente a estrutura do cabeçalho e se adapta.
    """
    
    dados_extraidos = []
    metricas = {
        'total_paginas': 0,
        'paginas_com_tabela': 0,
        'total_linhas_extraidas': 0,
        'nf_unicas': set(),
        'colunas_detectadas': [],
        'valor_total_itens': 0.0,
        'valor_total_difal': 0.0,
        'erros': [],
        'layout_detectado': ''
    }
    
    with pdfplumber.open(arquivo_pdf) as pdf:
        metricas['total_paginas'] = len(pdf.pages)
        
        mapeamento_colunas = None  # Será definido na primeira tabela encontrada
        
        for num_pagina, pagina in enumerate(pdf.pages):
            # Extrair tabelas da página
            tabelas = pagina.extract_tables()
            
            if not tabelas:
                continue
            
            for tabela in tabelas:
                if not tabela or len(tabela) < 2:
                    continue
                
                # Se ainda não temos mapeamento, procurar cabeçalho
                if mapeamento_colunas is None:
                    for i, linha in enumerate(tabela):
                        if not linha or not any(linha):
                            continue
                        
                        # Verificar se esta linha parece um cabeçalho
                        linha_str = ' '.join([str(c) if c else '' for c in linha])
                        
                        # Verificar múltiplos padrões de cabeçalho
                        padroes_cabecalho = [
                            ['Data', 'Fornecedor', 'CFOP'],           # Layout 1
                            ['DATA', 'FORNECEDOR', 'CFOP'],           # Layout 1 (caps)
                            ['DATA', 'N. FISCAL', 'CHAVE'],           # Layout 2
                            ['DATA', 'NF', 'UF', 'CFOP'],             # Layout 3
                            ['DATA', 'CFOP', 'DESCRIÇÃO'],            # Layout 4
                            ['DATA', 'CFOP', 'DESCRICAO'],            # Layout 4 (sem acento)
                        ]
                        
                        for padrao in padroes_cabecalho:
                            if all(p in linha_str.upper() for p in [x.upper() for x in padrao]):
                                # Encontramos o cabeçalho!
                                mapeamento_colunas = mapear_cabecalho(linha)
                                metricas['colunas_detectadas'] = list(mapeamento_colunas.values())
                                
                                # Identificar layout
                                cols_detectadas = set(mapeamento_colunas.values())
                                if 'Fornecedor' in cols_detectadas and 'UF' in cols_detectadas:
                                    metricas['layout_detectado'] = 'Layout Padrão (Fornecedor + UF)'
                                elif 'UF' in cols_detectadas and 'NCM' in cols_detectadas:
                                    metricas['layout_detectado'] = 'Layout Alternativo (UF + NCM)'
                                else:
                                    metricas['layout_detectado'] = 'Layout Automático'
                                
                                break
                        
                        if mapeamento_colunas is not None:
                            cabecalho_idx = i
                            break
                    
                    if mapeamento_colunas is None:
                        continue
                else:
                    # Procurar cabeçalho novamente (pode mudar entre páginas)
                    for i, linha in enumerate(tabela):
                        if not linha or not any(linha):
                            continue
                        linha_str = ' '.join([str(c) if c else '' for c in linha])
                        if all(p in linha_str.upper() for p in ['DATA', 'CFOP']):
                            cabecalho_idx = i
                            break
                    else:
                        continue
                
                metricas['paginas_com_tabela'] += 1
                
                # Processar linhas de dados
                for linha in tabela[cabecalho_idx + 1:]:
                    if not linha or not any(linha):
                        continue
                    
                    # Limpar células
                    celulas = [str(c).strip() if c else '' for c in linha]
                    
                    # Pular totais
                    linha_completa = ' '.join(celulas).upper()
                    if any(p in linha_completa for p in ['TOTAL DO MÊS', 'TOTAIS DO MÊS', 'TOTAL GERAL', 'TOTAIS']):
                        continue
                    
                    # Construir dicionário de dados com base no mapeamento
                    item = {}
                    
                    for idx_col, nome_padronizado in mapeamento_colunas.items():
                        if idx_col < len(celulas):
                            valor_celula = celulas[idx_col]
                        else:
                            valor_celula = ''
                        
                        # Armazenar no campo padronizado
                        if nome_padronizado not in item:
                            item[nome_padronizado] = valor_celula
                    
                    # Validar campos essenciais
                    data = item.get('Data', '')
                    nf = item.get('NF', '')
                    
                    # Se não tem NF, tentar extrair de Chave_NFe (posições 25-34)
                    if (not nf or not nf.isdigit()) and item.get('Chave_NFe', ''):
                        chave = item.get('Chave_NFe', '')
                        if len(chave) >= 34:
                            nf = chave[25:34]
                            item['NF'] = nf
                    
                    # Validar data
                    if not re.match(r'\d{2}/\d{2}/\d{4}', data):
                        # Tentar encontrar data em outras colunas
                        for v in item.values():
                            if re.match(r'\d{2}/\d{2}/\d{4}', str(v)):
                                data = str(v)
                                item['Data'] = data
                                break
                        else:
                            continue
                    
                    # Validar NF
                    if not nf or not re.match(r'^\d+$', str(nf)):
                        continue
                    
                    # Converter valores numéricos
                    campos_numericos = ['Valor_Itens', 'BC_ICMS', 'ICMS_Origem', 'VR_DIFAL', 'Pct_Interna']
                    for campo in campos_numericos:
                        if campo in item:
                            item[campo] = limpar_valor(item[campo])
                    
                    # Garantir campos essenciais
                    item.setdefault('Fornecedor', '')
                    item.setdefault('UF', '')
                    item.setdefault('Descricao', '')
                    item.setdefault('CFOP', '')
                    item.setdefault('NCM', '')
                    item.setdefault('Chave_NFe', '')
                    item.setdefault('OBS', '')
                    item.setdefault('Valor_Itens', 0.0)
                    item.setdefault('BC_ICMS', 0.0)
                    item.setdefault('ICMS_Origem', 0.0)
                    item.setdefault('VR_DIFAL', 0.0)
                    
                    # Validar UF (se presente)
                    uf = str(item.get('UF', '')).strip().upper()
                    if uf and len(uf) > 2:
                        # UF pode estar misturada - extrair apenas 2 letras
                        match_uf = re.search(r'([A-Z]{2})', uf)
                        if match_uf:
                            item['UF'] = match_uf.group(1)
                        else:
                            item['UF'] = uf[:2]
                    
                    # Atualizar métricas
                    metricas['nf_unicas'].add(str(nf))
                    if len(data) >= 10:
                        mes = data[3:5] + '/' + data[6:10]
                        if not hasattr(metricas, 'meses_encontrados'):
                            metricas['meses_encontrados'] = set()
                        metricas['meses_encontrados'].add(mes)
                    
                    metricas['valor_total_itens'] += item.get('Valor_Itens', 0)
                    metricas['valor_total_difal'] += item.get('VR_DIFAL', 0)
                    metricas['total_linhas_extraidas'] += 1
                    
                    # Adicionar à lista final
                    dados_extraidos.append(item)
    
    # Finalizar métricas
    metricas['nf_unicas_count'] = len(metricas['nf_unicas'])
    if hasattr(metricas, 'meses_encontrados') and metricas.get('meses_encontrados'):
        metricas['meses_encontrados'] = sorted(list(metricas['meses_encontrados']))
        metricas['meses_count'] = len(metricas['meses_encontrados'])
    else:
        metricas['meses_encontrados'] = []
        metricas['meses_count'] = 0
    
    return dados_extraidos, metricas


# ============ INTERFACE STREAMLIT ============

st.set_page_config(page_title="Extrator de Relatórios - Apoena", page_icon="📊", layout="wide")
st.title("📊 Extrator de Relatórios - Apoena")

with st.sidebar:
    st.markdown("## ⚙️ Configurações")
    st.markdown("---")
    mostrar_metricas = st.checkbox("📈 Painel de auditoria", value=True)
    mostrar_preview = st.checkbox("👁️ Preview dos dados", value=True)
    mostrar_erros = st.checkbox("🔍 Mostrar erros de extração", value=False)
    mostrar_diagnostico = st.checkbox("🔬 Diagnóstico do layout detectado", value=True)

st.markdown("### 1. Tipo de relatório:")
tipo_selecionado = st.radio(
    "Escolha:",
    options=["notas_fiscais", "hotel", "exames", "refeicoes"],
    format_func=lambda x: {
        "notas_fiscais": "📋 Notas Fiscais (Detecção Automática de Layout)",
        "hotel": "🏨 Diárias e Consumo (Plaza Hotel)",
        "exames": "🩺 Exames Ocupacionais (Biomed)",
        "refeicoes": "🍽️ Mapa de Refeições"
    }[x],
    horizontal=True
)

st.markdown("### 2. Selecione os ficheiros PDF:")
arquivos_selecionados = st.file_uploader(
    "Arraste e solte ou clique", 
    type=['pdf'], 
    accept_multiple_files=True,
    help="Aceita qualquer layout de notas fiscais (detecção automática)"
)

if st.button("🚀 Extrair Dados e Gerar Excel", type="primary", use_container_width=True):
    if not arquivos_selecionados:
        st.warning("⚠️ Selecione pelo menos um ficheiro PDF.")
    else:
        todos_dados = []
        metricas_consolidadas = {}
        
        import time
        inicio = time.time()
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, arquivo_pdf in enumerate(arquivos_selecionados):
            nome_arquivo = arquivo_pdf.name
            status_text.text(f"Processando: {nome_arquivo} ({idx+1}/{len(arquivos_selecionados)})")
            
            if tipo_selecionado == "notas_fiscais":
                dados, metricas = extrair_notas_fiscais_universal(arquivo_pdf)
                metricas['nome_arquivo'] = nome_arquivo
                
                todos_dados.extend(dados)
                metricas_consolidadas = metricas
                
                # Mostrar diagnóstico do layout
                if mostrar_diagnostico and metricas.get('colunas_detectadas'):
                    st.info(f"""
                    **📋 Arquivo: {nome_arquivo}**
                    - **Layout detectado:** {metricas.get('layout_detectado', 'Automático')}
                    - **Colunas encontradas:** {', '.join(metricas.get('colunas_detectadas', []))}
                    - **Linhas extraídas:** {metricas.get('total_linhas_extraidas', 0)}
                    """)
            
            progress_bar.progress((idx + 1) / len(arquivos_selecionados))
        
        tempo_total = round(time.time() - inicio, 2)
        status_text.text("✅ Processamento concluído!")
        
        # ============ RESULTADOS ============
        if todos_dados:
            # Criar DataFrame com todas as colunas possíveis
            colunas_saida = COLUNAS_CONFIG['notas_fiscais']
            df = pd.DataFrame(todos_dados)
            
            # Garantir que todas as colunas de saída existam
            for col in colunas_saida:
                if col not in df.columns:
                    df[col] = ''
            
            df = df[colunas_saida]
            
            # ===== MÉTRICAS =====
            if mostrar_metricas and metricas_consolidadas:
                st.markdown("---")
                st.markdown("## 📊 Painel de Auditoria")
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("📄 Páginas", metricas_consolidadas.get('total_paginas', 'N/A'))
                with col2:
                    st.metric("🧾 NFs Únicas", metricas_consolidadas.get('nf_unicas_count', 0))
                with col3:
                    st.metric("📝 Linhas Extraídas", len(todos_dados))
                with col4:
                    st.metric("⏱️ Tempo", f"{tempo_total}s")
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("💰 Valor Total Itens", f"R$ {metricas_consolidadas.get('valor_total_itens', 0):,.2f}")
                with col2:
                    st.metric("💵 Total DIFAL", f"R$ {metricas_consolidadas.get('valor_total_difal', 0):,.2f}")
                with col3:
                    st.metric("🔍 Layout", metricas_consolidadas.get('layout_detectado', 'N/A'))
                
                if metricas_consolidadas.get('meses_encontrados'):
                    st.markdown(f"**📅 Período:** {' | '.join(metricas_consolidadas['meses_encontrados'])}")
            
            # ===== PREVIEW =====
            if mostrar_preview:
                st.markdown("---")
                st.markdown("## 👁️ Preview dos Dados Extraídos")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**Primeiros 5 registros:**")
                    st.dataframe(df.head(), use_container_width=True)
                with col2:
                    st.markdown("**Últimos 5 registros:**")
                    st.dataframe(df.tail(), use_container_width=True)
            
            # ===== DOWNLOAD =====
            st.markdown("---")
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Dados Extraídos')
                
                # Aba de métricas
                if metricas_consolidadas:
                    dados_met = [
                        ['Arquivo', metricas_consolidadas.get('nome_arquivo', 'N/A')],
                        ['Layout Detectado', metricas_consolidadas.get('layout_detectado', 'N/A')],
                        ['Colunas Encontradas', ', '.join(metricas_consolidadas.get('colunas_detectadas', []))],
                        ['Páginas', metricas_consolidadas.get('total_paginas', 0)],
                        ['Linhas Extraídas', len(todos_dados)],
                        ['NFs Únicas', metricas_consolidadas.get('nf_unicas_count', 0)],
                        ['Valor Total Itens (R$)', metricas_consolidadas.get('valor_total_itens', 0)],
                        ['Total DIFAL (R$)', metricas_consolidadas.get('valor_total_difal', 0)],
                        ['Tempo Processamento (s)', tempo_total],
                    ]
                    pd.DataFrame(dados_met, columns=['Métrica', 'Valor']).to_excel(
                        writer, index=False, sheet_name='Métricas'
                    )
            
            st.success(f"✅ {len(todos_dados)} registros extraídos com sucesso!")
            st.download_button(
                label="📥 Baixar Excel Completo",
                data=buffer.getvalue(),
                file_name=f"Extracao_NFs_{metricas_consolidadas.get('nf_unicas_count', 0)}_NFs.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.error("❌ Nenhum dado foi extraído. Verifique se o PDF contém tabelas reconhecíveis.")
