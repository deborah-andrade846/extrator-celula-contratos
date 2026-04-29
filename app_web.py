# -*- coding: utf-8 -*-
"""
Created on Mon Apr  6 11:22:01 2026

@author: deborah.goncalves

Extrator Universal de Notas Fiscais - Suporta múltiplos layouts no mesmo processamento
"""

import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from datetime import datetime

# 1. Configuração das colunas de saída (padronizadas e unificadas)
COLUNAS_SAIDA_NF = [
    "Arquivo_Origem",      # Nome do arquivo PDF de origem
    "Layout_Detectado",    # Qual layout foi identificado
    "Data", 
    "NF", 
    "Chave_NFe", 
    "Fornecedor", 
    "UF", 
    "NCM",
    "CFOP", 
    "Descricao", 
    "Valor_Itens", 
    "BC_ICMS", 
    "ICMS_Origem", 
    "Pct_Interna",
    "VR_DIFAL", 
    "OBS"
]

COLUNAS_CONFIG = {
    "hotel": ["Arquivo", "Data", "Informação adicional", "Qtde", "Unidade", "Total"],
    "exames": ["Arquivo", "Exame", "Valor"],
    "refeicoes": ["Arquivo", "Data", "Total"],
    "notas_fiscais": COLUNAS_SAIDA_NF
}

# 2. Lista de UFs brasileiras
UFS_BRASIL = [
    'AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 
    'MT', 'MS', 'MG', 'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 
    'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'
]

# 3. Mapeamento universal de nomes de colunas
MAPA_COLUNAS = {
    # Data
    'DATA': 'Data', 'DT': 'Data', 'DT.': 'Data',
    'DATA EMISSÃO': 'Data', 'DATA EMISSAO': 'Data',
    'EMISSÃO': 'Data', 'EMISSAO': 'Data',
    'DATA DA NOTA': 'Data', 'DT NOTA': 'Data',
    
    # NF
    'N. FISCAL': 'NF', 'Nº N.F': 'NF', 'N.F': 'NF', 'NF': 'NF',
    'NÚMERO NF': 'NF', 'NUMERO NF': 'NF', 'NR NF': 'NF',
    'NOTA FISCAL': 'NF', 'NÚMERO': 'NF', 'NUMERO': 'NF',
    'NR. NOTA': 'NF', 'NRO NF': 'NF',
    
    # Chave NFe
    'CHAVE DA NOTA FISCAL ELETRÔNICA': 'Chave_NFe',
    'CHAVE DA NOTA FISCAL ELETRONICA': 'Chave_NFe',
    'CHAVE NFE': 'Chave_NFe', 'CHAVE NF-E': 'Chave_NFe',
    'CHAVE DE ACESSO': 'Chave_NFe', 'CHAVE': 'Chave_NFe',
    'CHAVE ELETRÔNICA': 'Chave_NFe', 'CHAVE ELETRONICA': 'Chave_NFe',
    
    # Fornecedor
    'FORNECEDOR': 'Fornecedor', 'EMITENTE': 'Fornecedor',
    'RAZÃO SOCIAL': 'Fornecedor', 'RAZAO SOCIAL': 'Fornecedor',
    'NOME': 'Fornecedor', 'NOME EMITENTE': 'Fornecedor',
    'EMITENTE/REMETENTE': 'Fornecedor',
    
    # UF
    'UF': 'UF', 'ESTADO': 'UF', 'UF EMITENTE': 'UF',
    'UF ORIGEM': 'UF', 'ESTADO ORIGEM': 'UF',
    
    # Descrição
    'DESCRIÇÃO DA MERCADORIA': 'Descricao',
    'DESCRIÇÃO DA MERCADORIA/SERVIÇO': 'Descricao',
    'DESCRICAO DA MERCADORIA': 'Descricao',
    'DESCRICAO DA MERCADORIA/SERVICO': 'Descricao',
    'DESCRIÇÃO': 'Descricao', 'DESCRICAO': 'Descricao',
    'MERCADORIA': 'Descricao', 'PRODUTO': 'Descricao',
    'ITEM': 'Descricao', 'DISCRIMINAÇÃO': 'Descricao',
    'DISCRIMINACAO': 'Descricao',
    
    # CFOP
    'CFOP': 'CFOP', 'CÓDIGO CFOP': 'CFOP', 'CODIGO CFOP': 'CFOP',
    'CFOP CÓD': 'CFOP', 'CÓD. CFOP': 'CFOP', 'COD. CFOP': 'CFOP',
    
    # NCM
    'NCM': 'NCM', 'NCM/SH': 'NCM', 'CÓDIGO NCM': 'NCM',
    'CODIGO NCM': 'NCM', 'NCM CÓD': 'NCM',
    
    # Valor Itens
    'VALOR DOS ITENS': 'Valor_Itens', 'VALOR NF.': 'Valor_Itens',
    'VALOR NF': 'Valor_Itens', 'VALOR TOTAL': 'Valor_Itens',
    'VALOR DA NOTA': 'Valor_Itens', 'VL. TOTAL': 'Valor_Itens',
    'VLR TOTAL': 'Valor_Itens', 'TOTAL NF': 'Valor_Itens',
    'VALOR': 'Valor_Itens',
    
    # BC ICMS
    'BC ICMS': 'BC_ICMS', 'BASE DE CÁLCULO': 'BC_ICMS',
    'BASE DE CALCULO': 'BC_ICMS', 'BASE CÁLCULO': 'BC_ICMS',
    'BASE CALCULO': 'BC_ICMS', 'BC': 'BC_ICMS',
    'BASE DE CALCULO DO ICMS': 'BC_ICMS',
    
    # ICMS Origem
    'ICMS ORIGEM': 'ICMS_Origem', 'ICMS': 'ICMS_Origem',
    'ICMS DESTACADO': 'ICMS_Origem', 'VL. ICMS': 'ICMS_Origem',
    'VLR ICMS': 'ICMS_Origem', 'ICMS ORIG': 'ICMS_Origem',
    'VALOR ICMS': 'ICMS_Origem',
    
    # VR DIFAL
    'VR DIFAL': 'VR_DIFAL', 'DIFAL': 'VR_DIFAL',
    'ICMS DIFAL': 'VR_DIFAL', 'DIFERENCIAL': 'VR_DIFAL',
    'DIFERENCIAL DE ALÍQUOTA': 'VR_DIFAL',
    'DIFERENCIAL DE ALIQUOTA': 'VR_DIFAL',
    'DIF. ALÍQUOTA': 'VR_DIFAL', 'DIF. ALIQUOTA': 'VR_DIFAL',
    'VALOR DIFAL': 'VR_DIFAL',
    
    # % Interna
    '% INTERNA': 'Pct_Interna', 'ALÍQUOTA INTERNA': 'Pct_Interna',
    'ALIQUOTA INTERNA': 'Pct_Interna', '% INT': 'Pct_Interna',
    'ALQ INTERNA': 'Pct_Interna', 'ALÍQ INTERNA': 'Pct_Interna',
    'ALIQ INTERNA': 'Pct_Interna',
    
    # OBS
    'OBS': 'OBS', 'OBS.': 'OBS', 'OBSERVAÇÃO': 'OBS',
    'OBSERVACAO': 'OBS', 'OBSERVAÇÕES': 'OBS',
    'OBSERVACOES': 'OBS', 'COMPLEMENTO': 'OBS',
}


def normalizar_nome_coluna(nome):
    """Converte qualquer nome de coluna para o nome padronizado"""
    if not nome:
        return None
    
    nome_upper = str(nome).upper().strip()
    nome_upper = re.sub(r'\s+', ' ', nome_upper)
    
    # Busca exata
    if nome_upper in MAPA_COLUNAS:
        return MAPA_COLUNAS[nome_upper]
    
    # Busca por similaridade (contém)
    for chave, valor in MAPA_COLUNAS.items():
        if chave in nome_upper or nome_upper in chave:
            return valor
    
    return None


def mapear_cabecalho(cabecalho):
    """
    Recebe uma linha de cabeçalho e retorna:
    {indice_coluna: nome_padronizado}
    """
    mapeamento = {}
    
    for idx, coluna in enumerate(cabecalho):
        if coluna and str(coluna).strip():
            nome_padronizado = normalizar_nome_coluna(str(coluna).strip())
            if nome_padronizado:
                mapeamento[idx] = nome_padronizado
    
    return mapeamento


def limpar_valor(valor_str):
    """Converte string de valor para float de forma robusta"""
    if not valor_str:
        return 0.0
    
    valor_str = str(valor_str).strip()
    
    try:
        return float(valor_str)
    except ValueError:
        pass
    
    valor_limpo = re.sub(r'[^\d.,\-]', '', valor_str)
    
    if not valor_limpo:
        return 0.0
    
    if ',' in valor_limpo:
        valor_limpo = valor_limpo.replace('.', '').replace(',', '.')
    
    try:
        return float(valor_limpo)
    except ValueError:
        return 0.0


def identificar_layout(mapeamento):
    """
    Identifica qual layout foi detectado baseado nas colunas encontradas.
    Retorna uma string descritiva.
    """
    colunas = set(mapeamento.values())
    
    if 'Fornecedor' in colunas and 'UF' in colunas and 'Descricao' in colunas and 'CFOP' in colunas:
        if 'NCM' in colunas:
            return "Layout 1 - Completo (Fornecedor + UF + NCM)"
        return "Layout 1 - Padrão (Fornecedor + UF)"
    
    if 'UF' in colunas and 'NCM' in colunas and 'CFOP' in colunas and 'Descricao' in colunas:
        return "Layout 2 - Alternativo (UF + NCM + CFOP)"
    
    if 'UF' in colunas and 'Descricao' in colunas and 'CFOP' in colunas:
        return "Layout 3 - Simplificado (UF + CFOP)"
    
    return "Layout Automático"


def extrair_notas_fiscais_universal(arquivo_pdf, nome_arquivo):
    """
    Extrator UNIVERSAL de notas fiscais.
    Detecta automaticamente o layout e extrai todos os dados.
    
    Retorna: (dados_extraidos, metricas)
    """
    
    dados_extraidos = []
    metricas = {
        'nome_arquivo': nome_arquivo,
        'total_paginas': 0,
        'paginas_com_tabela': 0,
        'total_linhas_extraidas': 0,
        'nf_unicas': set(),
        'colunas_detectadas': [],
        'valor_total_itens': 0.0,
        'valor_total_difal': 0.0,
        'erros': [],
        'layout_detectado': 'Não identificado',
        'meses_encontrados': set(),
    }
    
    try:
        with pdfplumber.open(arquivo_pdf) as pdf:
            metricas['total_paginas'] = len(pdf.pages)
            
            mapeamento_colunas = None
            layout_nome = "Não identificado"
            
            for num_pagina, pagina in enumerate(pdf.pages):
                # Extrair tabelas
                tabelas = pagina.extract_tables()
                
                if not tabelas:
                    continue
                
                for tabela in tabelas:
                    if not tabela or len(tabela) < 2:
                        continue
                    
                    # Procurar cabeçalho se ainda não temos mapeamento
                    # ou se mudou de página (revalidar)
                    cabecalho_encontrado = False
                    
                    for i, linha in enumerate(tabela):
                        if not linha or not any(linha):
                            continue
                        
                        linha_str = ' '.join([str(c) if c else '' for c in linha])
                        
                        # Padrões para identificar cabeçalho
                        padroes = [
                            ['DATA', 'CFOP'],
                            ['DATA', 'FORNECEDOR'],
                            ['DATA', 'N. FISCAL'],
                            ['DATA', 'NF'],
                        ]
                        
                        for padrao in padroes:
                            if all(p in linha_str.upper() for p in [x.upper() for x in padrao]):
                                novo_mapeamento = mapear_cabecalho(linha)
                                
                                if len(novo_mapeamento) >= 5:  # Mínimo de 5 colunas
                                    mapeamento_colunas = novo_mapeamento
                                    layout_nome = identificar_layout(mapeamento_colunas)
                                    metricas['colunas_detectadas'] = list(mapeamento_colunas.values())
                                    metricas['layout_detectado'] = layout_nome
                                    cabecalho_encontrado = True
                                    cabecalho_idx = i
                                    break
                        
                        if cabecalho_encontrado:
                            break
                    
                    if not cabecalho_encontrado:
                        continue
                    
                    metricas['paginas_com_tabela'] += 1
                    
                    # Processar linhas de dados
                    for linha in tabela[cabecalho_idx + 1:]:
                        if not linha or not any(linha):
                            continue
                        
                        celulas = [str(c).strip() if c else '' for c in linha]
                        linha_completa = ' '.join(celulas).upper()
                        
                        # Pular totais e títulos
                        if any(p in linha_completa for p in [
                            'TOTAL DO MÊS', 'TOTAIS DO MÊS', 'TOTAL GERAL', 
                            'TOTAIS', 'DEMONSTRATIVO', 'PÁGINA'
                        ]):
                            continue
                        
                        # Construir item
                        item = {}
                        for idx_col, nome_padronizado in mapeamento_colunas.items():
                            if idx_col < len(celulas):
                                item[nome_padronizado] = celulas[idx_col]
                            else:
                                item[nome_padronizado] = ''
                        
                        # Tentar extrair NF da chave se necessário
                        data = item.get('Data', '')
                        nf = item.get('NF', '')
                        chave = item.get('Chave_NFe', '')
                        
                        # Se tem chave mas não tem NF, extrair da chave
                        if chave and len(chave) >= 34 and (not nf or not nf.isdigit()):
                            item['NF'] = chave[25:34]
                            nf = item['NF']
                        
                        # Validar data
                        if not re.match(r'\d{2}/\d{2}/\d{4}', str(data)):
                            # Procurar data em qualquer coluna
                            for k, v in item.items():
                                if re.match(r'\d{2}/\d{2}/\d{4}', str(v)):
                                    item['Data'] = str(v)
                                    data = str(v)
                                    break
                            else:
                                continue
                        
                        # Validar NF (deve existir e ser numérica)
                        if not nf or not re.match(r'^\d+$', str(nf)):
                            continue
                        
                        # Converter campos numéricos
                        campos_numericos = ['Valor_Itens', 'BC_ICMS', 'ICMS_Origem', 'VR_DIFAL', 'Pct_Interna']
                        for campo in campos_numericos:
                            if campo in item:
                                item[campo] = limpar_valor(item[campo])
                        
                        # Garantir campos existentes
                        for campo in COLUNAS_SAIDA_NF:
                            if campo not in item:
                                item[campo] = ''
                        
                        # Limpar UF
                        uf = str(item.get('UF', '')).strip().upper()
                        if uf and len(uf) > 2:
                            match_uf = re.search(r'([A-Z]{2})', uf)
                            if match_uf:
                                item['UF'] = match_uf.group(1)
                            else:
                                item['UF'] = uf[:2]
                        
                        # Adicionar metadados
                        item['Arquivo_Origem'] = nome_arquivo
                        item['Layout_Detectado'] = layout_nome
                        
                        # Atualizar métricas
                        metricas['nf_unicas'].add(str(nf))
                        if len(str(data)) >= 10:
                            mes = str(data)[3:5] + '/' + str(data)[6:10]
                            metricas['meses_encontrados'].add(mes)
                        
                        metricas['valor_total_itens'] += item.get('Valor_Itens', 0) or 0
                        metricas['valor_total_difal'] += item.get('VR_DIFAL', 0) or 0
                        metricas['total_linhas_extraidas'] += 1
                        
                        dados_extraidos.append(item)
        
    except Exception as e:
        metricas['erros'].append({
            'arquivo': nome_arquivo,
            'erro': str(e)[:200]
        })
    
    # Finalizar métricas
    metricas['nf_unicas_count'] = len(metricas['nf_unicas'])
    metricas['meses_encontrados'] = sorted(list(metricas['meses_encontrados']))
    metricas['meses_count'] = len(metricas['meses_encontrados'])
    
    return dados_extraidos, metricas


# ============ INTERFACE STREAMLIT ============

st.set_page_config(
    page_title="Extrator Universal de Notas Fiscais - Apoena", 
    page_icon="📊", 
    layout="wide"
)

st.title("📊 Extrator Universal de Notas Fiscais")
st.markdown("### 🧠 Detecta automaticamente o layout e consolida tudo em um único Excel")

with st.sidebar:
    st.markdown("## ⚙️ Configurações")
    st.markdown("---")
    mostrar_metricas = st.checkbox("📈 Painel de auditoria consolidado", value=True)
    mostrar_preview = st.checkbox("👁️ Preview dos dados extraídos", value=True)
    mostrar_detalhes_arquivos = st.checkbox("📋 Detalhes por arquivo", value=True)
    mostrar_erros = st.checkbox("🔍 Mostrar erros de extração", value=False)

st.markdown("### 📂 Selecione os ficheiros PDF (múltiplos layouts suportados):")
arquivos_selecionados = st.file_uploader(
    "Arraste e solte ou clique para procurar - Pode misturar layouts diferentes!",
    type=['pdf'], 
    accept_multiple_files=True,
    help="Suporta Layout 1 (Fornecedor+UF) e Layout 2 (UF+NCM+CFOP) - detecta automaticamente"
)

if arquivos_selecionados:
    st.info(f"📁 **{len(arquivos_selecionados)} arquivo(s) selecionado(s)**")
    
    # Mostrar nomes dos arquivos
    with st.expander("📋 Ver lista de arquivos"):
        for i, arq in enumerate(arquivos_selecionados, 1):
            st.markdown(f"{i}. **{arq.name}** ({arq.size:,} bytes)")

if st.button("🚀 Extrair Todos os Dados e Gerar Excel Consolidado", type="primary", use_container_width=True):
    if not arquivos_selecionados:
        st.warning("⚠️ Selecione pelo menos um ficheiro PDF.")
    else:
        todos_dados = []
        metricas_por_arquivo = []
        metricas_consolidadas = {
            'total_arquivos': len(arquivos_selecionados),
            'total_nfs_unicas': set(),
            'total_linhas': 0,
            'total_valor_itens': 0.0,
            'total_valor_difal': 0.0,
            'arquivos_com_erro': 0,
            'arquivos_sem_dados': 0,
        }
        
        import time
        inicio = time.time()
        
        # Barra de progresso
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # Container para mostrar progresso por arquivo
        progress_container = st.container()
        
        for idx, arquivo_pdf in enumerate(arquivos_selecionados):
            nome_arquivo = arquivo_pdf.name
            status_text.text(f"🔍 Processando {idx+1}/{len(arquivos_selecionados)}: {nome_arquivo}")
            
            # Extrair dados
            dados, metricas = extrair_notas_fiscais_universal(arquivo_pdf, nome_arquivo)
            
            # Acumular
            todos_dados.extend(dados)
            metricas_por_arquivo.append(metricas)
            
            # Atualizar métricas consolidadas
            metricas_consolidadas['total_nfs_unicas'].update(metricas['nf_unicas'])
            metricas_consolidadas['total_linhas'] += metricas['total_linhas_extraidas']
            metricas_consolidadas['total_valor_itens'] += metricas['valor_total_itens']
            metricas_consolidadas['total_valor_difal'] += metricas['valor_total_difal']
            
            if metricas['erros']:
                metricas_consolidadas['arquivos_com_erro'] += 1
            
            if metricas['total_linhas_extraidas'] == 0:
                metricas_consolidadas['arquivos_sem_dados'] += 1
            
            # Mostrar progresso individual
            with progress_container:
                if metricas['total_linhas_extraidas'] > 0:
                    st.success(f"✅ **{nome_arquivo}**: {metricas['total_linhas_extraidas']} linhas | Layout: {metricas['layout_detectado']}")
                else:
                    st.warning(f"⚠️ **{nome_arquivo}**: Nenhum dado extraído")
            
            progress_bar.progress((idx + 1) / len(arquivos_selecionados))
        
        tempo_total = round(time.time() - inicio, 2)
        status_text.text("✅ Processamento concluído!")
        
        # Limpar container de progresso
        progress_container.empty()
        
        # ============ RESULTADO FINAL ============
        
        if todos_dados:
            # Criar DataFrame consolidado
            df_final = pd.DataFrame(todos_dados)
            
            # Garantir todas as colunas de saída
            for col in COLUNAS_SAIDA_NF:
                if col not in df_final.columns:
                    df_final[col] = ''
            
            df_final = df_final[COLUNAS_SAIDA_NF]
            
            # Ordenar por arquivo e data
            df_final = df_final.sort_values(['Arquivo_Origem', 'Data', 'NF'])
            
            # ===== PAINEL CONSOLIDADO =====
            if mostrar_metricas:
                st.markdown("---")
                st.markdown("## 📊 Painel de Auditoria Consolidado")
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("📁 Arquivos Processados", len(arquivos_selecionados))
                with col2:
                    st.metric("🧾 Total NFs Únicas", len(metricas_consolidadas['total_nfs_unicas']))
                with col3:
                    st.metric("📝 Total Linhas Extraídas", len(todos_dados))
                with col4:
                    st.metric("⏱️ Tempo Total", f"{tempo_total}s")
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("💰 Valor Total Itens", f"R$ {metricas_consolidadas['total_valor_itens']:,.2f}")
                with col2:
                    st.metric("💵 Total DIFAL", f"R$ {metricas_consolidadas['total_valor_difal']:,.2f}")
                with col3:
                    layouts_detectados = set(m['layout_detectado'] for m in metricas_por_arquivo if m['total_linhas_extraidas'] > 0)
                    st.metric("🔍 Layouts Detectados", len(layouts_detectados))
                with col4:
                    taxa_sucesso = ((len(arquivos_selecionados) - metricas_consolidadas['arquivos_sem_dados']) / len(arquivos_selecionados)) * 100
                    st.metric("✅ Taxa de Sucesso", f"{taxa_sucesso:.0f}%")
            
            # ===== DETALHES POR ARQUIVO =====
            if mostrar_detalhes_arquivos:
                st.markdown("---")
                st.markdown("## 📋 Detalhes por Arquivo")
                
                df_detalhes = pd.DataFrame([
                    {
                        'Arquivo': m['nome_arquivo'],
                        'Layout Detectado': m['layout_detectado'],
                        'Páginas': m['total_paginas'],
                        'Linhas Extraídas': m['total_linhas_extraidas'],
                        'NFs Únicas': m['nf_unicas_count'],
                        'Valor Itens (R$)': m['valor_total_itens'],
                        'Total DIFAL (R$)': m['valor_total_difal'],
                        'Erros': len(m['erros']),
                        'Status': '✅' if m['total_linhas_extraidas'] > 0 else '❌'
                    }
                    for m in metricas_por_arquivo
                ])
                
                st.dataframe(df_detalhes, use_container_width=True)
                
                # Distribuição por layout
                st.markdown("### Distribuição por Layout")
                df_layout = df_final.groupby('Layout_Detectado').agg(
                    Linhas=('NF', 'count'),
                    NFs_Unicas=('NF', 'nunique'),
                    Valor_Itens=('Valor_Itens', 'sum'),
                    Total_DIFAL=('VR_DIFAL', 'sum')
                ).reset_index()
                
                col1, col2 = st.columns(2)
                with col1:
                    st.dataframe(df_layout, use_container_width=True)
                with col2:
                    st.bar_chart(df_layout.set_index('Layout_Detectado')['Linhas'])
            
            # ===== ERROS =====
            if mostrar_erros:
                erros_total = []
                for m in metricas_por_arquivo:
                    for e in m['erros']:
                        erros_total.append(e)
                
                if erros_total:
                    st.markdown("---")
                    st.markdown(f"## ⚠️ Erros de Extração ({len(erros_total)})")
                    st.dataframe(pd.DataFrame(erros_total), use_container_width=True)
            
            # ===== PREVIEW =====
            if mostrar_preview:
                st.markdown("---")
                st.markdown("## 👁️ Preview dos Dados Consolidados")
                
                st.markdown(f"**Total: {len(df_final)} registros de {len(arquivos_selecionados)} arquivos**")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**Primeiros 5 registros:**")
                    st.dataframe(df_final.head(), use_container_width=True)
                with col2:
                    st.markdown("**Últimos 5 registros:**")
                    st.dataframe(df_final.tail(), use_container_width=True)
                
                # Amostra por arquivo
                with st.expander("🔍 Ver distribuição por arquivo"):
                    df_amostra = df_final.groupby('Arquivo_Origem').agg(
                        Registros=('NF', 'count'),
                        NFs=('NF', 'nunique'),
                        Layout=('Layout_Detectado', 'first'),
                        Primeira_Data=('Data', 'min'),
                        Ultima_Data=('Data', 'max'),
                        Valor_Itens=('Valor_Itens', 'sum'),
                        DIFAL=('VR_DIFAL', 'sum')
                    ).reset_index()
                    st.dataframe(df_amostra, use_container_width=True)
            
            # ===== DOWNLOAD =====
            st.markdown("---")
            st.markdown("## 📥 Download do Arquivo Consolidado")
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                # Aba 1: Dados consolidados
                df_final.to_excel(writer, index=False, sheet_name='Dados Consolidados')
                
                # Aba 2: Resumo por arquivo
                if mostrar_detalhes_arquivos:
                    df_detalhes.to_excel(writer, index=False, sheet_name='Resumo por Arquivo')
                
                # Aba 3: Distribuição por layout
                if 'df_layout' in locals():
                    df_layout.to_excel(writer, index=False, sheet_name='Distribuição por Layout')
                
                # Aba 4: Métricas consolidadas
                metricas_resumo = [
                    ['Data Processamento', datetime.now().strftime("%d/%m/%Y %H:%M")],
                    ['Total Arquivos', len(arquivos_selecionados)],
                    ['Total NFs Únicas', len(metricas_consolidadas['total_nfs_unicas'])],
                    ['Total Linhas Extraídas', len(todos_dados)],
                    ['Total Valor Itens (R$)', metricas_consolidadas['total_valor_itens']],
                    ['Total DIFAL (R$)', metricas_consolidadas['total_valor_difal']],
                    ['Layouts Detectados', ', '.join(layouts_detectados)],
                    ['Tempo Processamento (s)', tempo_total],
                    ['Arquivos com Erro', metricas_consolidadas['arquivos_com_erro']],
                    ['Arquivos sem Dados', metricas_consolidadas['arquivos_sem_dados']],
                ]
                pd.DataFrame(metricas_resumo, columns=['Métrica', 'Valor']).to_excel(
                    writer, index=False, sheet_name='Métricas Consolidadas'
                )
            
            # Download
            st.success(f"""
            ### ✅ Extração Concluída com Sucesso!
            - **{len(todos_dados)}** registros extraídos
            - **{len(metricas_consolidadas['total_nfs_unicas'])}** notas fiscais únicas
            - **{len(arquivos_selecionados)}** arquivos processados
            - **{len(layouts_detectados)}** layout(s) detectado(s)
            """)
            
            st.download_button(
                label=f"📥 Baixar Excel Consolidado ({len(todos_dados)} registros)",
                data=buffer.getvalue(),
                file_name=f"Notas_Fiscais_Consolidadas_{len(metricas_consolidadas['total_nfs_unicas'])}_NFs_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        else:
            st.error("""
            ### ❌ Nenhum dado foi extraído
            
            Possíveis causas:
            1. Os PDFs não contêm tabelas reconhecíveis
            2. O formato do cabeçalho não foi identificado
            3. As datas não estão no formato DD/MM/AAAA
            
            Verifique se os arquivos estão corretos e tente novamente.
            """)
