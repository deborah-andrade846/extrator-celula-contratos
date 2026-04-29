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

# 1. Configuração das colunas de saída (padronizadas e unificadas)
COLUNAS_CONFIG = {
    "hotel": ["Arquivo", "Data", "Informação adicional", "Qtde", "Unidade", "Total"],
    "exames": ["Arquivo", "Exame", "Valor"],
    "refeicoes": ["Arquivo", "Data", "Total"],
    "notas_fiscais": ["Arquivo_Origem", "Layout_Detectado", "Data", "NF", "Chave_NFe", 
                      "Fornecedor", "UF", "NCM", "Descricao", "CFOP", 
                      "Valor_Itens", "BC_ICMS", "ICMS_Origem", "VR_DIFAL", 
                      "Pct_Interna", "OBS", "Pagina"]
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
    'N FISCAL': 'NF',
    
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
    
    # % Interna
    '% INTERNA': 'Pct_Interna',
    'ALÍQUOTA INTERNA': 'Pct_Interna',
    'ALIQUOTA INTERNA': 'Pct_Interna',
    '% INT': 'Pct_Interna',
    'ALQ INTERNA': 'Pct_Interna',
}


def normalizar_nome_coluna(nome):
    """Converte nome de coluna para o nome padronizado"""
    nome_upper = nome.upper().strip()
    nome_upper = re.sub(r'\s+', ' ', nome_upper)
    
    # Busca exata
    if nome_upper in MAPA_COLUNAS:
        return MAPA_COLUNAS[nome_upper]
    
    # Busca parcial (contém)
    for chave, valor in MAPA_COLUNAS.items():
        if chave in nome_upper or nome_upper in chave:
            return valor
    
    return nome


def mapear_cabecalho(cabecalho):
    """Mapeia índices das colunas para nomes padronizados"""
    mapeamento = {}
    for idx, coluna in enumerate(cabecalho):
        if coluna and str(coluna).strip():
            nome_original = str(coluna).strip()
            nome_padronizado = normalizar_nome_coluna(nome_original)
            mapeamento[idx] = nome_padronizado
    return mapeamento


def limpar_valor(valor_str):
    """Converte string de valor para float"""
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


def identificar_layout(cabecalho_linha):
    """
    Identifica qual layout está sendo usado baseado nas colunas do cabeçalho.
    Retorna: 'layout_fornecedor_uf', 'layout_uf_ncm', ou 'layout_desconhecido'
    """
    linha_str = ' '.join([str(c) if c else '' for c in cabecalho_linha]).upper()
    colunas_str = [str(c).upper().strip() if c else '' for c in cabecalho_linha]
    
    # Layout 1: Data | NF | Chave | Fornecedor | UF | Descricao | CFOP | Valor | ICMS | DIFAL
    if any('FORNECEDOR' in c for c in colunas_str):
        return 'Layout 1 - Fornecedor + UF'
    
    # Layout 2: DATA | N. FISCAL | Chave | UF | NCM | CFOP | Descricao | VALOR NF | BC ICMS | ICMS | % INTERNA | DIFAL | OBS
    if any('NCM' in c for c in colunas_str) and any('BC ICMS' in c for c in colunas_str):
        return 'Layout 2 - UF + NCM + BC ICMS'
    
    # Layout 2 alternativo
    if '% INTERNA' in linha_str or 'ALIQUOTA INTERNA' in linha_str:
        return 'Layout 2 - UF + NCM + % Interna'
    
    return 'Layout Automático'


def extrair_notas_fiscais_universal(arquivo_pdf, nome_arquivo):
    """
    Extrator UNIVERSAL - processa QUALQUER layout de notas fiscais
    e consolida em um único formato de saída.
    """
    
    dados_extraidos = []
    metricas = {
        'nome_arquivo': nome_arquivo,
        'total_paginas': 0,
        'paginas_com_tabela': 0,
        'total_linhas_extraidas': 0,
        'nf_unicas': set(),
        'layouts_detectados': set(),
        'valor_total_itens': 0.0,
        'valor_total_difal': 0.0,
        'erros': [],
        'colunas_detectadas_por_layout': {}
    }
    
    with pdfplumber.open(arquivo_pdf) as pdf:
        metricas['total_paginas'] = len(pdf.pages)
        
        for num_pagina, pagina in enumerate(pdf.pages):
            tabelas = pagina.extract_tables()
            
            if not tabelas:
                continue
            
            for tabela in tabelas:
                if not tabela or len(tabela) < 2:
                    continue
                
                cabecalho_idx = -1
                layout_detectado = None
                mapeamento_colunas = None
                
                # Procurar cabeçalho em cada tabela (pode haver múltiplas por página)
                for i, linha in enumerate(tabela):
                    if not linha or not any(linha):
                        continue
                    
                    linha_str = ' '.join([str(c) if c else '' for c in linha]).upper()
                    
                    # Verificar se é cabeçalho (múltiplos padrões)
                    padroes = [
                        ['DATA', 'FORNECEDOR', 'CFOP'],
                        ['DATA', 'N. FISCAL', 'CHAVE'],
                        ['DATA', 'NF', 'UF', 'CFOP'],
                        ['DATA', 'CFOP', 'DESCRI'],
                        ['DATA', 'NCM', 'CFOP'],
                    ]
                    
                    for padrao in padroes:
                        if all(p in linha_str for p in padrao):
                            cabecalho_idx = i
                            mapeamento_colunas = mapear_cabecalho(linha)
                            layout_detectado = identificar_layout(linha)
                            metricas['layouts_detectados'].add(layout_detectado)
                            
                            # Registrar colunas por layout
                            if layout_detectado not in metricas['colunas_detectadas_por_layout']:
                                metricas['colunas_detectadas_por_layout'][layout_detectado] = list(mapeamento_colunas.values())
                            break
                    
                    if cabecalho_idx >= 0:
                        break
                
                if cabecalho_idx < 0:
                    continue
                
                metricas['paginas_com_tabela'] += 1
                
                # Processar linhas de dados
                for linha in tabela[cabecalho_idx + 1:]:
                    if not linha or not any(linha):
                        continue
                    
                    celulas = [str(c).strip() if c else '' for c in linha]
                    
                    # Pular totais
                    linha_completa = ' '.join(celulas).upper()
                    if any(p in linha_completa for p in ['TOTAL DO MÊS', 'TOTAIS DO MÊS', 'TOTAL GERAL', 'TOTAIS']):
                        continue
                    
                    # Construir item com base no mapeamento
                    item = {
                        'Arquivo_Origem': nome_arquivo,
                        'Layout_Detectado': layout_detectado,
                        'Pagina': num_pagina + 1
                    }
                    
                    for idx_col, nome_padronizado in mapeamento_colunas.items():
                        if idx_col < len(celulas):
                            item[nome_padronizado] = celulas[idx_col]
                    
                    # Extrair NF da chave se necessário
                    if (not item.get('NF') or not str(item.get('NF', '')).isdigit()) and item.get('Chave_NFe', ''):
                        chave = str(item.get('Chave_NFe', ''))
                        if len(chave) >= 34:
                            item['NF'] = chave[25:34]
                    
                    # Validar data
                    data = str(item.get('Data', ''))
                    if not re.match(r'\d{2}/\d{2}/\d{4}', data):
                        for v in item.values():
                            if re.match(r'\d{2}/\d{2}/\d{4}', str(v)):
                                item['Data'] = str(v)
                                break
                        else:
                            continue
                    
                    # Validar NF
                    nf = str(item.get('NF', ''))
                    if not nf or not re.match(r'^\d+$', nf):
                        continue
                    
                    # Converter valores numéricos
                    campos_numericos = ['Valor_Itens', 'BC_ICMS', 'ICMS_Origem', 'VR_DIFAL', 'Pct_Interna']
                    for campo in campos_numericos:
                        if campo in item:
                            item[campo] = limpar_valor(item[campo])
                    
                    # Garantir todos os campos
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
                    item.setdefault('Pct_Interna', 0.0)
                    
                    # Limpar UF
                    uf = str(item.get('UF', '')).strip().upper()
                    if uf and len(uf) > 2:
                        match_uf = re.search(r'([A-Z]{2})', uf)
                        item['UF'] = match_uf.group(1) if match_uf else uf[:2]
                    
                    # Limpar fornecedor
                    if item.get('Fornecedor'):
                        item['Fornecedor'] = re.sub(r'\s+', ' ', str(item['Fornecedor'])).strip()
                    
                    # Limpar descrição
                    if item.get('Descricao'):
                        item['Descricao'] = re.sub(r'\s+', ' ', str(item['Descricao'])).strip()
                    
                    # Atualizar métricas
                    metricas['nf_unicas'].add(nf)
                    metricas['valor_total_itens'] += item.get('Valor_Itens', 0)
                    metricas['valor_total_difal'] += item.get('VR_DIFAL', 0)
                    metricas['total_linhas_extraidas'] += 1
                    
                    dados_extraidos.append(item)
    
    metricas['nf_unicas_count'] = len(metricas['nf_unicas'])
    metricas['layouts_detectados'] = list(metricas['layouts_detectados'])
    
    return dados_extraidos, metricas


# ============ INTERFACE STREAMLIT ============

st.set_page_config(page_title="Extrator de Notas Fiscais - Apoena", page_icon="📊", layout="wide")
st.title("📊 Extrator Universal de Notas Fiscais")
st.markdown("### Consolida automaticamente múltiplos layouts em um único arquivo")

with st.sidebar:
    st.markdown("## ⚙️ Configurações")
    st.markdown("---")
    mostrar_metricas = st.checkbox("📈 Painel de auditoria consolidado", value=True)
    mostrar_preview = st.checkbox("👁️ Preview dos dados", value=True)
    mostrar_diagnostico = st.checkbox("🔬 Diagnóstico por arquivo", value=True)

st.markdown("### 📁 Selecione os ficheiros PDF:")
st.caption("Arraste PDFs de qualquer layout - o sistema detecta automaticamente e consolida tudo.")

arquivos_selecionados = st.file_uploader(
    "Formatos aceitos: Layout 1 (Fornecedor+UF) e Layout 2 (UF+NCM+BC ICMS)",
    type=['pdf'], 
    accept_multiple_files=True,
    help="Processa simultaneamente PDFs de diferentes layouts"
)

if arquivos_selecionados:
    st.info(f"📂 **{len(arquivos_selecionados)} arquivo(s) selecionado(s)**")
    
    # Mostrar nomes dos arquivos
    with st.expander("📋 Ver arquivos selecionados"):
        for i, arq in enumerate(arquivos_selecionados):
            st.write(f"{i+1}. {arq.name} ({arq.size/1024:.1f} KB)")

if st.button("🚀 Extrair e Consolidar Tudo em Excel Único", type="primary", use_container_width=True):
    if not arquivos_selecionados:
        st.warning("⚠️ Selecione pelo menos um ficheiro PDF.")
    else:
        todos_dados = []
        metricas_por_arquivo = []
        total_geral = {
            'total_arquivos': len(arquivos_selecionados),
            'total_linhas': 0,
            'total_nfs_unicas': set(),
            'total_valor_itens': 0.0,
            'total_difal': 0.0,
            'layouts_encontrados': set(),
            'arquivos_com_erro': 0
        }
        
        import time
        inicio = time.time()
        
        # Barra de progresso principal
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # Container para diagnósticos
        diag_container = st.container()
        
        for idx, arquivo_pdf in enumerate(arquivos_selecionados):
            nome_arquivo = arquivo_pdf.name
            status_text.text(f"🔄 Processando {idx+1}/{len(arquivos_selecionados)}: {nome_arquivo}")
            
            try:
                dados, metricas = extrair_notas_fiscais_universal(arquivo_pdf, nome_arquivo)
                
                todos_dados.extend(dados)
                metricas_por_arquivo.append(metricas)
                
                # Atualizar totais gerais
                total_geral['total_linhas'] += metricas['total_linhas_extraidas']
                total_geral['total_nfs_unicas'].update(metricas['nf_unicas'])
                total_geral['total_valor_itens'] += metricas['valor_total_itens']
                total_geral['total_difal'] += metricas['valor_total_difal']
                total_geral['layouts_encontrados'].update(metricas['layouts_detectados'])
                
                # Diagnóstico individual
                if mostrar_diagnostico:
                    with diag_container:
                        with st.expander(f"📄 {nome_arquivo}", expanded=(len(arquivos_selecionados) <= 3)):
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Linhas extraídas", metricas['total_linhas_extraidas'])
                            with col2:
                                st.metric("NFs únicas", metricas['nf_unicas_count'])
                            with col3:
                                st.metric("Layout detectado", metricas['layouts_detectados'][0] if metricas['layouts_detectados'] else 'N/A')
                            
                            if metricas.get('colunas_detectadas_por_layout'):
                                for layout, colunas in metricas['colunas_detectadas_por_layout'].items():
                                    st.caption(f"**Colunas ({layout}):** {', '.join(colunas)}")
                
            except Exception as e:
                total_geral['arquivos_com_erro'] += 1
                st.error(f"❌ Erro ao processar {nome_arquivo}: {str(e)}")
            
            progress_bar.progress((idx + 1) / len(arquivos_selecionados))
        
        tempo_total = round(time.time() - inicio, 2)
        status_text.text("✅ Processamento concluído!")
        
        # ============ RESULTADO FINAL CONSOLIDADO ============
        if todos_dados:
            st.markdown("---")
            st.markdown("## ✅ Extração Concluída - Consolidado Único")
            
            # Criar DataFrame unificado
            colunas_saida = COLUNAS_CONFIG['notas_fiscais']
            df = pd.DataFrame(todos_dados)
            
            # Garantir colunas
            for col in colunas_saida:
                if col not in df.columns:
                    df[col] = ''
            
            df = df[colunas_saida]
            
            # ===== PAINEL CONSOLIDADO =====
            if mostrar_metricas:
                st.markdown("### 📊 Painel Consolidado")
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("📁 Arquivos Processados", 
                             f"{total_geral['total_arquivos'] - total_geral['arquivos_com_erro']}/{total_geral['total_arquivos']}")
                with col2:
                    st.metric("📝 Total de Linhas", total_geral['total_linhas'])
                with col3:
                    st.metric("🧾 Total NFs Únicas", len(total_geral['total_nfs_unicas']))
                with col4:
                    st.metric("⏱️ Tempo Total", f"{tempo_total}s")
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("💰 Valor Total Itens", f"R$ {total_geral['total_valor_itens']:,.2f}")
                with col2:
                    st.metric("💵 Total DIFAL", f"R$ {total_geral['total_difal']:,.2f}")
                with col3:
                    layouts_str = ', '.join(total_geral['layouts_encontrados'])
                    st.metric("🔍 Layouts Detectados", layouts_str if layouts_str else 'N/A')
                
                # Distribuição por layout
                if 'Layout_Detectado' in df.columns and len(df['Layout_Detectado'].unique()) > 1:
                    st.markdown("#### Distribuição por Layout")
                    layout_counts = df['Layout_Detectado'].value_counts()
                    
                    cols = st.columns(len(layout_counts))
                    for i, (layout, count) in enumerate(layout_counts.items()):
                        with cols[i]:
                            st.metric(layout, f"{count} linhas")
                
                # Distribuição por arquivo
                st.markdown("#### Distribuição por Arquivo")
                df_por_arquivo = df.groupby('Arquivo_Origem').agg(
                    Linhas=('NF', 'count'),
                    NFs_Unicas=('NF', 'nunique'),
                    Valor_Itens=('Valor_Itens', 'sum'),
                    DIFAL=('VR_DIFAL', 'sum')
                ).reset_index()
                
                st.dataframe(df_por_arquivo, use_container_width=True)
            
            # ===== PREVIEW =====
            if mostrar_preview:
                st.markdown("---")
                st.markdown("### 👁️ Preview do Consolidado")
                
                tab1, tab2, tab3 = st.tabs(["📋 Primeiros Registros", "📋 Últimos Registros", "🔍 Por Layout"])
                
                with tab1:
                    st.dataframe(df.head(10), use_container_width=True)
                
                with tab2:
                    st.dataframe(df.tail(10), use_container_width=True)
                
                with tab3:
                    if 'Layout_Detectado' in df.columns and len(df['Layout_Detectado'].unique()) > 1:
                        layout_selecionado = st.selectbox("Selecione o layout:", df['Layout_Detectado'].unique())
                        df_layout = df[df['Layout_Detectado'] == layout_selecionado]
                        st.dataframe(df_layout.head(10), use_container_width=True)
                        st.caption(f"Total de linhas neste layout: {len(df_layout)}")
            
            # ===== DOWNLOAD CONSOLIDADO =====
            st.markdown("---")
            st.markdown("### 📥 Download do Arquivo Consolidado")
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                # Aba principal: todos os dados
                df.to_excel(writer, index=False, sheet_name='Consolidado')
                
                # Aba de métricas consolidadas
                dados_metricas = [
                    ['Total de Arquivos Processados', total_geral['total_arquivos']],
                    ['Arquivos com Erro', total_geral['arquivos_com_erro']],
                    ['Total de Linhas Extraídas', total_geral['total_linhas']],
                    ['Total de NFs Únicas', len(total_geral['total_nfs_unicas'])],
                    ['Valor Total Itens (R$)', f"{total_geral['total_valor_itens']:,.2f}"],
                    ['Total DIFAL (R$)', f"{total_geral['total_difal']:,.2f}"],
                    ['Layouts Detectados', ', '.join(total_geral['layouts_encontrados'])],
                    ['Tempo de Processamento', f"{tempo_total}s"],
                    ['Data/Hora Extração', time.strftime('%d/%m/%Y %H:%M:%S')],
                ]
                pd.DataFrame(dados_metricas, columns=['Métrica', 'Valor']).to_excel(
                    writer, index=False, sheet_name='Métricas Consolidadas'
                )
                
                # Aba de resumo por arquivo
                if df_por_arquivo is not None:
                    df_por_arquivo.to_excel(writer, index=False, sheet_name='Resumo por Arquivo')
                
                # Aba de métricas detalhadas por arquivo
                if metricas_por_arquivo:
                    detalhes = []
                    for m in metricas_por_arquivo:
                        detalhes.append({
                            'Arquivo': m.get('nome_arquivo', 'N/A'),
                            'Páginas': m.get('total_paginas', 0),
                            'Linhas Extraídas': m.get('total_linhas_extraidas', 0),
                            'NFs Únicas': m.get('nf_unicas_count', 0),
                            'Valor Itens (R$)': m.get('valor_total_itens', 0),
                            'DIFAL (R$)': m.get('valor_total_difal', 0),
                            'Layout': ', '.join(m.get('layouts_detectados', [])),
                            'Colunas': ', '.join(list(m.get('colunas_detectadas_por_layout', {}).keys())[:1]) if m.get('colunas_detectadas_por_layout') else 'N/A'
                        })
                    pd.DataFrame(detalhes).to_excel(writer, index=False, sheet_name='Detalhes por Arquivo')
            
            # Botão de download
            nfs_count = len(total_geral['total_nfs_unicas'])
            st.success(f"✅ **{total_geral['total_linhas']} registros** de **{nfs_count} NFs** extraídos e consolidados!")
            
            st.download_button(
                label=f"📥 Baixar Excel Consolidado ({total_geral['total_linhas']} linhas, {nfs_count} NFs)",
                data=buffer.getvalue(),
                file_name=f"NFs_Consolidadas_{nfs_count}_NFs_{time.strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            # Informações adicionais
            with st.expander("📋 Detalhes do Arquivo Gerado"):
                st.markdown(f"""
                **Arquivo Excel contém 4 abas:**
                1. **Consolidado** - Todos os dados unificados ({total_geral['total_linhas']} linhas)
                2. **Métricas Consolidadas** - Resumo geral da extração
                3. **Resumo por Arquivo** - Totais por arquivo de origem
                4. **Detalhes por Arquivo** - Métricas detalhadas de cada PDF
                
                **Colunas no consolidado:**
                - `Arquivo_Origem` - Nome do PDF de origem
                - `Layout_Detectado` - Qual layout foi identificado
                - `Data`, `NF`, `Chave_NFe`, `Fornecedor`, `UF`, `NCM`
                - `Descricao`, `CFOP`, `Valor_Itens`, `BC_ICMS`
                - `ICMS_Origem`, `VR_DIFAL`, `Pct_Interna`, `OBS`
                - `Pagina` - Número da página no PDF
                """)
        
        else:
            st.error("❌ Nenhum dado foi extraído de nenhum arquivo.")
            if total_geral['arquivos_com_erro'] > 0:
                st.warning(f"⚠️ {total_geral['arquivos_com_erro']} arquivo(s) com erro. Verifique as mensagens acima.")
