# -*- coding: utf-8 -*-
"""
Extrator Universal de Notas Fiscais - Apoena
Compatível com múltiplos layouts de PDF
"""

import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import time

# Configuração inicial da página DEVE ser a primeira chamada Streamlit
st.set_page_config(
    page_title="Extrator de Notas Fiscais - Apoena", 
    page_icon="📊", 
    layout="wide"
)

# ============ CONFIGURAÇÕES ============

COLUNAS_SAIDA = [
    "Arquivo_Origem", "Layout_Detectado", "Data", "NF", "Chave_NFe",
    "Fornecedor", "UF", "NCM", "Descricao", "CFOP",
    "Valor_Itens", "BC_ICMS", "ICMS_Origem", "VR_DIFAL",
    "Pct_Interna", "OBS", "Pagina"
]

UFS_BRASIL = [
    'AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA',
    'MT', 'MS', 'MG', 'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN',
    'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'
]

# Mapeamento de nomes de colunas para padronizado
MAPA_COLUNAS = {
    'DATA': 'Data', 'DT': 'Data', 'DT.': 'Data',
    'DATA EMISSÃO': 'Data', 'DATA EMISSAO': 'Data',
    'EMISSÃO': 'Data', 'EMISSAO': 'Data',
    
    'N. FISCAL': 'NF', 'Nº N.F': 'NF', 'N.F': 'NF', 'NF': 'NF',
    'NÚMERO NF': 'NF', 'NUMERO NF': 'NF', 'NR NF': 'NF',
    'NOTA FISCAL': 'NF', 'N FISCAL': 'NF',
    
    'CHAVE DA NOTA FISCAL ELETRÔNICA': 'Chave_NFe',
    'CHAVE DA NOTA FISCAL ELETRONICA': 'Chave_NFe',
    'CHAVE NFE': 'Chave_NFe', 'CHAVE NF-E': 'Chave_NFe',
    'CHAVE DE ACESSO': 'Chave_NFe', 'CHAVE': 'Chave_NFe',
    
    'FORNECEDOR': 'Fornecedor', 'EMITENTE': 'Fornecedor',
    'RAZÃO SOCIAL': 'Fornecedor', 'RAZAO SOCIAL': 'Fornecedor',
    'NOME': 'Fornecedor',
    
    'UF': 'UF', 'ESTADO': 'UF',
    
    'DESCRIÇÃO DA MERCADORIA': 'Descricao',
    'DESCRIÇÃO DA MERCADORIA/SERVIÇO': 'Descricao',
    'DESCRICAO DA MERCADORIA': 'Descricao',
    'DESCRICAO DA MERCADORIA/SERVICO': 'Descricao',
    'DESCRIÇÃO': 'Descricao', 'DESCRICAO': 'Descricao',
    'MERCADORIA': 'Descricao', 'PRODUTO': 'Descricao', 'ITEM': 'Descricao',
    
    'CFOP': 'CFOP', 'CÓDIGO CFOP': 'CFOP', 'CODIGO CFOP': 'CFOP',
    
    'NCM': 'NCM', 'NCM/SH': 'NCM', 'CÓDIGO NCM': 'NCM', 'CODIGO NCM': 'NCM',
    
    'VALOR DOS ITENS': 'Valor_Itens', 'VALOR NF.': 'Valor_Itens',
    'VALOR NF': 'Valor_Itens', 'VALOR TOTAL': 'Valor_Itens',
    'VALOR DA NOTA': 'Valor_Itens', 'VL. TOTAL': 'Valor_Itens',
    'VLR TOTAL': 'Valor_Itens', 'TOTAL NF': 'Valor_Itens',
    
    'BC ICMS': 'BC_ICMS', 'BASE DE CÁLCULO': 'BC_ICMS',
    'BASE DE CALCULO': 'BC_ICMS', 'BC': 'BC_ICMS',
    
    'ICMS ORIGEM': 'ICMS_Origem', 'ICMS': 'ICMS_Origem',
    'ICMS DESTACADO': 'ICMS_Origem', 'VL. ICMS': 'ICMS_Origem',
    'VLR ICMS': 'ICMS_Origem',
    
    'VR DIFAL': 'VR_DIFAL', 'DIFAL': 'VR_DIFAL',
    'ICMS DIFAL': 'VR_DIFAL', 'DIFERENCIAL': 'VR_DIFAL',
    'DIFERENCIAL DE ALÍQUOTA': 'VR_DIFAL',
    
    'OBS': 'OBS', 'OBS.': 'OBS', 'OBSERVAÇÃO': 'OBS',
    'OBSERVACAO': 'OBS', 'COMPLEMENTO': 'OBS',
    
    '% INTERNA': 'Pct_Interna', 'ALÍQUOTA INTERNA': 'Pct_Interna',
    'ALIQUOTA INTERNA': 'Pct_Interna', '% INT': 'Pct_Interna',
}


# ============ FUNÇÕES ============

def normalizar_nome_coluna(nome):
    """Converte nome de coluna para o nome padronizado"""
    if not nome:
        return ""
    nome_upper = str(nome).upper().strip()
    nome_upper = re.sub(r'\s+', ' ', nome_upper)
    
    if nome_upper in MAPA_COLUNAS:
        return MAPA_COLUNAS[nome_upper]
    
    for chave, valor in MAPA_COLUNAS.items():
        if chave in nome_upper or nome_upper in chave:
            return valor
    return nome


def limpar_valor(valor_str):
    """Converte string de valor para float"""
    if not valor_str:
        return 0.0
    try:
        valor_str = str(valor_str).strip()
        return float(valor_str)
    except ValueError:
        pass
    
    valor_limpo = re.sub(r'[^\d.,\-]', '', str(valor_str))
    if not valor_limpo:
        return 0.0
    
    if ',' in valor_limpo:
        valor_limpo = valor_limpo.replace('.', '').replace(',', '.')
    
    try:
        return float(valor_limpo)
    except ValueError:
        return 0.0


def extrair_dados_tabela(arquivo_pdf, nome_arquivo):
    """
    Extrai dados de notas fiscais de forma universal.
    Retorna (dados_extraidos, metricas)
    """
    dados = []
    metricas = {
        'nome_arquivo': nome_arquivo,
        'total_paginas': 0,
        'paginas_com_dados': 0,
        'total_linhas': 0,
        'nf_unicas': set(),
        'layouts_encontrados': set(),
        'valor_total': 0.0,
        'difal_total': 0.0,
        'erros': 0
    }
    
    try:
        with pdfplumber.open(arquivo_pdf) as pdf:
            metricas['total_paginas'] = len(pdf.pages)
            
            for num_pagina, pagina in enumerate(pdf.pages):
                # Método 1: Tentar extrair tabelas
                tabelas = pagina.extract_tables()
                
                processou_pagina = False
                
                if tabelas:
                    for tabela in tabelas:
                        if not tabela or len(tabela) < 2:
                            continue
                        
                        # Procurar cabeçalho
                        cabecalho_idx = -1
                        mapeamento = {}
                        layout = ""
                        
                        for i, linha in enumerate(tabela):
                            if not linha or not any(linha):
                                continue
                            
                            linha_str = ' '.join([str(c) if c else '' for c in linha]).upper()
                            
                            # Verificar padrões de cabeçalho
                            if any([
                                'DATA' in linha_str and 'FORNECEDOR' in linha_str and 'CFOP' in linha_str,
                                'DATA' in linha_str and 'NCM' in linha_str and 'CFOP' in linha_str,
                                'DATA' in linha_str and 'CHAVE' in linha_str and 'CFOP' in linha_str,
                                'DATA' in linha_str and 'N. FISCAL' in linha_str,
                            ]):
                                cabecalho_idx = i
                                
                                # Mapear colunas
                                for j, col in enumerate(linha):
                                    if col and str(col).strip():
                                        mapeamento[j] = normalizar_nome_coluna(str(col))
                                
                                # Identificar layout
                                cols = list(mapeamento.values())
                                if 'Fornecedor' in cols and 'UF' in cols:
                                    layout = 'Layout 1 - Fornecedor+UF'
                                elif 'NCM' in cols and 'BC_ICMS' in cols:
                                    layout = 'Layout 2 - UF+NCM+BC ICMS'
                                elif 'NCM' in cols:
                                    layout = 'Layout 2 - UF+NCM'
                                else:
                                    layout = 'Layout Automático'
                                
                                metricas['layouts_encontrados'].add(layout)
                                break
                        
                        if cabecalho_idx < 0:
                            continue
                        
                        # Processar linhas
                        for linha in tabela[cabecalho_idx + 1:]:
                            if not linha or not any(linha):
                                continue
                            
                            celulas = [str(c).strip() if c else '' for c in linha]
                            
                            # Pular totais
                            linha_upper = ' '.join(celulas).upper()
                            if any(p in linha_upper for p in ['TOTAL DO MÊS', 'TOTAIS DO MÊS', 'TOTAL GERAL', 'TOTAIS']):
                                continue
                            
                            # Construir registro
                            registro = {
                                'Arquivo_Origem': nome_arquivo,
                                'Layout_Detectado': layout,
                                'Pagina': num_pagina + 1,
                                'Fornecedor': '', 'UF': '', 'NCM': '', 'Descricao': '',
                                'CFOP': '', 'Chave_NFe': '', 'OBS': '',
                                'Valor_Itens': 0.0, 'BC_ICMS': 0.0,
                                'ICMS_Origem': 0.0, 'VR_DIFAL': 0.0, 'Pct_Interna': 0.0
                            }
                            
                            for idx_col, nome_campo in mapeamento.items():
                                if idx_col < len(celulas) and celulas[idx_col]:
                                    registro[nome_campo] = celulas[idx_col]
                            
                            # Extrair NF da chave se necessário
                            nf = str(registro.get('NF', ''))
                            if (not nf or not nf.isdigit()) and registro.get('Chave_NFe'):
                                chave = str(registro['Chave_NFe'])
                                if len(chave) >= 34:
                                    nf = chave[25:34]
                                    registro['NF'] = nf
                            
                            # Validar
                            data = str(registro.get('Data', ''))
                            if not re.match(r'\d{2}/\d{2}/\d{4}', data):
                                continue
                            
                            if not nf or not re.match(r'^\d+$', str(nf)):
                                continue
                            
                            # Converter valores
                            for campo in ['Valor_Itens', 'BC_ICMS', 'ICMS_Origem', 'VR_DIFAL', 'Pct_Interna']:
                                registro[campo] = limpar_valor(registro.get(campo, 0))
                            
                            # Limpar UF
                            uf = str(registro.get('UF', '')).strip().upper()
                            if uf and len(uf) > 2:
                                match = re.search(r'([A-Z]{2})', uf)
                                registro['UF'] = match.group(1) if match else uf[:2]
                            
                            # Atualizar métricas
                            metricas['nf_unicas'].add(nf)
                            metricas['valor_total'] += registro['Valor_Itens']
                            metricas['difal_total'] += registro['VR_DIFAL']
                            metricas['total_linhas'] += 1
                            
                            dados.append(registro)
                            processou_pagina = True
                
                if processou_pagina:
                    metricas['paginas_com_dados'] += 1
        
        metricas['nf_unicas_count'] = len(metricas['nf_unicas'])
        metricas['layouts_encontrados'] = list(metricas['layouts_encontrados'])
        
    except Exception as e:
        metricas['erros'] = str(e)[:200]
        st.warning(f"⚠️ Aviso ao processar {nome_arquivo}: {str(e)[:150]}")
    
    return dados, metricas


# ============ INTERFACE ============

st.title("📊 Extrator Universal de Notas Fiscais")
st.markdown("### Consolida automaticamente múltiplos layouts em um único Excel")

# Sidebar
with st.sidebar:
    st.markdown("## ⚙️ Opções")
    mostrar_metricas = st.checkbox("📈 Painel consolidado", value=True)
    mostrar_preview = st.checkbox("👁️ Preview dos dados", value=True)
    mostrar_detalhes = st.checkbox("📋 Detalhes por arquivo", value=True)

# Upload
st.markdown("### 📁 Selecione os arquivos PDF")
st.caption("Formatos aceitos: Layout 1 (Fornecedor+UF) e Layout 2 (UF+NCM+BC ICMS)")

arquivos = st.file_uploader(
    "Arraste os arquivos ou clique para selecionar",
    type=['pdf'],
    accept_multiple_files=True
)

if arquivos:
    st.info(f"📂 **{len(arquivos)} arquivo(s)** selecionado(s)")

# Botão de extração
if st.button("🚀 Extrair e Consolidar", type="primary", use_container_width=True):
    if not arquivos:
        st.warning("⚠️ Selecione pelo menos um arquivo PDF.")
    else:
        todos_dados = []
        todas_metricas = []
        
        # Totais gerais
        total_nfs = set()
        total_valor = 0.0
        total_difal = 0.0
        erros_count = 0
        
        inicio = time.time()
        progresso = st.progress(0)
        status = st.empty()
        
        for i, arq in enumerate(arquivos):
            status.text(f"Processando {i+1}/{len(arquivos)}: {arq.name}")
            
            dados, met = extrair_dados_tabela(arq, arq.name)
            
            if dados:
                todos_dados.extend(dados)
                todas_metricas.append(met)
                total_nfs.update(met['nf_unicas'])
                total_valor += met['valor_total']
                total_difal += met['difal_total']
            else:
                erros_count += 1
            
            progresso.progress((i + 1) / len(arquivos))
        
        tempo_total = round(time.time() - inicio, 2)
        status.text("✅ Concluído!")
        
        # ============ RESULTADOS ============
        if todos_dados:
            st.markdown("---")
            st.success(f"✅ Extração concluída! **{len(todos_dados)} registros** em **{tempo_total}s**")
            
            # DataFrame
            df = pd.DataFrame(todos_dados)
            for col in COLUNAS_SAIDA:
                if col not in df.columns:
                    df[col] = ''
            df = df[COLUNAS_SAIDA]
            
            # ===== MÉTRICAS =====
            if mostrar_metricas:
                st.markdown("## 📊 Resumo Consolidado")
                
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("📁 Arquivos", f"{len(arquivos) - erros_count}/{len(arquivos)}")
                c2.metric("📝 Registros", len(todos_dados))
                c3.metric("🧾 NFs Únicas", len(total_nfs))
                c4.metric("⏱️ Tempo", f"{tempo_total}s")
                
                c1, c2 = st.columns(2)
                c1.metric("💰 Valor Total Itens", f"R$ {total_valor:,.2f}")
                c2.metric("💵 Total DIFAL", f"R$ {total_difal:,.2f}")
                
                # Layouts encontrados
                layouts = set()
                for m in todas_metricas:
                    layouts.update(m.get('layouts_encontrados', []))
                if layouts:
                    st.info(f"🔍 **Layouts detectados:** {', '.join(layouts)}")
            
            # ===== DETALHES POR ARQUIVO =====
            if mostrar_detalhes and todas_metricas:
                st.markdown("## 📋 Detalhes por Arquivo")
                
                resumo = []
                for m in todas_metricas:
                    resumo.append({
                        'Arquivo': m['nome_arquivo'],
                        'Páginas': m['total_paginas'],
                        'Linhas': m['total_linhas'],
                        'NFs Únicas': m['nf_unicas_count'],
                        'Valor Itens': f"R$ {m['valor_total']:,.2f}",
                        'DIFAL': f"R$ {m['difal_total']:,.2f}",
                        'Layout': ', '.join(m.get('layouts_encontrados', []))
                    })
                
                df_resumo = pd.DataFrame(resumo)
                st.dataframe(df_resumo, use_container_width=True)
            
            # ===== PREVIEW =====
            if mostrar_preview:
                st.markdown("## 👁️ Preview dos Dados")
                
                tab1, tab2 = st.tabs(["Primeiros 10", "Últimos 10"])
                with tab1:
                    st.dataframe(df.head(10), use_container_width=True)
                with tab2:
                    st.dataframe(df.tail(10), use_container_width=True)
            
            # ===== DOWNLOAD =====
            st.markdown("---")
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                # Aba 1: Dados
                df.to_excel(writer, index=False, sheet_name='Consolidado')
                
                # Aba 2: Métricas
                metricas_df = pd.DataFrame([
                    ['Arquivos Processados', len(arquivos)],
                    ['Registros Extraídos', len(todos_dados)],
                    ['NFs Únicas', len(total_nfs)],
                    ['Valor Total Itens', f"R$ {total_valor:,.2f}"],
                    ['Total DIFAL', f"R$ {total_difal:,.2f}"],
                    ['Tempo', f"{tempo_total}s"],
                    ['Data Extração', time.strftime('%d/%m/%Y %H:%M')],
                ], columns=['Métrica', 'Valor'])
                metricas_df.to_excel(writer, index=False, sheet_name='Métricas')
                
                # Aba 3: Resumo por arquivo
                if resumo:
                    pd.DataFrame(resumo).to_excel(writer, index=False, sheet_name='Por Arquivo')
            
            st.download_button(
                label=f"📥 Baixar Excel ({len(todos_dados)} registros)",
                data=buffer.getvalue(),
                file_name=f"NFs_Consolidadas_{len(total_nfs)}_NFs.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        else:
            st.error("❌ Nenhum dado extraído.")
            if erros_count > 0:
                st.warning(f"⚠️ {erros_count} arquivo(s) sem dados detectados.")
