# -*- coding: utf-8 -*-
"""
Extrator Universal de Notas Fiscais - Versão Simplificada e Robusta
"""

import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from datetime import datetime

# ============ CONFIGURAÇÃO ============

COLUNAS_SAIDA = [
    "Arquivo", "Data", "NF", "Chave_NFe", "Fornecedor", "UF", 
    "NCM", "CFOP", "Descricao", "Valor_Itens", "BC_ICMS", 
    "ICMS_Origem", "VR_DIFAL", "OBS"
]

UFS_BRASIL = ['AC','AL','AP','AM','BA','CE','DF','ES','GO','MA','MT','MS','MG','PA','PB','PR','PE','PI','RJ','RN','RS','RO','RR','SC','SP','SE','TO']


# ============ FUNÇÃO DE EXTRAÇÃO ============

def extrair_tabelas_pdf(arquivo_pdf, nome_arquivo):
    """
    Extrai TODAS as tabelas do PDF sem depender de regex complexo.
    Usa extract_tables() do pdfplumber que é muito mais confiável.
    """
    
    todos_dados = []
    metricas = {
        'arquivo': nome_arquivo,
        'paginas': 0,
        'paginas_com_tabela': 0,
        'linhas_extraidas': 0,
        'nfs_unicas': set(),
        'valor_itens': 0.0,
        'valor_difal': 0.0,
        'erros': []
    }
    
    with pdfplumber.open(arquivo_pdf) as pdf:
        metricas['paginas'] = len(pdf.pages)
        
        for num_pag, pagina in enumerate(pdf.pages, 1):
            tabelas = pagina.extract_tables()
            
            if not tabelas:
                continue
            
            for tabela in tabelas:
                if not tabela or len(tabela) < 2:
                    continue
                
                # Encontrar linha do cabeçalho
                idx_cabecalho = -1
                colunas_mapeadas = {}
                
                for i, linha in enumerate(tabela):
                    if not linha:
                        continue
                    
                    linha_upper = ' '.join([str(c).upper() if c else '' for c in linha])
                    
                    # Verifica se é cabeçalho (contém Data + (CFOP ou Fornecedor))
                    if 'DATA' in linha_upper and ('CFOP' in linha_upper or 'FORNECEDOR' in linha_upper or 'N. FISCAL' in linha_upper):
                        idx_cabecalho = i
                        
                        # Mapear colunas
                        for j, col in enumerate(linha):
                            if col and str(col).strip():
                                col_upper = str(col).upper().strip()
                                
                                if 'DATA' in col_upper:
                                    colunas_mapeadas[j] = 'Data'
                                elif 'CHAVE' in col_upper:
                                    colunas_mapeadas[j] = 'Chave_NFe'
                                elif 'FORNECEDOR' in col_upper or 'EMITENTE' in col_upper or 'RAZÃO' in col_upper:
                                    colunas_mapeadas[j] = 'Fornecedor'
                                elif col_upper in ['UF', 'ESTADO']:
                                    colunas_mapeadas[j] = 'UF'
                                elif 'NCM' in col_upper:
                                    colunas_mapeadas[j] = 'NCM'
                                elif 'CFOP' in col_upper:
                                    colunas_mapeadas[j] = 'CFOP'
                                elif 'DESCRI' in col_upper or 'MERCADORIA' in col_upper or 'PRODUTO' in col_upper:
                                    colunas_mapeadas[j] = 'Descricao'
                                elif 'VALOR' in col_upper and 'NF' in col_upper:
                                    colunas_mapeadas[j] = 'Valor_Itens'
                                elif 'BC' in col_upper and 'ICMS' in col_upper:
                                    colunas_mapeadas[j] = 'BC_ICMS'
                                elif 'ICMS' in col_upper and 'ORIGEM' in col_upper:
                                    colunas_mapeadas[j] = 'ICMS_Origem'
                                elif col_upper == 'ICMS':
                                    colunas_mapeadas[j] = 'ICMS_Origem'
                                elif 'DIFAL' in col_upper or 'DIFERENCIAL' in col_upper:
                                    colunas_mapeadas[j] = 'VR_DIFAL'
                                elif 'OBS' in col_upper:
                                    colunas_mapeadas[j] = 'OBS'
                                elif 'N. FISCAL' in col_upper or col_upper in ['NF', 'N.F', 'Nº N.F']:
                                    colunas_mapeadas[j] = 'NF'
                                elif 'VALOR' in col_upper or 'TOTAL' in col_upper:
                                    colunas_mapeadas[j] = 'Valor_Itens'
                        
                        break
                
                if idx_cabecalho == -1:
                    continue
                
                metricas['paginas_com_tabela'] += 1
                
                # Processar linhas de dados
                for linha in tabela[idx_cabecalho + 1:]:
                    if not linha:
                        continue
                    
                    # Filtrar linhas vazias
                    valores_nao_vazios = [c for c in linha if c and str(c).strip()]
                    if len(valores_nao_vazios) < 3:
                        continue
                    
                    # Pular totais
                    linha_str = ' '.join([str(c) if c else '' for c in linha]).upper()
                    if any(p in linha_str for p in ['TOTAL', 'TOTAIS', 'PÁGINA', 'PAGINA']):
                        continue
                    
                    # Construir registro
                    registro = {
                        'Arquivo': nome_arquivo,
                        'Data': '', 'NF': '', 'Chave_NFe': '', 'Fornecedor': '', 'UF': '',
                        'NCM': '', 'CFOP': '', 'Descricao': '', 'Valor_Itens': 0.0,
                        'BC_ICMS': 0.0, 'ICMS_Origem': 0.0, 'VR_DIFAL': 0.0, 'OBS': ''
                    }
                    
                    for idx_col, nome_campo in colunas_mapeadas.items():
                        if idx_col < len(linha) and linha[idx_col]:
                            registro[nome_campo] = str(linha[idx_col]).strip()
                    
                    # Validar data
                    data = registro.get('Data', '')
                    if not re.match(r'\d{2}/\d{2}/\d{4}', data):
                        continue
                    
                    # Validar NF - se não tem NF, tentar extrair da chave
                    nf = registro.get('NF', '')
                    chave = registro.get('Chave_NFe', '')
                    
                    if (not nf or not nf.isdigit()) and chave and len(chave) >= 34:
                        nf = chave[25:34]
                        registro['NF'] = nf
                    
                    if not nf or not nf.isdigit():
                        continue
                    
                    # Limpar UF
                    uf = registro.get('UF', '').upper().strip()
                    if len(uf) > 2:
                        match = re.search(r'([A-Z]{2})', uf)
                        registro['UF'] = match.group(1) if match else uf[:2]
                    
                    # Converter valores numéricos
                    for campo in ['Valor_Itens', 'BC_ICMS', 'ICMS_Origem', 'VR_DIFAL']:
                        valor_str = registro.get(campo, '0')
                        try:
                            # Remove caracteres não numéricos (exceto , e .)
                            valor_limpo = re.sub(r'[^\d,.\-]', '', str(valor_str))
                            if ',' in valor_limpo:
                                valor_limpo = valor_limpo.replace('.', '').replace(',', '.')
                            registro[campo] = float(valor_limpo) if valor_limpo else 0.0
                        except:
                            registro[campo] = 0.0
                    
                    # Atualizar métricas
                    metricas['nfs_unicas'].add(nf)
                    metricas['linhas_extraidas'] += 1
                    metricas['valor_itens'] += registro['Valor_Itens']
                    metricas['valor_difal'] += registro['VR_DIFAL']
                    
                    todos_dados.append(registro)
    
    metricas['nfs_count'] = len(metricas['nfs_unicas'])
    return todos_dados, metricas


# ============ INTERFACE STREAMLIT ============

st.set_page_config(page_title="Extrator NF - Apoena", page_icon="📊", layout="wide")
st.title("📊 Extrator de Notas Fiscais - Consolidado")
st.markdown("### 🧠 Detecta automaticamente os dois layouts e gera um único Excel")

# Sidebar
with st.sidebar:
    st.markdown("## ⚙️ Opções")
    mostrar_detalhes = st.checkbox("📋 Mostrar detalhes por arquivo", value=True)
    mostrar_preview = st.checkbox("👁️ Mostrar preview dos dados", value=True)

# Upload
st.markdown("### 📂 Selecione os arquivos PDF:")
arquivos = st.file_uploader(
    "Arraste os PDFs aqui (misture Layout 1 e Layout 2)",
    type=['pdf'],
    accept_multiple_files=True
)

if st.button("🚀 Extrair e Gerar Excel Consolidado", type="primary", use_container_width=True):
    if not arquivos:
        st.warning("⚠️ Selecione pelo menos um arquivo PDF.")
    else:
        todos_dados = []
        todas_metricas = []
        total_nfs = set()
        
        progress = st.progress(0)
        status = st.empty()
        
        for i, arq in enumerate(arquivos):
            status.text(f"Processando {i+1}/{len(arquivos)}: {arq.name}")
            
            dados, met = extrair_tabelas_pdf(arq, arq.name)
            
            todos_dados.extend(dados)
            todas_metricas.append(met)
            
            if dados:
                total_nfs.update(met['nfs_unicas'])
            
            progress.progress((i+1)/len(arquivos))
        
        status.text("✅ Concluído!")
        
        if todos_dados:
            # DataFrame final
            df = pd.DataFrame(todos_dados, columns=COLUNAS_SAIDA)
            
            # Métricas consolidadas
            total_linhas = len(df)
            total_valor_itens = df['Valor_Itens'].sum()
            total_difal = df['VR_DIFAL'].sum()
            
            # ===== PAINEL DE MÉTRICAS =====
            st.markdown("---")
            st.markdown("## 📊 Resumo da Extração")
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("📁 Arquivos", len(arquivos))
            c2.metric("🧾 NFs Únicas", len(total_nfs))
            c3.metric("📝 Linhas Extraídas", total_linhas)
            c4.metric("⏱️ Tempo", "Concluído")
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("💰 Valor Itens", f"R$ {total_valor_itens:,.2f}")
            c2.metric("💵 Total DIFAL", f"R$ {total_difal:,.2f}")
            c3.metric("📊 Média/NF", f"R$ {total_difal/len(total_nfs):,.2f}" if total_nfs else "R$ 0,00")
            c4.metric("✅ Sucesso", f"{(len([m for m in todas_metricas if m['linhas_extraidas'] > 0])/len(arquivos))*100:.0f}%")
            
            # ===== DETALHES POR ARQUIVO =====
            if mostrar_detalhes:
                st.markdown("---")
                st.markdown("## 📋 Detalhes por Arquivo")
                
                df_det = pd.DataFrame([
                    {
                        'Arquivo': m['arquivo'],
                        'Páginas': m['paginas'],
                        'Linhas': m['linhas_extraidas'],
                        'NFs Únicas': m['nfs_count'],
                        'Valor Itens (R$)': f"{m['valor_itens']:,.2f}",
                        'DIFAL (R$)': f"{m['valor_difal']:,.2f}",
                        'Status': '✅' if m['linhas_extraidas'] > 0 else '❌'
                    }
                    for m in todas_metricas
                ])
                
                st.dataframe(df_det, use_container_width=True)
            
            # ===== PREVIEW =====
            if mostrar_preview:
                st.markdown("---")
                st.markdown("## 👁️ Preview dos Dados")
                
                st.markdown(f"**{total_linhas} registros extraídos**")
                
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("*Primeiros 5:*")
                    st.dataframe(df.head(), use_container_width=True)
                with c2:
                    st.markdown("*Últimos 5:*")
                    st.dataframe(df.tail(), use_container_width=True)
            
            # ===== DOWNLOAD =====
            st.markdown("---")
            st.markdown("## 📥 Download")
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                # Aba principal
                df.to_excel(writer, index=False, sheet_name='Notas Fiscais')
                
                # Aba de resumo
                df_resumo = pd.DataFrame([
                    ['Data Extração', datetime.now().strftime("%d/%m/%Y %H:%M")],
                    ['Arquivos Processados', len(arquivos)],
                    ['NFs Únicas', len(total_nfs)],
                    ['Total Linhas', total_linhas],
                    ['Valor Total Itens', f"R$ {total_valor_itens:,.2f}"],
                    ['Total DIFAL', f"R$ {total_difal:,.2f}"],
                ], columns=['Métrica', 'Valor'])
                df_resumo.to_excel(writer, index=False, sheet_name='Resumo')
            
            st.success(f"✅ **{total_linhas}** registros de **{len(total_nfs)}** NFs extraídos com sucesso!")
            
            st.download_button(
                label=f"📥 Baixar Excel ({total_linhas} linhas)",
                data=buffer.getvalue(),
                file_name=f"NFs_Consolidadas_{len(total_nfs)}_NFs.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.error("❌ Nenhum dado foi extraído. Verifique se os PDFs contêm tabelas.")
