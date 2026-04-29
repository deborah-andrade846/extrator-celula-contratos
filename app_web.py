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

# 1. Configuração direta das colunas
COLUNAS_CONFIG = {
    "hotel": ["Arquivo", "Data", "Informação adicional", "Qtde", "Unidade", "Total"],
    "exames": ["Arquivo", "Exame", "Valor"],
    "refeicoes": ["Arquivo", "Data", "Total"],
    "notas_fiscais": ["Data", "NF", "Chave_NFe", "Fornecedor", "UF", "Descricao", "CFOP", "Valor_Itens", "ICMS_Origem", "VR_DIFAL"]
}

# 2. Lista de UFs brasileiras
UFS_BRASIL = [
    'AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 
    'MT', 'MS', 'MG', 'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 
    'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'
]


def extrair_notas_fiscais_pdfplumber(arquivo_pdf):
    """
    Extrai dados usando a funcionalidade de tabelas do pdfplumber.
    Esta é a abordagem mais confiável para PDFs com estrutura tabular.
    """
    
    dados_extraidos = []
    metricas = {
        'total_paginas': 0,
        'paginas_com_tabela': 0,
        'total_linhas_extraidas': 0,
        'nf_unicas': set(),
        'meses_encontrados': set(),
        'valor_total_itens': 0.0,
        'valor_total_difal': 0.0,
        'erros': []
    }
    
    with pdfplumber.open(arquivo_pdf) as pdf:
        metricas['total_paginas'] = len(pdf.pages)
        
        for num_pagina, pagina in enumerate(pdf.pages):
            # Extrair tabelas da página
            tabelas = pagina.extract_tables()
            
            if not tabelas:
                continue
            
            for tabela in tabelas:
                if not tabela or len(tabela) < 2:
                    continue
                
                # Encontrar linha do cabeçalho
                cabecalho_idx = -1
                for i, linha in enumerate(tabela):
                    if not linha or not any(linha):
                        continue
                    # Verificar se é cabeçalho
                    linha_str = ' '.join([str(c) if c else '' for c in linha])
                    if all(palavra in linha_str for palavra in ['Data', 'Fornecedor', 'CFOP']):
                        cabecalho_idx = i
                        break
                
                if cabecalho_idx == -1:
                    continue
                
                metricas['paginas_com_tabela'] += 1
                
                # Processar linhas após o cabeçalho
                for linha in tabela[cabecalho_idx + 1:]:
                    if not linha or not any(linha):
                        continue
                    
                    # Limpar células
                    celulas = [str(c).strip() if c else '' for c in linha]
                    
                    # Pular totais
                    linha_completa = ' '.join(celulas)
                    if any(p in linha_completa.upper() for p in ['TOTAL DO MÊS', 'TOTAIS DO MÊS', 'TOTAL GERAL']):
                        continue
                    
                    # Verificar se tem dados mínimos (pelo menos 8 células preenchidas)
                    celulas_preenchidas = [c for c in celulas if c]
                    if len(celulas_preenchidas) < 8:
                        continue
                    
                    try:
                        # Mapear colunas (assumindo ordem padrão)
                        data = celulas[0] if len(celulas) > 0 else ''
                        nf = celulas[1] if len(celulas) > 1 else ''
                        chave_nfe = celulas[2] if len(celulas) > 2 else ''
                        fornecedor = celulas[3] if len(celulas) > 3 else ''
                        uf = celulas[4] if len(celulas) > 4 else ''
                        descricao = celulas[5] if len(celulas) > 5 else ''
                        cfop = celulas[6] if len(celulas) > 6 else ''
                        
                        # Os valores podem estar nas últimas 3 colunas
                        valor_itens_str = celulas[7] if len(celulas) > 7 else '0'
                        icms_origem_str = celulas[8] if len(celulas) > 8 else '0'
                        vr_difal_str = celulas[9] if len(celulas) > 9 else '0'
                        
                        # Validar data
                        if not re.match(r'\d{2}/\d{2}/\d{4}', data):
                            continue
                        
                        # Validar NF (deve ser número)
                        if not nf.isdigit():
                            continue
                        
                        # Função para limpar e converter valor
                        def limpar_valor(valor_str):
                            if not valor_str:
                                return 0.0
                            # Remove tudo exceto dígitos, vírgula e ponto
                            valor_limpo = re.sub(r'[^\d.,]', '', valor_str)
                            # Se não tem nada, retorna 0
                            if not valor_limpo:
                                return 0.0
                            # Remove pontos de milhar e substitui vírgula por ponto
                            if ',' in valor_limpo:
                                # Formato brasileiro: 1.234,56
                                valor_limpo = valor_limpo.replace('.', '').replace(',', '.')
                            return float(valor_limpo)
                        
                        valor_itens = limpar_valor(valor_itens_str)
                        icms_origem = limpar_valor(icms_origem_str)
                        vr_difal = limpar_valor(vr_difal_str)
                        
                        # Validar UF
                        if uf and len(uf) > 2:
                            # UF pode estar misturada com o fornecedor
                            # Tentar extrair apenas as 2 letras
                            match_uf = re.search(r'([A-Z]{2})', uf)
                            if match_uf:
                                uf = match_uf.group(1)
                            else:
                                uf = uf[:2] if len(uf) >= 2 else uf
                        
                        # Limpar fornecedor (remover sufixos/prefixos estranhos)
                        fornecedor = re.sub(r'\s+', ' ', fornecedor).strip()
                        # Remover UF do final do fornecedor se estiver duplicada
                        if uf and fornecedor.endswith(uf):
                            fornecedor = fornecedor[:-len(uf)].strip()
                        
                        # Atualizar métricas
                        metricas['nf_unicas'].add(nf)
                        if data and len(data) >= 10:
                            mes = data[3:5] + '/' + data[6:10]
                            metricas['meses_encontrados'].add(mes)
                        metricas['valor_total_itens'] += valor_itens
                        metricas['valor_total_difal'] += vr_difal
                        metricas['total_linhas_extraidas'] += 1
                        
                        dados_extraidos.append({
                            "Data": data,
                            "NF": nf,
                            "Chave_NFe": chave_nfe,
                            "Fornecedor": fornecedor,
                            "UF": uf,
                            "Descricao": descricao,
                            "CFOP": cfop,
                            "Valor_Itens": round(valor_itens, 2),
                            "ICMS_Origem": round(icms_origem, 2),
                            "VR_DIFAL": round(vr_difal, 2)
                        })
                        
                    except Exception as e:
                        metricas['erros'].append({
                            'pagina': num_pagina + 1,
                            'linha': celulas[:5] if len(celulas) >= 5 else celulas,
                            'erro': str(e)[:100]
                        })
    
    metricas['nf_unicas_count'] = len(metricas['nf_unicas'])
    metricas['meses_count'] = len(metricas['meses_encontrados'])
    metricas['meses_encontrados'] = sorted(list(metricas['meses_encontrados']))
    
    return dados_extraidos, metricas


def extrair_notas_fiscais_texto_fallback(texto_completo):
    """
    Fallback: extrai do texto quando a extração por tabela falha.
    Usa regex mais flexível para lidar com texto desformatado.
    """
    
    linhas = texto_completo.split('\n')
    dados_extraidos = []
    metricas = {
        'total_linhas_extraidas': 0,
        'nf_unicas': set(),
        'valor_total_itens': 0.0,
        'valor_total_difal': 0.0,
    }
    
    cabecalho_encontrado = False
    
    for linha in linhas:
        linha_limpa = linha.strip()
        
        if not cabecalho_encontrado:
            if all(p in linha_limpa for p in ['Data', 'Nº N.F', 'Fornecedor']):
                cabecalho_encontrado = True
            continue
        
        if not linha_limpa:
            continue
            
        if any(p in linha_limpa.upper() for p in ['TOTAL', 'DEMONSTRATIVO']):
            continue
        
        # Verificar se parece uma linha de dados (começa com data)
        if not re.match(r'\d{2}/\d{2}/\d{4}', linha_limpa):
            continue
        
        try:
            # Estratégia: procurar o padrão de valores no final primeiro
            # Últimos 3 valores numéricos + CFOP antes deles
            match_final = re.search(r'(\d{4})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$', linha_limpa)
            
            if not match_final:
                continue
            
            cfop = match_final.group(1)
            
            def converter_valor(v):
                v = re.sub(r'[^\d.,]', '', v)
                if ',' in v:
                    v = v.replace('.', '').replace(',', '.')
                return float(v) if v else 0.0
            
            valor_itens = converter_valor(match_final.group(2))
            icms_origem = converter_valor(match_final.group(3))
            vr_difal = converter_valor(match_final.group(4))
            
            # Parte antes dos valores
            resto = linha_limpa[:match_final.start()].strip()
            
            # Data (10 caracteres)
            data = resto[:10]
            resto = resto[10:].strip()
            
            # NF (primeiro número)
            match_nf = re.match(r'(\d+)', resto)
            if not match_nf:
                continue
            nf = match_nf.group(1)
            resto = resto[match_nf.end():].strip()
            
            # Chave NFe (44 dígitos)
            match_chave = re.match(r'(\d{44})', resto)
            if not match_chave:
                continue
            chave_nfe = match_chave.group(1)
            resto = resto[match_chave.end():].strip()
            
            # Encontrar UF (últimas 2 letras maiúsculas antes da descrição final)
            match_uf = re.search(r'\s+([A-Z]{2})\s+', resto)
            if match_uf:
                fornecedor = resto[:match_uf.start()].strip()
                uf = match_uf.group(1)
                descricao = resto[match_uf.end():].strip()
            else:
                # Fallback: assumir que as últimas 2 letras são UF
                palavras = resto.split()
                if len(palavras) >= 2 and len(palavras[-1]) == 2 and palavras[-1].isupper():
                    uf = palavras[-1]
                    descricao = ''  # Não conseguimos separar
                    fornecedor = ' '.join(palavras[:-1])
                else:
                    continue
            
            fornecedor = re.sub(r'\s+', ' ', fornecedor).strip()
            descricao = re.sub(r'\s+', ' ', descricao).strip()
            
            metricas['nf_unicas'].add(nf)
            metricas['valor_total_itens'] += valor_itens
            metricas['valor_total_difal'] += vr_difal
            metricas['total_linhas_extraidas'] += 1
            
            dados_extraidos.append({
                "Data": data,
                "NF": nf,
                "Chave_NFe": chave_nfe,
                "Fornecedor": fornecedor if fornecedor else "NÃO_IDENTIFICADO",
                "UF": uf,
                "Descricao": descricao,
                "CFOP": cfop,
                "Valor_Itens": round(valor_itens, 2),
                "ICMS_Origem": round(icms_origem, 2),
                "VR_DIFAL": round(vr_difal, 2)
            })
            
        except Exception:
            continue
    
    metricas['nf_unicas_count'] = len(metricas['nf_unicas'])
    
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

st.markdown("### 1. Tipo de relatório:")
tipo_selecionado = st.radio(
    "Escolha:",
    options=["notas_fiscais", "hotel", "exames", "refeicoes"],
    format_func=lambda x: {
        "notas_fiscais": "📋 Notas Fiscais com Produtos (ICMS DIFAL)",
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
    accept_multiple_files=True
)

if st.button("🚀 Extrair Dados e Gerar Excel", type="primary", use_container_width=True):
    if not arquivos_selecionados:
        st.warning("⚠️ Selecione pelo menos um ficheiro PDF.")
    else:
        dados_finais = []
        metricas_gerais = {}
        erros_extração = []
        
        import time
        inicio = time.time()
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, arquivo_pdf in enumerate(arquivos_selecionados):
            nome_arquivo = arquivo_pdf.name
            status_text.text(f"Processando: {nome_arquivo}")
            
            if tipo_selecionado == "notas_fiscais":
                # PRIMEIRO: Tentar extração por tabela (mais confiável)
                dados, metricas = extrair_notas_fiscais_pdfplumber(arquivo_pdf)
                
                # Se extraiu poucas linhas, tentar fallback por texto
                if metricas['total_linhas_extraidas'] < 10:
                    st.warning(f"⚠️ Extração por tabela retornou apenas {metricas['total_linhas_extraidas']} linhas. Tentando fallback por texto...")
                    
                    with pdfplumber.open(arquivo_pdf) as pdf:
                        texto_completo = "\n".join([pagina.extract_text() or "" for pagina in pdf.pages])
                    
                    dados_fallback, metricas_fallback = extrair_notas_fiscais_texto_fallback(texto_completo)
                    
                    if metricas_fallback['total_linhas_extraidas'] > metricas['total_linhas_extraidas']:
                        st.success(f"✅ Fallback extraiu {metricas_fallback['total_linhas_extraidas']} linhas (vs {metricas['total_linhas_extraidas']} por tabela)")
                        dados = dados_fallback
                        metricas.update(metricas_fallback)
                
                metricas['nome_arquivo'] = nome_arquivo
                metricas['num_paginas'] = metricas.get('total_paginas', 0)
                
                dados_finais.extend(dados)
                metricas_gerais = metricas
                erros_extração = metricas.get('erros', [])
            
            progress_bar.progress((idx + 1) / len(arquivos_selecionados))
        
        tempo_total = round(time.time() - inicio, 2)
        status_text.text("✅ Processamento concluído!")
        
        # ============ RESULTADOS ============
        if dados_finais:
            df = pd.DataFrame(dados_finais, columns=COLUNAS_CONFIG[tipo_selecionado])
            
            # ===== MÉTRICAS =====
            if mostrar_metricas and metricas_gerais:
                st.markdown("---")
                st.markdown("## 📊 Painel de Auditoria")
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("📄 Páginas", metricas_gerais.get('num_paginas', metricas_gerais.get('total_paginas', 'N/A')))
                with col2:
                    st.metric("🧾 NFs Únicas", metricas_gerais.get('nf_unicas_count', 0))
                with col3:
                    st.metric("📝 Linhas Extraídas", len(dados_finais))
                with col4:
                    st.metric("⏱️ Tempo", f"{tempo_total}s")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("💰 Valor Total Itens", f"R$ {metricas_gerais.get('valor_total_itens', 0):,.2f}")
                with col2:
                    st.metric("💵 Total DIFAL", f"R$ {metricas_gerais.get('valor_total_difal', 0):,.2f}")
                
                if metricas_gerais.get('meses_encontrados'):
                    st.markdown(f"**📅 Período:** {' | '.join(metricas_gerais['meses_encontrados'])}")
            
            # ===== ERROS =====
            if mostrar_erros and erros_extração:
                st.markdown("---")
                st.markdown(f"## ⚠️ Erros de Extração ({len(erros_extração)})")
                st.dataframe(pd.DataFrame(erros_extração), use_container_width=True)
            
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
                
                # Amostra de fornecedores e UFs
                with st.expander("🔍 Verificar Fornecedores e UFs extraídos"):
                    df_uf = df[['Fornecedor', 'UF']].drop_duplicates().head(30)
                    st.dataframe(df_uf, use_container_width=True)
            
            # ===== DOWNLOAD =====
            st.markdown("---")
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Dados Extraídos')
                
                # Aba de métricas
                if metricas_gerais:
                    dados_met = [
                        ['Arquivo', metricas_gerais.get('nome_arquivo', 'N/A')],
                        ['Páginas', metricas_gerais.get('num_paginas', metricas_gerais.get('total_paginas', 0))],
                        ['Linhas Extraídas', len(dados_finais)],
                        ['NFs Únicas', metricas_gerais.get('nf_unicas_count', 0)],
                        ['Meses', metricas_gerais.get('meses_count', 0)],
                        ['Valor Total Itens (R$)', metricas_gerais.get('valor_total_itens', 0)],
                        ['Total DIFAL (R$)', metricas_gerais.get('valor_total_difal', 0)],
                        ['Tempo Processamento (s)', tempo_total],
                        ['Método', 'pdfplumber (tabelas) + fallback texto'],
                    ]
                    pd.DataFrame(dados_met, columns=['Métrica', 'Valor']).to_excel(
                        writer, index=False, sheet_name='Métricas'
                    )
                
                # Aba de erros
                if erros_extração:
                    pd.DataFrame(erros_extração).to_excel(
                        writer, index=False, sheet_name='Erros Extração'
                    )
            
            st.success(f"✅ {len(dados_finais)} registros extraídos com sucesso!")
            st.download_button(
                label="📥 Baixar Excel Completo",
                data=buffer.getvalue(),
                file_name=f"Extracao_NFs_{metricas_gerais.get('nf_unicas_count', 0)}_NFs.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.error("❌ Nenhum dado foi extraído. O PDF pode não conter tabelas reconhecíveis.")
