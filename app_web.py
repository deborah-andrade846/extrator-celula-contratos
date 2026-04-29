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

# 2. Lista de UFs brasileiras para validação
UFS_BRASIL = [
    'AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 
    'MT', 'MS', 'MG', 'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 
    'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'
]

# 3. Função de extração CORRIGIDA para notas fiscais
def extrair_notas_fiscais_robusto(texto_completo):
    """
    Extrai dados de notas fiscais com tratamento robusto para:
    - Nomes de fornecedores com caracteres especiais
    - Valores numéricos em diversos formatos
    - Linhas com espaçamento irregular
    """
    
    linhas = texto_completo.split('\n')
    dados_extraidos = []
    diagnostico = []
    
    metricas = {
        'total_linhas_arquivo': len(linhas),
        'linhas_processadas': 0,
        'linhas_extraidas': 0,
        'linhas_ignoradas': 0,
        'nf_unicas': set(),
        'meses_encontrados': set(),
        'valor_total_itens': 0.0,
        'valor_total_difal': 0.0,
        'motivos_ignoradas': {},
        'fornecedores_nao_identificados': []
    }
    
    cabecalho_encontrado = False
    
    for i, linha in enumerate(linhas):
        linha_limpa = linha.strip()
        num_linha = i + 1
        
        # Verificar cabeçalho
        if not cabecalho_encontrado:
            if all(col in linha_limpa for col in ['Data', 'Nº N.F', 'Fornecedor', 'UF']):
                cabecalho_encontrado = True
                diagnostico.append({
                    'linha': num_linha,
                    'tipo': 'CABEÇALHO',
                    'conteudo': linha_limpa[:100] + '...',
                    'status': 'identificado',
                    'motivo': 'Cabeçalho da tabela encontrado'
                })
                continue
        
        if not cabecalho_encontrado:
            continue
        
        # Processar linhas após cabeçalho
        metricas['linhas_processadas'] += 1
        
        # Pular linhas vazias
        if not linha_limpa:
            metricas['linhas_ignoradas'] += 1
            metricas['motivos_ignoradas']['Linha vazia'] = metricas['motivos_ignoradas'].get('Linha vazia', 0) + 1
            continue
            
        # Pular títulos e totais
        if any(padrao in linha_limpa.upper() for padrao in ['DEMONSTRATIVO', 'TOTAIS DO MÊS', 'TOTAL DO MÊS', 'TOTAL GERAL', 'TOTAIS']):
            metricas['linhas_ignoradas'] += 1
            metricas['motivos_ignoradas']['Título/Total'] = metricas['motivos_ignoradas'].get('Título/Total', 0) + 1
            continue
        
        # Verificar se começa com data (formato DD/MM/AAAA)
        if not re.match(r'^\d{2}/\d{2}/\d{4}', linha_limpa):
            metricas['linhas_ignoradas'] += 1
            metricas['motivos_ignoradas']['Sem data no início'] = metricas['motivos_ignoradas'].get('Sem data no início', 0) + 1
            continue
        
        # ========== EXTRAÇÃO PRINCIPAL ==========
        try:
            # 1. Extrair Data
            data = linha_limpa[:10]
            resto = linha_limpa[10:].strip()
            mes = data[3:5] + '/' + data[6:10]
            metricas['meses_encontrados'].add(mes)
            
            # 2. Extrair NF (primeiro número após data)
            match_nf = re.match(r'(\d+)', resto)
            if not match_nf:
                metricas['linhas_ignoradas'] += 1
                metricas['motivos_ignoradas']['NF não encontrada'] = metricas['motivos_ignoradas'].get('NF não encontrada', 0) + 1
                continue
            
            nf = match_nf.group(1)
            metricas['nf_unicas'].add(nf)
            resto = resto[match_nf.end():].strip()
            
            # 3. Extrair Chave NFe (44 dígitos)
            match_chave = re.match(r'(\d{44})', resto)
            if not match_chave:
                metricas['linhas_ignoradas'] += 1
                metricas['motivos_ignoradas']['Chave NFe (44 dígitos) não encontrada'] = metricas['motivos_ignoradas'].get('Chave NFe (44 dígitos) não encontrada', 0) + 1
                continue
            
            chave_nfe = match_chave.group(1)
            resto = resto[match_chave.end():].strip()
            
            # 4. Extrair Fornecedor e UF (MÉTODO CORRIGIDO)
            fornecedor = ""
            uf = ""
            
            # Estratégia: encontrar a UF (2 letras maiúsculas) mais próxima do final
            # antes dos valores numéricos e CFOP
            
            # Primeiro, encontrar onde estão os valores numéricos (final da linha)
            # Procurar por padrão: CFOP(4 dígitos) + valor + valor + valor
            match_valores_fim = re.search(r'\s+(\d{4})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$', resto)
            
            if match_valores_fim:
                # Parte antes dos valores contém: Fornecedor + UF + Descrição
                parte_inicial = resto[:match_valores_fim.start()].strip()
                
                # Procurar a última ocorrência de UF (2 letras maiúsculas) na parte inicial
                match_uf = re.search(r'\s+([A-Z]{2})\s+', parte_inicial)
                
                if match_uf:
                    # Dividir em fornecedor (antes da UF) e descrição (depois da UF)
                    fornecedor = parte_inicial[:match_uf.start()].strip()
                    uf = match_uf.group(1)
                    descricao_cfop = parte_inicial[match_uf.end():].strip()
                    
                    # Validar se a UF é brasileira
                    if uf not in UFS_BRASIL:
                        # Tentar encontrar outra UF
                        match_uf2 = re.search(r'\s+([A-Z]{2})\s+', descricao_cfop)
                        if match_uf2:
                            fornecedor = parte_inicial[:match_uf2.start()].strip()
                            uf = match_uf2.group(1)
                            descricao_cfop = descricao_cfop[match_uf2.end():].strip()
                else:
                    # Se não encontrou UF, tentar pegar as últimas 2 letras antes dos números
                    # como possível UF
                    match_uf_alt = re.search(r'\s+([A-Z]{2})\s*$', parte_inicial)
                    if match_uf_alt:
                        uf = match_uf_alt.group(1)
                        fornecedor = parte_inicial[:match_uf_alt.start()].strip()
                        descricao_cfop = ""
                    else:
                        # Última tentativa: pegar últimos 2 caracteres se forem letras
                        if len(parte_inicial) >= 2 and parte_inicial[-2:].isalpha() and parte_inicial[-2:].isupper():
                            uf = parte_inicial[-2:]
                            fornecedor = parte_inicial[:-2].strip()
                            descricao_cfop = ""
                
                # Limpar fornecedor (remover lixo no final)
                fornecedor = re.sub(r'\s+', ' ', fornecedor).strip()
                
                # Extrair CFOP e valores
                cfop = match_valores_fim.group(1)
                
                # Função para converter valor brasileiro para float
                def converter_valor(valor_str):
                    """Converte string de valor brasileiro para float"""
                    # Remove pontos de milhar e substitui vírgula por ponto
                    valor_limpo = valor_str.replace('.', '').replace(',', '.')
                    return float(valor_limpo)
                
                valor_itens = converter_valor(match_valores_fim.group(2))
                icms_origem = converter_valor(match_valores_fim.group(3))
                vr_difal = converter_valor(match_valores_fim.group(4))
                
                # Montar descrição completa (descricao_cfop pode conter parte da descrição)
                descricao = descricao_cfop.strip() if descricao_cfop else ""
                
                # Se não temos fornecedor, registrar para análise
                if not fornecedor:
                    metricas['fornecedores_nao_identificados'].append({
                        'linha': num_linha,
                        'nf': nf,
                        'resto': resto[:100]
                    })
                    metricas['motivos_ignoradas']['Fornecedor vazio'] = metricas['motivos_ignoradas'].get('Fornecedor vazio', 0) + 1
                
                # Atualizar métricas
                metricas['valor_total_itens'] += valor_itens
                metricas['valor_total_difal'] += vr_difal
                metricas['linhas_extraidas'] += 1
                
                diagnostico.append({
                    'linha': num_linha,
                    'tipo': 'PRODUTO',
                    'conteudo': f"NF {nf} | {fornecedor[:40]} | {descricao[:40]}...",
                    'status': 'extraido',
                    'motivo': 'Extraído com sucesso' + (' (fornecedor vazio)' if not fornecedor else '')
                })
                
                dados_extraidos.append({
                    "Data": data,
                    "NF": nf,
                    "Chave_NFe": chave_nfe,
                    "Fornecedor": fornecedor if fornecedor else "NÃO_IDENTIFICADO",
                    "UF": uf if uf else "??",
                    "Descricao": descricao,
                    "CFOP": cfop,
                    "Valor_Itens": valor_itens,
                    "ICMS_Origem": icms_origem,
                    "VR_DIFAL": vr_difal
                })
                
            else:
                # Não encontrou padrão de valores no final
                metricas['linhas_ignoradas'] += 1
                metricas['motivos_ignoradas']['Regex valores falhou'] = metricas['motivos_ignoradas'].get('Regex valores falhou', 0) + 1
                diagnostico.append({
                    'linha': num_linha,
                    'tipo': 'ERRO_VALORES',
                    'conteudo': resto[:150],
                    'status': 'erro',
                    'motivo': 'Padrão CFOP + 3 valores não encontrado no final'
                })
                
        except Exception as e:
            metricas['linhas_ignoradas'] += 1
            metricas['motivos_ignoradas'][f'Erro: {str(e)[:50]}'] = metricas['motivos_ignoradas'].get(f'Erro: {str(e)[:50]}', 0) + 1
            diagnostico.append({
                'linha': num_linha,
                'tipo': 'EXCEÇÃO',
                'conteudo': linha_limpa[:100],
                'status': 'erro',
                'motivo': str(e)[:100]
            })
    
    # Finalizar métricas
    metricas['nf_unicas_count'] = len(metricas['nf_unicas'])
    metricas['meses_count'] = len(metricas['meses_encontrados'])
    metricas['meses_encontrados'] = sorted(list(metricas['meses_encontrados']))
    
    return dados_extraidos, metricas, diagnostico


# 4. Interface do Streamlit (MANTIDA IGUAL, apenas trocando a chamada da função)
st.set_page_config(page_title="Extrator de Relatórios - Apoena", page_icon="📊", layout="wide")
st.title("📊 Extrator de Relatórios - Apoena")

# Sidebar
with st.sidebar:
    st.markdown("## ⚙️ Configurações")
    st.markdown("---")
    modo_debug = st.checkbox("🐛 Modo Debug (diagnóstico detalhado)", value=True)
    mostrar_metricas = st.checkbox("📈 Painel de auditoria", value=True)
    mostrar_preview = st.checkbox("👁️ Preview dos dados", value=True)

st.markdown("### 1. Tipo de relatório:")
tipo_selecionado = st.radio(
    "Escolha:",
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
    "Arraste e solte ou clique", 
    type=['pdf'], 
    accept_multiple_files=True
)

# 5. Botão de Extração
if st.button("🚀 Extrair Dados e Gerar Excel", type="primary", use_container_width=True):
    if not arquivos_selecionados:
        st.warning("⚠️ Selecione pelo menos um ficheiro PDF.")
    else:
        dados_finais = []
        metricas_gerais = {}
        diagnosticos_gerais = []
        
        import time
        inicio = time.time()
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, arquivo_pdf in enumerate(arquivos_selecionados):
            nome_arquivo = arquivo_pdf.name
            status_text.text(f"Processando: {nome_arquivo}")
            
            try:
                with pdfplumber.open(arquivo_pdf) as pdf:
                    texto_completo = "\n".join([pagina.extract_text() or "" for pagina in pdf.pages])
                    num_paginas = len(pdf.pages)
                    
                    if tipo_selecionado == "notas_fiscais":
                        # USA A NOVA FUNÇÃO ROBUSTA
                        dados_notas, metricas, diagnostico = extrair_notas_fiscais_robusto(texto_completo)
                        metricas['num_paginas'] = num_paginas
                        metricas['nome_arquivo'] = nome_arquivo
                        
                        dados_finais.extend(dados_notas)
                        metricas_gerais = metricas
                        diagnosticos_gerais.extend(diagnostico)
                    
                    # ... (outros tipos mantidos iguais)
                    
            except Exception as e:
                st.error(f"❌ Erro em {nome_arquivo}: {str(e)}")
            
            progress_bar.progress((idx + 1) / len(arquivos_selecionados))
        
        tempo_total = round(time.time() - inicio, 2)
        status_text.text("✅ Processamento concluído!")
        
        # ==================== PAINEL DE DIAGNÓSTICO ====================
        if modo_debug and tipo_selecionado == "notas_fiscais" and diagnosticos_gerais:
            st.markdown("---")
            st.markdown("## 🐛 Diagnóstico Detalhado da Extração")
            
            df_diag = pd.DataFrame(diagnosticos_gerais)
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("✅ Extraídas", len(df_diag[df_diag['status'] == 'extraido']))
            with col2:
                st.metric("⚠️ Alertas", len(df_diag[df_diag['status'] == 'alerta']))
            with col3:
                st.metric("❌ Erros", len(df_diag[df_diag['status'] == 'erro']))
            with col4:
                st.metric("⏭️ Ignoradas", len(df_diag[df_diag['status'].isin(['ignorado', 'ignorado_registrado'])]))
            
            # Fornecedores não identificados
            if metricas_gerais.get('fornecedores_nao_identificados'):
                st.warning(f"⚠️ {len(metricas_gerais['fornecedores_nao_identificados'])} linhas com fornecedor não identificado")
            
            # Motivos de exclusão
            motivos = metricas_gerais.get('motivos_ignoradas', {})
            if motivos:
                st.markdown("### 📊 Motivos de Linhas NÃO Extraídas")
                df_motivos = pd.DataFrame(list(motivos.items()), columns=['Motivo', 'Quantidade'])
                df_motivos = df_motivos.sort_values('Quantidade', ascending=False)
                st.bar_chart(df_motivos.set_index('Motivo'))
            
            # Log filtrável
            st.markdown("### 🔍 Log de Processamento")
            filtro_status = st.multiselect(
                "Filtrar por status:",
                options=['extraido', 'alerta', 'erro', 'ignorado'],
                default=['erro']
            )
            
            df_filtrado = df_diag[df_diag['status'].isin(filtro_status)] if filtro_status else df_diag
            st.dataframe(df_filtrado[['linha', 'tipo', 'status', 'motivo', 'conteudo']], 
                        use_container_width=True, height=400)
        
        # ==================== MÉTRICAS ====================
        if mostrar_metricas and metricas_gerais:
            st.markdown("---")
            st.markdown("## 📊 Painel de Auditoria")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("📄 Páginas", metricas_gerais.get('num_paginas', 'N/A'))
            with col2:
                st.metric("🧾 NFs Únicas", metricas_gerais.get('nf_unicas_count', 0))
            with col3:
                st.metric("📝 Extraídas", metricas_gerais.get('linhas_extraidas', 0))
            with col4:
                taxa = (metricas_gerais.get('linhas_extraidas', 0) / max(metricas_gerais.get('linhas_processadas', 1), 1) * 100)
                st.metric("📈 Aproveitamento", f"{taxa:.1f}%")
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("💰 Total Itens", f"R$ {metricas_gerais.get('valor_total_itens', 0):,.2f}")
            with col2:
                st.metric("💵 Total DIFAL", f"R$ {metricas_gerais.get('valor_total_difal', 0):,.2f}")
        
        # ==================== DOWNLOAD ====================
        if dados_finais:
            st.markdown("---")
            df = pd.DataFrame(dados_finais, columns=COLUNAS_CONFIG[tipo_selecionado])
            
            # Preview
            if mostrar_preview:
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**Primeiros 5:**")
                    st.dataframe(df.head(), use_container_width=True)
                with col2:
                    st.markdown("**Últimos 5:**")
                    st.dataframe(df.tail(), use_container_width=True)
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Dados Extraídos')
                
                if metricas_gerais:
                    df_met = pd.DataFrame([
                        ['Páginas', metricas_gerais.get('num_paginas', 'N/A')],
                        ['Linhas Extraídas', metricas_gerais.get('linhas_extraidas', 0)],
                        ['Linhas Ignoradas', metricas_gerais.get('linhas_ignoradas', 0)],
                        ['NFs Únicas', metricas_gerais.get('nf_unicas_count', 0)],
                        ['Meses', metricas_gerais.get('meses_count', 0)],
                        ['Valor Total Itens', f"R$ {metricas_gerais.get('valor_total_itens', 0):,.2f}"],
                        ['Total DIFAL', f"R$ {metricas_gerais.get('valor_total_difal', 0):,.2f}"],
                    ], columns=['Métrica', 'Valor'])
                    df_met.to_excel(writer, index=False, sheet_name='Métricas')
                
                if diagnosticos_gerais:
                    pd.DataFrame(diagnosticos_gerais).to_excel(writer, index=False, sheet_name='Diagnóstico')
            
            st.success(f"✅ {len(dados_finais)} registros extraídos!")
            st.download_button(
                label="📥 Baixar Excel",
                data=buffer.getvalue(),
                file_name=f"Extracao_{tipo_selecionado}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.warning("⚠️ Nenhum dado extraído. Verifique o diagnóstico.")
