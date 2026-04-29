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

# 2. Função de diagnóstico para notas fiscais
def extrair_notas_fiscais_debug(texto_completo):
    """Extrai dados e também retorna diagnóstico completo de cada linha"""
    
    linhas = texto_completo.split('\n')
    dados_extraidos = []
    diagnostico = []
    
    metricas = {
        'total_linhas_arquivo': len(linhas),
        'linhas_processadas': 0,
        'linhas_extraidas': 0,
        'linhas_ignoradas': 0,
        'total_paginas': 1,
        'nf_unicas': set(),
        'meses_encontrados': set(),
        'valor_total_itens': 0.0,
        'valor_total_difal': 0.0,
        'motivos_ignoradas': {}
    }
    
    cabecalho_encontrado = False
    inicio_tabela = False
    
    for i, linha in enumerate(linhas):
        linha_limpa = linha.strip()
        num_linha = i + 1
        
        # Verificar se é o cabeçalho
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
            diagnostico.append({
                'linha': num_linha,
                'tipo': 'PRÉ-CABEÇALHO',
                'conteudo': linha_limpa[:100],
                'status': 'ignorado',
                'motivo': 'Linha antes do cabeçalho'
            })
            continue
        
        # Processar linhas após cabeçalho
        metricas['linhas_processadas'] += 1
        
        # Pular linhas vazias
        if not linha_limpa:
            metricas['linhas_ignoradas'] += 1
            metricas['motivos_ignoradas']['Linha vazia'] = metricas['motivos_ignoradas'].get('Linha vazia', 0) + 1
            diagnostico.append({
                'linha': num_linha,
                'tipo': 'VAZIA',
                'conteudo': '',
                'status': 'ignorado',
                'motivo': 'Linha vazia'
            })
            continue
            
        # Pular títulos e totais
        if 'DEMONSTRATIVO' in linha_limpa.upper():
            metricas['linhas_ignoradas'] += 1
            metricas['motivos_ignoradas']['Título/Demonstrativo'] = metricas['motivos_ignoradas'].get('Título/Demonstrativo', 0) + 1
            diagnostico.append({
                'linha': num_linha,
                'tipo': 'TÍTULO',
                'conteudo': linha_limpa[:100],
                'status': 'ignorado',
                'motivo': 'Título do demonstrativo'
            })
            continue
            
        if 'TOTAIS DO MÊS' in linha_limpa.upper() or 'TOTAL DO MÊS' in linha_limpa.upper():
            # Extrair informações do total para validação
            metricas['linhas_ignoradas'] += 1
            metricas['motivos_ignoradas']['Total do Mês'] = metricas['motivos_ignoradas'].get('Total do Mês', 0) + 1
            diagnostico.append({
                'linha': num_linha,
                'tipo': 'TOTAL_MÊS',
                'conteudo': linha_limpa[:100],
                'status': 'ignorado_registrado',
                'motivo': 'Total do mês (valores registrados para validação)'
            })
            continue
            
        if 'Total Geral' in linha_limpa or 'TOTAIS' in linha_limpa:
            metricas['linhas_ignoradas'] += 1
            metricas['motivos_ignoradas']['Total Geral'] = metricas['motivos_ignoradas'].get('Total Geral', 0) + 1
            diagnostico.append({
                'linha': num_linha,
                'tipo': 'TOTAL_GERAL',
                'conteudo': linha_limpa[:100],
                'status': 'ignorado_registrado',
                'motivo': 'Total geral (valores registrados para validação)'
            })
            continue
        
        # Verificar se é continuação de linha (não começa com data)
        if not re.match(r'^\d{2}/\d{2}/\d{4}', linha_limpa):
            # Verificar se parece continuação de produto
            if re.match(r'^\d', linha_limpa):
                metricas['linhas_ignoradas'] += 1
                metricas['motivos_ignoradas']['Possível continuação (começa com número)'] = metricas['motivos_ignoradas'].get('Possível continuação (começa com número)', 0) + 1
                diagnostico.append({
                    'linha': num_linha,
                    'tipo': 'CONTINUAÇÃO?',
                    'conteudo': linha_limpa[:100],
                    'status': 'alerta',
                    'motivo': 'Possível continuação de descrição - VERIFICAR'
                })
            else:
                metricas['linhas_ignoradas'] += 1
                metricas['motivos_ignoradas']['Texto não reconhecido'] = metricas['motivos_ignoradas'].get('Texto não reconhecido', 0) + 1
                diagnostico.append({
                    'linha': num_linha,
                    'tipo': 'TEXTO',
                    'conteudo': linha_limpa[:100],
                    'status': 'ignorado',
                    'motivo': 'Formato não reconhecido como linha de produto'
                })
            continue
        
        # Tentar extrair dados
        match_data = re.match(r'^(\d{2}/\d{2}/\d{4})\s+', linha_limpa)
        if not match_data:
            metricas['linhas_ignoradas'] += 1
            metricas['motivos_ignoradas']['Data não encontrada'] = metricas['motivos_ignoradas'].get('Data não encontrada', 0) + 1
            diagnostico.append({
                'linha': num_linha,
                'tipo': 'ERRO',
                'conteudo': linha_limpa[:100],
                'status': 'erro',
                'motivo': 'Regex de data falhou'
            })
            continue
            
        data = match_data.group(1)
        mes = data[3:5] + '/' + data[6:10]
        metricas['meses_encontrados'].add(mes)
        
        resto_linha = linha_limpa[match_data.end():]
        
        # NF
        match_nf = re.match(r'(\d+)\s+', resto_linha)
        if not match_nf:
            metricas['linhas_ignoradas'] += 1
            metricas['motivos_ignoradas']['NF não encontrada'] = metricas['motivos_ignoradas'].get('NF não encontrada', 0) + 1
            diagnostico.append({
                'linha': num_linha,
                'tipo': 'ERRO',
                'conteudo': linha_limpa[:100],
                'status': 'erro',
                'motivo': 'Regex NF falhou'
            })
            continue
            
        nf = match_nf.group(1)
        metricas['nf_unicas'].add(nf)
        resto_linha = resto_linha[match_nf.end():]
        
        # Chave NFe (44 dígitos)
        match_chave = re.match(r'(\d{44})\s+', resto_linha)
        if not match_chave:
            metricas['linhas_ignoradas'] += 1
            metricas['motivos_ignoradas']['Chave NFe não encontrada'] = metricas['motivos_ignoradas'].get('Chave NFe não encontrada', 0) + 1
            diagnostico.append({
                'linha': num_linha,
                'tipo': 'ERRO',
                'conteudo': linha_limpa[:100],
                'status': 'erro',
                'motivo': 'Chave NFe (44 dígitos) não encontrada'
            })
            continue
            
        chave_nfe = match_chave.group(1)
        resto_linha = resto_linha[match_chave.end():]
        
        # Fornecedor e UF
        match_fornecedor_uf = re.match(r'(.+?)\s+([A-Z]{2})\s+(.+)', resto_linha)
        if not match_fornecedor_uf:
            metricas['linhas_ignoradas'] += 1
            metricas['motivos_ignoradas']['Fornecedor/UF não identificado'] = metricas['motivos_ignoradas'].get('Fornecedor/UF não identificado', 0) + 1
            diagnostico.append({
                'linha': num_linha,
                'tipo': 'ERRO',
                'conteudo': linha_limpa[:100],
                'status': 'erro',
                'motivo': 'Padrão Fornecedor + UF não encontrado'
            })
            continue
            
        fornecedor = match_fornecedor_uf.group(1).strip()
        uf = match_fornecedor_uf.group(2)
        resto_final = match_fornecedor_uf.group(3)
        
        # Descricao, CFOP, Valores
        match_dados = re.match(r'(.+?)\s+(\d{4})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)', resto_final)
        if not match_dados:
            metricas['linhas_ignoradas'] += 1
            metricas['motivos_ignoradas']['Dados finais não extraídos'] = metricas['motivos_ignoradas'].get('Dados finais não extraídos', 0) + 1
            diagnostico.append({
                'linha': num_linha,
                'tipo': 'ERRO',
                'conteudo': linha_limpa[:150],
                'status': 'erro',
                'motivo': f'Regex valores falhou. Resto: "{resto_final[:80]}..."'
            })
            continue
            
        descricao = match_dados.group(1).strip()
        cfop = match_dados.group(2)
        
        try:
            valor_itens = float(match_dados.group(3).replace('.', '').replace(',', '.'))
            icms_origem = float(match_dados.group(4).replace('.', '').replace(',', '.'))
            vr_difal = float(match_dados.group(5).replace('.', '').replace(',', '.'))
        except ValueError:
            metricas['linhas_ignoradas'] += 1
            metricas['motivos_ignoradas']['Erro conversão numérica'] = metricas['motivos_ignoradas'].get('Erro conversão numérica', 0) + 1
            diagnostico.append({
                'linha': num_linha,
                'tipo': 'ERRO',
                'conteudo': linha_limpa[:100],
                'status': 'erro',
                'motivo': 'Erro ao converter valores numéricos'
            })
            continue
        
        metricas['valor_total_itens'] += valor_itens
        metricas['valor_total_difal'] += vr_difal
        metricas['linhas_extraidas'] += 1
        
        diagnostico.append({
            'linha': num_linha,
            'tipo': 'PRODUTO',
            'conteudo': f"NF {nf} | {descricao[:60]}...",
            'status': 'extraido',
            'motivo': 'Extraído com sucesso'
        })
        
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
    
    metricas['nf_unicas_count'] = len(metricas['nf_unicas'])
    metricas['meses_count'] = len(metricas['meses_encontrados'])
    metricas['meses_encontrados'] = sorted(list(metricas['meses_encontrados']))
    
    return dados_extraidos, metricas, diagnostico


# 3. Função de limpeza para hotel (mantida igual)
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


# 4. Interface do Streamlit
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
                        dados_notas, metricas, diagnostico = extrair_notas_fiscais_debug(texto_completo)
                        metricas['num_paginas'] = num_paginas
                        metricas['nome_arquivo'] = nome_arquivo
                        
                        dados_finais.extend(dados_notas)
                        metricas_gerais = metricas
                        diagnosticos_gerais.extend(diagnostico)
                    
                    elif tipo_selecionado == "hotel":
                        linhas = texto_completo.split('\n')
                        nome_hospede = "NÃO_IDENTIFICADO"
                        for linha in linhas:
                            if "Hóspede principal:" in linha:
                                try:
                                    nome_cru = linha.split("Hóspede principal:")[1].split("|")[0].strip()
                                    nome_hospede = nome_cru.split()[0]
                                except: pass
                                continue
                            if any(x in linha for x in ["PLAZA HOTEL", "Apartamento:", "Fechado", "Pagamentos", "Tarifário:"]):
                                continue
                            extraido = limpar_linha_hotel(linha, nome_hospede)
                            if extraido:
                                dados_finais.append(extraido)
                    
                    elif tipo_selecionado == "exames":
                        linhas = texto_completo.split('\n')
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
                        linhas = texto_completo.split('\n')
                        data_ref = "DATA_NAO_ENCONTRADA"
                        total_ref = "TOTAL_NAO_ENCONTRADO"
                        for linha in linhas:
                            if "Período:" in linha:
                                match = re.search(r'\d{2}/\d{2}/\d{4}', linha)
                                if match:
                                    data_ref = match.group(0)
                            if "Total Geral" in linha:
                                valor = linha.replace("Total Geral", "").replace("|", "").strip()
                                if valor:
                                    total_ref = valor
                        if data_ref != "DATA_NAO_ENCONTRADA" or total_ref != "TOTAL_NAO_ENCONTRADO":
                            dados_finais.append({
                                "Arquivo": nome_arquivo,
                                "Data": data_ref,
                                "Total": total_ref
                            })
                            
            except Exception as e:
                st.error(f"❌ Erro em {nome_arquivo}: {str(e)}")
            
            progress_bar.progress((idx + 1) / len(arquivos_selecionados))
        
        tempo_total = round(time.time() - inicio, 2)
        status_text.text("✅ Processamento concluído!")
        
        # ==================== PAINEL DE DIAGNÓSTICO ====================
        if modo_debug and tipo_selecionado == "notas_fiscais" and diagnosticos_gerais:
            st.markdown("---")
            st.markdown("## 🐛 Diagnóstico Detalhado da Extração")
            
            # Resumo de status
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
            
            # Distribuição de motivos de linhas NÃO extraídas
            st.markdown("### 📊 Motivos de Linhas NÃO Extraídas")
            
            motivos = metricas_gerais.get('motivos_ignoradas', {})
            if motivos:
                df_motivos = pd.DataFrame(list(motivos.items()), columns=['Motivo', 'Quantidade'])
                df_motivos = df_motivos.sort_values('Quantidade', ascending=False)
                
                # Gráfico de barras
                st.bar_chart(df_motivos.set_index('Motivo'))
                
                # Tabela detalhada
                st.dataframe(df_motivos, use_container_width=True)
            
            # Tabela completa de diagnóstico (filtrável)
            st.markdown("### 🔍 Log Completo de Processamento")
            
            # Filtros
            col1, col2 = st.columns(2)
            with col1:
                filtro_status = st.multiselect(
                    "Filtrar por status:",
                    options=['extraido', 'alerta', 'erro', 'ignorado', 'ignorado_registrado'],
                    default=['alerta', 'erro']
                )
            with col2:
                filtro_tipo = st.multiselect(
                    "Filtrar por tipo:",
                    options=df_diag['tipo'].unique().tolist(),
                    default=[]
                )
            
            # Aplicar filtros
            df_filtrado = df_diag.copy()
            if filtro_status:
                df_filtrado = df_filtrado[df_filtrado['status'].isin(filtro_status)]
            if filtro_tipo:
                df_filtrado = df_filtrado[df_filtrado['tipo'].isin(filtro_tipo)]
            
            st.dataframe(df_filtrado[['linha', 'tipo', 'status', 'motivo', 'conteudo']], 
                        use_container_width=True,
                        height=400)
            
            # Botão para exportar diagnóstico
            buffer_diag = io.BytesIO()
            with pd.ExcelWriter(buffer_diag, engine='xlsxwriter') as writer:
                df_diag.to_excel(writer, index=False, sheet_name='Diagnóstico')
                if motivos:
                    df_motivos.to_excel(writer, index=False, sheet_name='Resumo Motivos')
            
            st.download_button(
                label="📥 Baixar Diagnóstico Completo (Excel)",
                data=buffer_diag.getvalue(),
                file_name="diagnostico_extracao.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # ==================== PAINEL DE MÉTRICAS ====================
        if mostrar_metricas and metricas_gerais:
            st.markdown("---")
            st.markdown("## 📊 Painel de Auditoria")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("📄 Páginas", metricas_gerais.get('num_paginas', 'N/A'))
            with col2:
                st.metric("🧾 NFs Únicas", metricas_gerais.get('nf_unicas_count', 0))
            with col3:
                st.metric("📅 Meses", metricas_gerais.get('meses_count', 0))
            with col4:
                st.metric("⏱️ Tempo", f"{tempo_total}s")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("📝 Linhas Extraídas", metricas_gerais.get('linhas_extraidas', 0))
            with col2:
                st.metric("⏭️ Linhas Ignoradas", metricas_gerais.get('linhas_ignoradas', 0))
            with col3:
                total_proc = metricas_gerais.get('linhas_processadas', 1)
                taxa = (metricas_gerais.get('linhas_extraidas', 0) / total_proc * 100) if total_proc > 0 else 0
                st.metric("📈 Taxa Aproveitamento", f"{taxa:.1f}%")
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("💰 Valor Total Itens", f"R$ {metricas_gerais.get('valor_total_itens', 0):,.2f}")
            with col2:
                st.metric("💵 Total DIFAL Extraído", f"R$ {metricas_gerais.get('valor_total_difal', 0):,.2f}")
        
        # ==================== PREVIEW ====================
        if mostrar_preview and dados_finais:
            st.markdown("---")
            st.markdown("## 👁️ Preview dos Dados")
            
            df = pd.DataFrame(dados_finais, columns=COLUNAS_CONFIG[tipo_selecionado])
            
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**Primeiros 5:**")
                st.dataframe(df.head(), use_container_width=True)
            with col2:
                st.markdown("**Últimos 5:**")
                st.dataframe(df.tail(), use_container_width=True)
        
        # ==================== DOWNLOAD ====================
        if dados_finais:
            st.markdown("---")
            df = pd.DataFrame(dados_finais, columns=COLUNAS_CONFIG[tipo_selecionado])
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
            
            st.success(f"✅ {len(dados_finais)} registros extraídos com sucesso!")
            st.download_button(
                label="📥 Baixar Excel",
                data=buffer.getvalue(),
                file_name=f"Extracao_{tipo_selecionado}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.warning("⚠️ Nenhum dado foi extraído. Verifique o diagnóstico acima.")
