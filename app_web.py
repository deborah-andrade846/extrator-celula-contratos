import pandas as pd
import re
from io import BytesIO

# Mapeamento padrão cobrindo os dois formatos observados
MAPA_COLUNAS_PADRAO = {
    'data': 'data_emissao',
    'dataemissao': 'data_emissao',
    'nº n.f': 'numero_nf',
    'n. fiscal': 'numero_nf',
    'nº nota': 'numero_nf',
    'chave da nota fiscal eletrônica': 'chave_nfe',
    'chave da nota fiscal eletrónica': 'chave_nfe',
    'fornecedor': 'fornecedor',
    'uf': 'uf',
    'ncm': 'ncm',
    'cfop': 'cfop',
    'descrição da mercadoria': 'descricao',
    'descrição da mercadoria/serviço': 'descricao',
    'descrição do documento': 'descricao',
    'descricao': 'descricao',
    'valor dos itens': 'valor_total',
    'valor nf.': 'valor_total',
    'valor total': 'valor_total',
    'bc icms': 'bc_icms',
    'base icms': 'bc_icms',
    'icms origem': 'icms_origem',
    'icms': 'icms_origem',
    '% interna': 'percentual_interna',
    'alíquota interna': 'percentual_interna',
    'vr difal': 'valor_difal',
    'difal': 'valor_difal',
    'obs': 'observacao',
}

COLUNAS_OBRIGATORIAS = ['data_emissao', 'numero_nf', 'valor_total']

def extrair_dados_fiscais(arquivo, mapeamento=None, colunas_obrigatorias=None):
    """
    Extrai dados fiscais de um arquivo Excel ou CSV, normalizando as colunas.

    Parâmetros
    ----------
    arquivo : str, pathlib.Path ou objeto file-like (UploadedFile do Streamlit)
        Caminho para o arquivo ou objeto de arquivo carregado.
    mapeamento : dict, opcional
        Dicionário que mapeia nomes de coluna originais (em minúsculas) para nomes padronizados.
        Se None, utiliza MAPA_COLUNAS_PADRAO.
    colunas_obrigatorias : list, opcional
        Lista de nomes padronizados que devem estar presentes. Se None, usa COLUNAS_OBRIGATORIAS.

    Retorna
    -------
    pd.DataFrame
        DataFrame com colunas padronizadas e dados limpos.
    """
    if mapeamento is None:
        mapeamento = MAPA_COLUNAS_PADRAO
    if colunas_obrigatorias is None:
        colunas_obrigatorias = COLUNAS_OBRIGATORIAS

    # 1. Carregar o arquivo conforme extensão
    if isinstance(arquivo, (str,)):
        nome_arquivo = arquivo
        if nome_arquivo.lower().endswith('.csv'):
            df_raw = _ler_csv_com_delimitador(nome_arquivo)
        else:
            df_raw = pd.read_excel(nome_arquivo, header=None, dtype=str)
    else:
        # Supõe objeto file-like (UploadedFile)
        nome_arquivo = getattr(arquivo, 'name', 'arquivo')
        ext = nome_arquivo.split('.')[-1].lower() if '.' in nome_arquivo else ''
        if ext == 'csv':
            # Lê como bytes e interpreta
            content = arquivo.read()
            df_raw = _ler_csv_com_delimitador(BytesIO(content))
        else:
            df_raw = pd.read_excel(arquivo, header=None, dtype=str)

    # Se houver múltiplas abas, varremos até encontrar dados ou usamos a primeira
    if isinstance(df_raw, dict):  # quando pd.read_excel retorna dict de sheets
        for sheet_name, df_sheet in df_raw.items():
            try:
                return _processar_sheet(df_sheet, mapeamento, colunas_obrigatorias)
            except ValueError:
                continue
        raise ValueError("Nenhuma aba contém as colunas obrigatórias.")

    return _processar_sheet(df_raw, mapeamento, colunas_obrigatorias)


def _ler_csv_com_delimitador(arquivo):
    """Tenta ler CSV detectando delimitador (vírgula, ponto-e-vírgula, tab, pipe)."""
    try:
        # Lê como texto para detecção
        if isinstance(arquivo, BytesIO):
            raw_bytes = arquivo.read()
            arquivo.seek(0)
            sample = raw_bytes[:4096].decode('utf-8', errors='ignore')
        else:
            with open(arquivo, 'rb') as f:
                sample = f.read(4096).decode('utf-8', errors='ignore')

        # Contagem simples de pontuações
        delimitadores = [',', ';', '\t', '|']
        contagens = {d: sample.count(d) for d in delimitadores}
        # Pega o delimitador que mais aparece (excluindo ponto como possível decimal)
        melhor = max(contagens, key=contagens.get)
        if contagens[melhor] < 2:  # muito pouco, assume vírgula
            melhor = ','
        return pd.read_csv(arquivo, sep=melhor, dtype=str, header=None, encoding='utf-8')
    except Exception:
        # Fallback: vírgula com encoding latin1
        arquivo.seek(0)
        return pd.read_csv(arquivo, sep=',', dtype=str, header=None, encoding='latin1')


def _processar_sheet(df_raw, mapeamento, colunas_obrigatorias):
    """Encontra cabeçalho e extrai dados de um DataFrame bruto."""
    # Converte tudo para string e limpa espaços
    df = df_raw.astype(str).applymap(lambda x: x.strip())
    # Substitui strings 'nan', 'None' por NaN real
    df.replace(['nan', 'None', '', ' '], pd.NA, inplace=True)

    # 2. Localizar linha do cabeçalho
    idx_cabecalho = None
    cabecalho_original = None

    # Mapeamento reverso: original -> padronizado (chaves em minúsculas)
    # Vamos construir um padrão de busca com todas as chaves
    chaves_normalizadas = {k.lower(): v for k, v in mapeamento.items()}

    for i, row in df.iterrows():
        # Remove NaNs, deixa apenas strings
        valores_linha = [str(v).lower().strip() for v in row if pd.notna(v)]
        # Verifica quantas colunas casam com as chaves do mapeamento
        matches = sum(1 for v in valores_linha if v in chaves_normalizadas)
        # Exige pelo menos 2 ou uma quantidade razoável (metade das colunas obrigatórias)
        if matches >= max(2, len(colunas_obrigatorias) // 2):
            idx_cabecalho = i
            cabecalho_original = row.tolist()
            break

    if idx_cabecalho is None:
        raise ValueError("Não foi possível localizar a linha de cabeçalho. Verifique o arquivo.")

    # 3. Mapear posições das colunas
    mapa_colunas_pos = {}  # indice -> nome padronizado
    colunas_nao_mapeadas = []
    for idx_col, nome_original in enumerate(cabecalho_original):
        if pd.isna(nome_original):
            continue
        chave = str(nome_original).lower().strip()
        if chave in chaves_normalizadas:
            mapa_colunas_pos[idx_col] = chaves_normalizadas[chave]
        else:
            colunas_nao_mapeadas.append(idx_col)

    # Verifica colunas obrigatórias
    padronizadas_presentes = set(mapa_colunas_pos.values())
    faltantes = set(colunas_obrigatorias) - padronizadas_presentes
    if faltantes:
        raise ValueError(f"Colunas obrigatórias não encontradas: {faltantes}. "
                         f"Cabeçalhos detectados: {list(cabecalho_original)}")

    # 4. Extrair bloco de dados (após cabeçalho até primeira linha totalmente vazia)
    dados = df.iloc[idx_cabecalho + 1:].copy()

    # Remove linhas completamente vazias
    dados = dados.dropna(how='all')

    # Remove linhas que pareçam títulos (ex: contêm apenas uma palavra como "Total")
    mascara_titulo = dados.apply(lambda r: r.astype(str).str.match(r'^(Total|Subtotal|Pagina|Página)\b').any(), axis=1)
    dados = dados[~mascara_titulo]

    # Se não sobrar nada, erro
    if dados.empty:
        raise ValueError("Nenhuma linha de dados encontrada após o cabeçalho.")

    # 5. Selecionar apenas colunas mapeadas e renomeá-las
    colunas_utilizadas = sorted(mapa_colunas_pos.keys())
    df_final = dados.iloc[:, colunas_utilizadas].copy()
    df_final.columns = [mapa_colunas_pos[c] for c in colunas_utilizadas]

    # 6. Conversão de tipos
    # Datas: formato brasileiro dd/mm/aaaa
    if 'data_emissao' in df_final.columns:
        df_final['data_emissao'] = pd.to_datetime(df_final['data_emissao'], dayfirst=True, errors='coerce')

    # Colunas numéricas: formato brasileiro (1.234,56)
    colunas_numericas = ['valor_total', 'bc_icms', 'icms_origem', 'percentual_interna', 'valor_difal']
    for col in colunas_numericas:
        if col in df_final.columns:
            # Converte string: remove pontos (milhar) e troca vírgula por ponto
            df_final[col] = df_final[col].astype(str).str.replace('.', '', regex=False)
            df_final[col] = df_final[col].str.replace(',', '.', regex=False)
            df_final[col] = pd.to_numeric(df_final[col], errors='coerce')

    # Outras colunas numéricas: ncm, cfop, numero_nf podem ficar como string ou int
    # mas se desejar, pode converter para inteiro se não houver zeros à esquerda.

    return df_final.reset_index(drop=True)
