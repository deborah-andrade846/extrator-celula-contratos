# No seu app_web.py, adicione:

def processar_excel_upload(arquivo_upload):
    """
    Processa arquivo Excel/CSV enviado pelo Streamlit.
    
    Args:
        arquivo_upload: Objeto UploadedFile do Streamlit
        
    Returns:
        DataFrame padronizado
    """
    nome = arquivo_upload.name
    
    # Verificar extensão
    if nome.endswith('.csv'):
        # CSV: ler como texto e passar para pandas
        conteudo = arquivo_upload.read().decode('utf-8')
        arquivo_upload.seek(0)
        df = extrair_dados_excel(io.StringIO(conteudo), nome)
    else:
        # Excel: ler diretamente
        df = extrair_dados_excel(arquivo_upload, nome)
    
    return df
