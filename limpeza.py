import pandas as pd

def limpar_dados_para_sql(caminho_entrada, caminho_saida):
    print(f"Iniciando a leitura do arquivo base: {caminho_entrada} ...")
    
    df = pd.read_excel(caminho_entrada)
    linhas_antes, colunas_antes = df.shape
    print(f"-> Tamanho Original: {linhas_antes} linhas e {colunas_antes} colunas.\n")
    
    colunas_controle = ['Arquivo_Origem', 'Aba_Origem']
    colunas_dados = [col for col in df.columns if col not in colunas_controle]
    
    print("Realizando limpezas e padronizações...")

    df.dropna(subset=colunas_dados, how='all', inplace=True)
    
    df.dropna(axis=1, how='all', inplace=True)
    
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime('%Y-%m-%d')
            print(f"  - Coluna '{col}' padronizada para data do SQL.")
            
    linhas_depois, colunas_depois = df.shape
    print("\n-> Limpeza concluída!")
    print(f"-> Tamanho Final Pronto pro SQL: {linhas_depois} linhas e {colunas_depois} colunas.")
    
    print(f"\nSalvando o arquivo final de saída como '{caminho_saida}'...")
    df.to_excel(caminho_saida, index=False)
    
    print("Sucesso! Tabela limpa, enxugada e com datas padronizadas.")


if __name__ == "__main__":
    arquivo_entrada = "Base_Consolidada.xlsx"
    arquivo_saida = "Base_Limpa_SQL.xlsx"
    
    try:
        limpar_dados_para_sql(arquivo_entrada, arquivo_saida)
    except FileNotFoundError:
        print(f"Erro: O arquivo '{arquivo_entrada}' não foi encontrado. Execute o arquivo 'etl.py' primeiro para juntar os dados.")
    except Exception as e:
        print(f"Ocorreu um erro inesperado durante a limpeza: {e}")
