import pandas as pd
import os
import glob

def extrair_dados_excel(caminho_arquivo):
   
    print(f"Lendo o arquivo: {caminho_arquivo}")
    
    try:
        dicionario_planilhas = pd.read_excel(caminho_arquivo, sheet_name=None)
        
        total_linhas_arquivo = 0
        
        print("-" * 50)
        print(f"{'Nome da Planilha (Aba)':<25} | {'Linhas':<8} | {'Colunas':<8}")
        print("-" * 50)
        
        for nome_aba, df in dicionario_planilhas.items():
            linhas, colunas = df.shape
            total_linhas_arquivo += linhas
            
            print(f"{nome_aba:<25} | {linhas:<8} | {colunas:<8}")
            
        print("-" * 50)
        print(f"Total de Abas encontradas no arquivo: {len(dicionario_planilhas)}")
        print(f"Total de Linhas do Arquivo: {total_linhas_arquivo}\n")
        
        return dicionario_planilhas
        
    except FileNotFoundError:
        print(f"Erro: O arquivo '{caminho_arquivo}' não foi encontrado.\n")
        return None
    except Exception as e:
        print(f"Ocorreu um erro ao processar o arquivo '{caminho_arquivo}': {e}\n")
        return None


if __name__ == "__main__":
    pasta_dados = "Dados"
    
    caminho_busca = os.path.join(pasta_dados, "*.xlsx")
    arquivos_excel = glob.glob(caminho_busca)
    
    if not arquivos_excel:
        print(f"Nenhum arquivo .xlsx encontrado na pasta '{pasta_dados}'.")
    else:
        print(f"Encontrados {len(arquivos_excel)} arquivos .xlsx na pasta '{pasta_dados}'.\n")
        print("=" * 60)
        
        lista_de_dataframes = []
        total_linhas_geral_todos_arquivos = 0
        
        for arquivo in arquivos_excel:
            dados_planilha = extrair_dados_excel(arquivo)
            
            if dados_planilha is not None:
                nome_base = os.path.basename(arquivo)
                
                for nome_aba, df in dados_planilha.items():

                    df['Arquivo_Origem'] = nome_base
                    df['Aba_Origem'] = nome_aba
                    
                    lista_de_dataframes.append(df)
                    total_linhas_geral_todos_arquivos += df.shape[0]
                
        print("=" * 60)
        print(f"TOTAL DE LINHAS GERAL (SOMA DE TODOS OS ARQUIVOS REUNIDOS): {total_linhas_geral_todos_arquivos}")
        print("=" * 60)
        
        if lista_de_dataframes:
            print("\nIniciando a união de todos os arquivos e abas apontados. Aguarde...")
   
            df_consolidado = pd.concat(lista_de_dataframes, ignore_index=True)
            
            arquivo_saida_excel = "Base_Consolidada.xlsx"
            
            print(f"Salvando o relatório num único arquivo de Excel: '{arquivo_saida_excel}' ...")
            
            df_consolidado.to_excel(arquivo_saida_excel, index=False)
            
            print("Processamento total e união finalizados com SUCESSO!")
            print(f"O seu arquivo unificado '{arquivo_saida_excel}' foi gerado na sua pasta atual.")
        else:
            print("Não houve dados extraídos para consolidar.")
