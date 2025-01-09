import os
import pandas as pd

def gerar_relatorio():
    pasta = os.getcwd()
    lista_final = []

    for nome_arquivo in os.listdir(pasta):
        if nome_arquivo.lower().endswith(".xlsx"):
            caminho = os.path.join(pasta, nome_arquivo)
            if "Analise" in nome_arquivo:
                continue  # Ignora o próprio script caso esteja em XLSX
            try:
                dados = pd.read_excel(caminho, sheet_name="Dados", header=0)
                resumo = pd.read_excel(caminho, sheet_name="Resumo", header=0)

                rastreavel = dados.iloc[0, 11]    # col L linha 2
                endereco_inicial = dados.iloc[-1, 1]  
                endereco_final = dados.iloc[0, 4]     
                inicio_jornada = dados.iloc[-1, 0]    
                fim_jornada = dados.iloc[0, 3]        
                tempo_total = resumo.iloc[0, 2]       

                lista_final.append({
                    "Rastreavel": rastreavel,
                    "Endereço Inicial": endereco_inicial,
                    "Endereço Final": endereco_final,
                    "Início da Jornada": inicio_jornada,
                    "Fim da Jornada": fim_jornada,
                    "Tempo Total": tempo_total
                })
            except:
                pass

    df_final = pd.DataFrame(lista_final)
    nome_arquivo = input("Informe o nome do arquivo (sem extensão): ")
    caminho_salvar = os.path.join(os.getcwd(), nome_arquivo + ".xlsx")
    df_final.to_excel(caminho_salvar, index=False)
    print("Relatório gerado com sucesso!")

if __name__ == "__main__":
    gerar_relatorio()