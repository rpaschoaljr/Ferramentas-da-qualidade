import tkinter as tk
from tkinter import filedialog
import pandas as pd

def ler_excel(nome_arquivo):
    # Lendo o arquivo Excel, pulando a primeira linha (índice 0)
    df = pd.read_excel(nome_arquivo, skiprows=1)
    
    # Criando o dicionário para armazenar os dados
    dados = {}
    
    # Iterando sobre as linhas do DataFrame
    for index, linha in df.iterrows():
        # Verificar se a linha está vazia
        if linha.isnull().values.all():
            break  # Parar o loop se a linha estiver vazia
        
        # Verificar se há dados vazios na linha
        if linha.isnull().any():
            continue  # Passar para a próxima iteração se houver dados vazios
        
        # Lendo o valor da coluna "Inspeção"
        inspecao = linha.iloc[0]
        
        # Convertendo "Inspeção aprovada" para 1 e "Inspeção reprovada" para 0
        if inspecao == "Inspeção aprovada":
            ponto_quantidade0 = 1
        elif inspecao == "Inspeção reprovada":
            ponto_quantidade0  = 0
        
        # Lendo os valores das outras colunas
        codigo = linha.iloc[1]
        razao_social = str(linha.iloc[2]).upper()  # Convertendo para maiúsculas
        data_prevista = pd.to_datetime(linha.iloc[3])  # Convertendo para datetime
        data_real = pd.to_datetime(linha.iloc[4])  # Convertendo para datetime
        
        # Convertendo a quantidade prevista e real para float
        quantidade_prevista = float(str(linha.iloc[5]).replace(".", "").replace(",", "."))
        quantidade_real = float(str(linha.iloc[6]).replace(".", "").replace(",", "."))
        
        # Calculando os pontos de qualidade
        ponto_qualidade = ponto_quantidade0 
        
        # Calculando os pontos de quantidade
        if quantidade_real < quantidade_prevista:
            ponto_quantidade = 0
        else:
            ponto_quantidade = 1
        
        # Calculando os pontos de pontualidade
        if data_real > data_prevista:
            ponto_pontualidade = 0
        else:
            ponto_pontualidade = 1
        
        # Calculando o IMF
        imf = ponto_qualidade*6 + ponto_quantidade*2.5 + ponto_pontualidade*1.5
        
        # Calculando o resultado
        resultado = "Aprovado" if imf > 7 else "Reprovado"
        
        # Adicionando os valores ao dicionário
        if (codigo, razao_social) in dados:
            # Se o código e a razão social já existirem no dicionário, adicione os pontos aos valores existentes
            dados[(codigo, razao_social)]['Ponto de Qualidade'] += ponto_qualidade
            dados[(codigo, razao_social)]['Ponto de Quantidade'] += ponto_quantidade
            dados[(codigo, razao_social)]['Ponto de Pontualidade'] += ponto_pontualidade
            dados[(codigo, razao_social)]['IMF'] += imf
            dados[(codigo, razao_social)]['Resultado'] = resultado
        else:
            # Se o código e a razão social não existirem no dicionário, crie uma nova entrada
            dados[(codigo, razao_social)] = {
                'Ponto de Qualidade': ponto_qualidade,
                'Ponto de Quantidade': ponto_quantidade,
                'Ponto de Pontualidade': ponto_pontualidade,
                'IMF': imf,
                'Resultado': resultado
            }
    
    # Calcular a média dos pontos para cada grupo
    for key in dados:
        dados[key]['Ponto de Qualidade'] /= len(dados[key])
        dados[key]['Ponto de Quantidade'] /= len(dados[key])
        dados[key]['Ponto de Pontualidade'] /= len(dados[key])
        dados[key]['IMF'] /= len(dados[key])
        
        # Formatar os números com duas casas decimais e substituir os pontos por vírgulas
        dados[key]['IMF'] = str(round(dados[key]['IMF'], 2)).replace(".", ",")
    
    # Remover a chave 'Quantidade de Inspeções' do dicionário
    for key in dados:
        if 'Quantidade de Inspeções' in dados[key]:
            del dados[key]['Quantidade de Inspeções']
    
    # Ordenar os dados por ordem alfabética da razão social
    dados_ordenados = dict(sorted(dados.items(), key=lambda x: x[0][1]))
    
    return dados_ordenados

def selecionar_arquivo():
    root = tk.Tk()
    root.withdraw()
    nome_arquivo = filedialog.askopenfilename()
    root.destroy()  # Fechar a janela do Tkinter após selecionar o arquivo
    return nome_arquivo

# Selecionar o arquivo Excel usando uma janela de diálogo
nome_arquivo = selecionar_arquivo()

# Se o usuário cancelar a seleção do arquivo, nome_arquivo será uma string vazia
if nome_arquivo:
    # Chamando a função para ler o arquivo Excel e armazenar os dados em um dicionário
    dados = ler_excel(nome_arquivo)

    # Exibindo os dados agrupados
    for key, value in dados.items():
        print(key, value)
else:
    print("Nenhum arquivo selecionado.")
