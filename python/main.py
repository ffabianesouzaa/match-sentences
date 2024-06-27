# O objetivo é padronizar o nome da ocupação nas planilhas 'ocupadas' e 'ofertadas'
# conforme a planilha 'Skills' (nome com primeira letra maiúscula e demais minúsculas, com todos os acentos

import pandas as pd
from difflib import get_close_matches

# Função que compara strings buscando as que mais são compativeis
def encontra_correspondencia(palavra, lista_palavras):
    correspondencias = get_close_matches(palavra, lista_palavras)
    if correspondencias:
        return correspondencias[0]
    else:
        return None

# Planilha que contém a ocupação conforme desejado
planilha_padrao = pd.read_excel('C:/Users/DELL/Dropbox/match-sentences/Sheets/CBO.xlsx')
# Planilha que será alterada
planilha = pd.read_excel('C:/Users/DELL/Dropbox/match-sentences/Sheets/Ofertadas.xlsx')

# Lê todas as strings armazenadas na coluna 2 da planilha padrão
strings_padrao = planilha_padrao.iloc[:, 1].tolist()

# Iterar sobre as células da coluna 2 da planilha
for indice, celula in enumerate(planilha.iloc[:, 1]):
    # Verificar se a célula existe na lista de strings da planilha 1
    if celula in strings_padrao:
        # Substituir na planilha 2 pela string exata
        planilha.iat[indice, 1] = celula
    else:
        # Encontrar a correspondência mais próxima
        correspondencia = encontra_correspondencia(celula, strings_padrao)
        if correspondencia:
            # Substituir na planilha 2 pela correspondência mais próxima
            planilha.iat[indice, 1] = correspondencia
    print("Feito!", indice, celula)

# Salvar a planilha atualizada
planilha.to_excel('planilha-atualizada.xlsx', index=False)
