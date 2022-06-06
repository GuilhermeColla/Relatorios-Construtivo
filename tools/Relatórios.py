'''
Este pacote de funções terá como entrada um pd.DataFrame com o formato genérico:

Código | Atributos (Versão, autor, data de emissão etc) | Fluxo 

Sendo que código é composto por uma única coluna, e as outras podem conter n
colunas.
Fluxo deve conter a informação da data que o fluxo foi alterado.

Um exemplo pode ser observado com a classe implementada no módulo 
"Formatar_relatório"
'''


import pandas as pd
from datetime import datetime


def dias_com_agentes(
    relatorio_base: pd.DataFrame,
    coluna_estado_atual: str
    ) -> pd.DataFrame:

    #TODO: Os estados de "Alerta" acontecem após o estado normal. Devemos fazer 
    #o calculo de acordo com o normal ou com o alerta?.=
    
    diferencas = []
    for codigo, _ in relatorio_base.iterrows():
        estado_atual = relatorio_base.loc[codigo][coluna_estado_atual]
        dia_estado_atual = datetime.strptime(relatorio_base.loc[codigo][estado_atual], "%d/%m/%Y %H:%M")
        diferencas.append(str(datetime.now() - dia_estado_atual))
    
    relatorio_base["Dias"] = pd.Series(diferencas).values
    
    return relatorio_base
    

