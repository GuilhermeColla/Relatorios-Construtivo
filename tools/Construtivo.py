'''
Utilizando PANDAS esse módulo vai formatar os relatórios extraídos do construtivo,
adequando-os para a ferramenta em desenvolvimento.
'''

import re
import pandas as pd
from datetime import datetime


class Relatórios_Construtivo():
    def __init__(self, caminho_relatorio_gerencial: str, caminho_relatorio_planejamento):
        #TODO: Processar o relatório de Planejamento.
        '''
        Cria um objeto "Relatório".
        '''
        self.relatorio_gerencial = pd.read_csv(caminho_relatorio_gerencial, sep=";")
        self.relatorio_planejamento = pd.read_csv(caminho_relatorio_planejamento, sep=";", index_col="Codificação")
        self.obter_título_dos_documentos()
    

    @staticmethod
    def remover_colunas(relatorio, colunas_a_remover: list) -> None:
        relatorio.drop(columns=colunas_a_remover, inplace=True)
    
    
    @staticmethod
    def remover_estados(relatorio: pd.DataFrame, nome_coluna_estado_atual: str, estados_a_remover: list) -> None:
        '''
        A função remove os estados contidos na lista "estados_a_remover" do dataframe.
        Isto é feito pois, no momento, documentos neste estado não tem utilidade.
        '''
        
        for estado in estados_a_remover:
            relatorio.drop(relatorio[relatorio[nome_coluna_estado_atual] == estado].index, inplace=True)
        
    
    def obter_título_dos_documentos(self) -> None:
        '''
        A função obtém o título (código) dos documentos removendo a extensão do arquivo
        (.doc, .docx, .xls, etc...) e a versão do documento no relatório gerencial.
        '''
        códigos = []
        revisões = []
        for título in self.relatorio_gerencial["Título"]:
            
            match = re.search("(?P<codigo>[A-Z|0-9|-]+)-(?P<versao>[\w]+)", título)

            if match:
                código, revisão = match.groups()

            else:
                código, revisão = None, None

            códigos.append(código)
            revisões.append(revisão)
                
        self.relatorio_gerencial.drop(columns="Título", inplace=True)
        self.relatorio_gerencial.insert(0, "Revisão",revisões)
        self.relatorio_gerencial.insert(0, "Código", códigos)


    def dias_com_agentes(self, coluna_estado_atual: str) -> None:

        #TODO: Os estados de "Alerta" acontecem após o estado normal. Devemos fazer 
        #o calculo de acordo com o normal ou com o alerta?
        '''Adiciona uma coluna com os dias que se passaram desde a mudança de
        estado do documento'''
        
        diferencas = []

        for codigo, _ in self.relatorio_planejamento.iterrows():

            estado_atual = self.relatorio_planejamento.loc[codigo][coluna_estado_atual]
            dia_estado_atual = datetime.strptime(self.relatorio_gerencial.loc[codigo, estado_atual][estado_atual][0], "%d/%m/%Y %H:%M")
            
            match = re.search("(^[0-9]+) day", str((datetime.now() - dia_estado_atual)))

            if match:
                diferencas.append(match.groups()[0])
            else:
                diferencas.append("0")

            
        self.relatorio_planejamento["Dias"] = diferencas


    def selecionar_estados(self, estados: list) -> pd.DataFrame:
        return self.relatorio_planejamento.query("`Estado Workflow` == @estados")
    

    def com_projetista(self) -> pd.DataFrame:
        return self.selecionar_estados([
            "Em Análise",
            "Alerta de Em Análise",
            "Para Revisar - pre-analise",
            "Alerta de Para Revisar - pre-analise",
            "Não Aprovado",
            "Alerta de Não Aprovado",
            "Para Revisão",
            "Alerta para Revisão",
            "Aprovado",
            "Alerta Aprovado",
            "Aprovado com Comentários",
            "Alerta Aprovado com Comentários"
            ])


    def com_CPFL(self) -> pd.DataFrame:
        return self.selecionar_estados([
            "Para Análise CPFL",
            "Alerta de Para Análise CPFL",
            "Liberação CPFL",
        ])

    
    def com_acessadas(self) -> pd.DataFrame:
        return self.selecionar_estados([
            "Para Análise Acessado",
            "Para Análise Acessado_",
        ])
