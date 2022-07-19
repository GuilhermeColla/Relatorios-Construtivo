'''
Arquivo principal que gera os relatórios para os empreendimentos CPFL.
'''

import os
from pandas import ExcelWriter
from tools import Construtivo
from datetime import datetime


def criar_relatorio(empreendimento: Construtivo.Relatórios_Construtivo):
    empreendimento.remover_colunas(empreendimento.relatorio_gerencial, ["Seq.", "Autor", "Unnamed: 32"])
    empreendimento.relatorio_planejamento = empreendimento.relatorio_planejamento[["Pasta", "Descricao", "Revisão", "Estado Workflow"]]
    empreendimento.relatorio_gerencial.set_index(["Código", "Estado Atual"], inplace=True)
    empreendimento.relatorio_gerencial.sort_index(inplace=True)
    empreendimento.remover_estados(empreendimento.relatorio_planejamento, "Estado Workflow", ["--"])
   # empreendimento.remover_estados(empreendimento.relatorio_gerencial, 
   #                                 [
   #                                 "Obsoleto",
   #                                 "Cancelado",
   #                                 "Documento Previsto",
   #                                 "Liberado para Execução"
   #                                 ])
    empreendimento.dias_com_agentes("Estado Workflow")


def exportar_excel(empreendimento: Construtivo.Relatórios_Construtivo, colunas: list, nome_arquivo: str):
    if not os.path.exists(f"Relatorios"):
        os.mkdir(f"Relatorios")
    
    with ExcelWriter(f"Relatorios/{nome_arquivo}.xlsx") as writer:
        empreendimento.com_acessadas().to_excel(writer, sheet_name= "ComAcessadas", columns=colunas)
        empreendimento.com_CPFL().to_excel(writer, sheet_name="ComCPFL", columns=colunas)
        empreendimento.com_projetista().to_excel(writer, sheet_name="ComsSIEMENS", columns=colunas)
    

if __name__ == "__main__":


    Construtivo.download_gerencial_planejamento("CPFL_Sul_II_VMT_PE")
    Construtivo.download_gerencial_planejamento("CPFL_Sul_II_OSO3_PE")
    Construtivo.download_gerencial_planejamento("CPFL_Sul_II_PAL1_PE")
    

    Osório = Construtivo.Relatórios_Construtivo("Construtivo/gerencial_CPFL_Sul_II_OSO3_PE.csv", "Construtivo/planejamento_CPFL_Sul_II_OSO3_PE.csv")
    VilaMaria = Construtivo.Relatórios_Construtivo("Construtivo/gerencial_CPFL_Sul_II_VMT_PE.csv", "Construtivo/planejamento_CPFL_Sul_II_VMT_PE.csv")
    PortoAlegre = Construtivo.Relatórios_Construtivo("Construtivo/gerencial_CPFL_Sul_II_PAL1_PE.csv", "Construtivo/planejamento_CPFL_Sul_II_PAL1_PE.csv")
    criar_relatorio(Osório)
    criar_relatorio(VilaMaria)
    criar_relatorio(PortoAlegre)
    colunas = ["Pasta", "Descricao", "Revisão", "Estado Workflow", "Dias"]
    exportar_excel(Osório, colunas, "Oso")
    exportar_excel(VilaMaria, colunas, "Vmt")
    exportar_excel(PortoAlegre, colunas, "Pal")
    print(f"Done!\n{datetime.now()}")
