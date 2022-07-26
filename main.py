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
        empreendimento.com_CPFL().to_excel(writer, sheet_name="ComCPFL", columns=colunas)
        empreendimento.com_acessadas().to_excel(writer, sheet_name= "ComAcessadas", columns=colunas)
        empreendimento.com_projetista().to_excel(writer, sheet_name="ComProjetista", columns=colunas)
    

if __name__ == "__main__":

    empreendimentos = ["CPFL_Sul_II_OSO3_PE",
                        "CPFL_Sul_II_VMT_PE",
                        "CPFL_Sul_II_PAL1_PE",
                        "CPFL_Sul_I_PE"
                        ]
    colunas = ["Pasta", "Descricao", "Revisão", "Estado Workflow", "Dias"]

    download = False
    if download := input("Realizar o download dos relatórios? (s/n)\n") == "s":
        downloader = Construtivo.Download_Relatorios()

    for empreendimento in empreendimentos:

        while True:
            try:    
                downloader.acessar_empreendimento(empreendimento)
                downloader.download_relatorio_gerencial()
                downloader.download_visualiza_planejamento()
            except Exception:
                continue
            break

        relatorio = Construtivo.Relatórios_Construtivo(f"Construtivo/gerencial_{empreendimento}.csv", f"Construtivo/planejamento_{empreendimento}.csv")
        criar_relatorio(relatorio)
        exportar_excel(relatorio, colunas, empreendimento)
    
    if downloader:
        downloader.fechar()

    print(f"Done!\n{datetime.now()}")