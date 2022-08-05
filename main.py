'''
Arquivo principal que gera os relatórios para os empreendimentos CPFL.
'''

import os
from pandas import ExcelWriter
from tools import Construtivo
from datetime import datetime

def criar_relatorio(empreendimento: Construtivo.Relatórios_Construtivo):
    empreendimento.remover_colunas(empreendimento.relatorio_gerencial, ["Seq.", "Autor", "Unnamed: 32"])
    empreendimento.relatorio_planejamento = empreendimento.relatorio_planejamento[["Pasta", "Descricao", "Revisão", "Estado Workflow", "Data 1° Emissão"]]
    empreendimento.relatorio_gerencial.set_index(["Código", "Estado Atual"], inplace=True)
    empreendimento.relatorio_gerencial.sort_index(inplace=True)
    empreendimento.remover_estados(empreendimento.relatorio_planejamento, "Estado Workflow", ["--"])
    empreendimento.dias_com_agentes("Estado Workflow")
    empreendimento.dias_para_liberar()


def exportar_excel(empreendimento: Construtivo.Relatórios_Construtivo, colunas: list, nome_arquivo: str):
    if not os.path.exists(f"Relatorios"):
        os.mkdir(f"Relatorios")
    
    with ExcelWriter(f"Relatorios/{nome_arquivo}.xlsx") as writer:
        empreendimento.com_CPFL().to_excel(writer, sheet_name="ComCPFL", columns=colunas)
        empreendimento.com_acessadas().to_excel(writer, sheet_name= "ComAcessadas", columns=colunas)
        empreendimento.com_projetista().to_excel(writer, sheet_name="ComProjetista", columns=colunas)
        empreendimento.aprovados().to_excel(writer, sheet_name="Aprovados", columns=colunas)
        empreendimento.aprovados().to_excel(writer, sheet_name="Aprovados", columns=["Pasta", "Descricao", "Revisão", "Estado Workflow", "Data 1° Emissão", "Com CPFL", "Com Projetista", "Com Acessadas", "Tempo total de fluxo"])

        
if __name__ == "__main__":

    empreendimentos = ["CPFL_Sul_II_OSO3_PE",
                        "CPFL_Sul_II_VMT_PE",
                        "CPFL_Sul_II_PAL1_PE",
                        "CPFL_Sul_I_PE"
                        ]
    colunas = ["Pasta", "Descricao", "Revisão", "Estado Workflow", "Dias"]

    downloader = False
    if input("Realizar o download dos relatórios? (s/n)\n") == "s":
        while True:
            try:
                downloader = Construtivo.Download_Relatorios()
            except Exception:
                continue
            break

    for empreendimento in empreendimentos:
        if downloader:
            downloader.acessar_empreendimento(empreendimento)

        while True:
            try:
                if downloader:
                    downloader.driver.switch_to.window(downloader.janela_principal)
                    downloader.download_relatorio_gerencial()
                    downloader.download_visualiza_planejamento()
            except Exception as err:
                print(f"\n\n{err}")
                continue
            else:
                break

        relatorio = Construtivo.Relatórios_Construtivo(f"Construtivo/gerencial_{empreendimento}.csv", f"Construtivo/planejamento_{empreendimento}.csv")
        criar_relatorio(relatorio)
        exportar_excel(relatorio, colunas, empreendimento)
    

    if downloader:
        downloader.fechar()

    print(f"Done!\n{datetime.now()}")