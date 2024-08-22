import pandas
import numpy
import SalvaEspessura
import Laser_Time
import AchaDiretorio
import xlsxwriter


def ReadWrite(path):

    try:
        #Lê certas collunas do excel
        dfNome = (pandas.read_excel(path, usecols = "B:B")).values.tolist()
        dfUnit = (pandas.read_excel(path, usecols = "G:G")).values.tolist()
        dfName = (pandas.read_excel(path, usecols = "I:I")).values.tolist()
        dfMass = (pandas.read_excel(path, usecols = "BD:BD")).values.tolist()

        Mass = []
        Unit = []
        Material_var = []
        Nome = []

        #transforma em dataframes do panda para listas e já organiza as planilhas
        for i in range(len(dfNome)):

            if dfUnit[i][0] == "kilograms":

                Mass.append(float(dfMass[i - 1][0]))
                Unit.append(dfUnit[i][0])
                Material_var.append(dfName[i][0])
                Nome.append(dfNome[i - 1][0])
            
        #acerta o material e a espessura

        EspessuraMaterial = []
        MaterialMaterial = []

        for j in range(len(Material_var)):

            if Material_var[j].startswith("sheet metal") or Material_var[j].startswith("metal sheet"):
            
                espessura, material = SalvaEspessura.achaEspessura(Material_var[j])

                if (material != False) and (espessura != False):
                
                    EspessuraMaterial.append(espessura)
                    MaterialMaterial.append(material)

                else:
                    EspessuraMaterial.append("Erro de material")
                    MaterialMaterial.append("Erro de material")

            else:
                EspessuraMaterial.append("Erro de material")
                MaterialMaterial.append("Erro de material")

        diretorioFiles = path[:-31]
        LaserTempos = []
        DobraNumero = []
        DobraTempos = []
        Area = []

        #Faz o cálculo de laser, dobra, area
        for k in range(len(EspessuraMaterial)):
            try:
                diretorio = AchaDiretorio.pathFinderExcel(Nome[k], diretorioFiles)
                
                if diretorio != 0:
                    if EspessuraMaterial[k] != "Erro de material" and MaterialMaterial != "Erro de material":
                        laser, dobra = Laser_Time.laser_calculator(diretorio, EspessuraMaterial[k], MaterialMaterial[k])
                        LaserTempos.append(laser)
                        DobraNumero.append(dobra)

                        if MaterialMaterial[k] == "Carbono":
                            densidade = 7800
                    
                        else:
                            densidade = 8000

                        Area.append(round(2*float(Mass[k])/(densidade*(float(EspessuraMaterial[k])/1000)),3))
                        DobraTempos.append(dobra*0.5)

                    else:
                        LaserTempos.append("Erro ao calcular")
                        DobraNumero.append("Erro ao calcular")
                        DobraTempos.append("Erro ao calcular")
                        Area.append("Erro ao calcular")

                else:
                    LaserTempos.append("Diretório não encontrado")
                    DobraTempos.append("Diretório não encontrado")
                    Area.append("Diretório não encontrado")
                    DobraNumero.append("Diretório não encontrado")
            
            except: 
                LaserTempos.append("Erro ao calcular")
                DobraTempos.append("Erro ao calcular")
                Area.append("Erro ao calcular")
                DobraNumero.append("Erro ao calcular")

        # LaserTempo DobraTempos Area EspessuraMaterial MaterialMaterial Nome Mass DobraNumero
        #Parte que irá remover os duplicados das listas.



        #Parte que irá criar a planilha do excel. 
        NomePlanilha = path[:-17] + "_CalculosDeTempo.xlsx"
        workbook = xlsxwriter.Workbook(NomePlanilha)
        worksheet = workbook.add_worksheet("Cálculo")

        worksheet.set_column("A:A", 20)
        worksheet.set_column("B:B", 16)
        worksheet.set_column("C:C", 22)
        worksheet.set_column("D:D", 15)
        worksheet.set_column("E:E", 23)
        worksheet.set_column("F:F", 23)
        worksheet.set_column("G:G", 23)
        worksheet.set_column("H:H", 23)

        worksheet.write(0, 0, "Código")
        worksheet.write(0, 1, "Matéria prima")
        worksheet.write(0, 2, "Massa individual (kg)")
        worksheet.write(0, 3, "Espessura")
        worksheet.write(0, 4, "Tempo de laser (min)")
        worksheet.write(0, 5, "Número de dobras")
        worksheet.write(0, 6, "Tempo de dobra (min)")
        worksheet.write(0, 7, "Área (m2)")

        format = workbook.add_format({'bg_color' : 'red'})

        dupRemovedNome = []
        dupRemovedMaterial = []
        dupRemovedMass = []
        dupRemovedEspessura = []
        dupRemovedLaser = []
        dupRemovedDobraNum = []
        dupRemovedDobraTem = []
        dupRemovedArea = []

        for t in range(len(LaserTempos)):
            if not(Nome[t] in dupRemovedNome):
            
                dupRemovedNome.append(Nome[t])
                dupRemovedArea.append(Area[t])
                dupRemovedMass.append(Mass[t])
                dupRemovedDobraNum.append(DobraNumero[t])
                dupRemovedDobraTem.append(DobraTempos[t])
                dupRemovedEspessura.append(EspessuraMaterial[t])
                dupRemovedLaser.append(LaserTempos[t])
                dupRemovedMaterial.append(MaterialMaterial[t])

        for u in range(len(dupRemovedLaser)):

            if DobraNumero[u] == 1:
                worksheet.write(u + 1, 0, str(dupRemovedNome[u]), format)
                worksheet.write(u + 1, 1, str(dupRemovedMaterial[u]), format)
                worksheet.write(u + 1, 2, str(dupRemovedMass[u]), format)
                worksheet.write(u + 1, 3, str(dupRemovedEspessura[u]), format)
                worksheet.write(u + 1, 4, str(dupRemovedLaser[u]), format)
                worksheet.write(u + 1, 5, str(dupRemovedDobraNum[u]), format)
                worksheet.write(u + 1, 6, str(dupRemovedDobraTem[u]), format)
                worksheet.write(u + 1, 7, str(dupRemovedArea[u]), format)

            else:
                worksheet.write(u + 1, 0, str(dupRemovedNome[u]))
                worksheet.write(u + 1, 1, str(dupRemovedMaterial[u]))
                worksheet.write(u + 1, 2, str(dupRemovedMass[u]))
                worksheet.write(u + 1, 3, str(dupRemovedEspessura[u]))
                worksheet.write(u + 1, 4, str(dupRemovedLaser[u]))
                worksheet.write(u + 1, 5, str(dupRemovedDobraNum[u]))
                worksheet.write(u + 1, 6, str(dupRemovedDobraTem[u]))
                worksheet.write(u + 1, 7, str(dupRemovedArea[u]))
        
        workbook.close()

        return(True)
    
    except:
        return(False)

ReadWrite(r"C:\Users\U57534\OneDrive - Bühler\Desktop\LBEB-41728-001\LBEB-41728-001_00_BOMLarge.xlsm")