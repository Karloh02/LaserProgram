import customtkinter
from CTkMessagebox import CTkMessagebox
import tkinter
from tkinter import ttk
from tkinter import filedialog
import xlsxwriter
import pandas
import PIL
import ezdxf
import os
import sys
import math

aplicativo = customtkinter.CTk()
customtkinter.set_default_color_theme("green")
customtkinter.set_appearance_mode("dark")
aplicativo.title("APP Bühler")
aplicativo.geometry("500x500")

def laser_calculator(directory_user, thickness_user, material_user):
    sheet_thickness = ["1", "1.5", "2", "2.5", "3", "4", "5", "6", "8", "10", "12", "15"]

    #carbon sheets parameters
    time_cut_carb = [0.35, 0.37, 0.40, 0.42, 0.45, 0.51, 0.58, 0.66, 0.85, 1.1, 1.42, 2.08]
    time_hole_carb = [0.0002, 0.0005, 0.0006, 0.0007, 0.0008, 0.0023, 0.0035, 0.0043, 0.0087, 0.0286, 0.0624, 0.298]

    #stainless sheets parameters
    time_cut_stain = [0.25, 0.29, 0.33, 0.39, 0.45, 0.6, 0.81, 1.08, 1.96, 3.54, 6.4]
    time_hole_stain = [0.00072, 0.00087, 0.00122, 0.00149, 0.00176, 0.0026, 0.0039, 0.0052, 0.00823, 0.013, 0.0273]

    #parameters to calculate
    total_perimeter = 0             #Total perimeter of the component
    bends = 0                       #Total number of bends                         
    interior_profiles = 0           #Total number of interior profiles - Holes

    #user inputs
    file = directory_user 
    material = material_user
    thickness = str(thickness_user)

    #reading the dxf file - saving it as a modelspace 
    dwg = ezdxf.readfile(file)
    msp = dwg.modelspace()

    #this for function will iterate trough all the modelspace parameters 
    for e in msp:
        
        #This part will determine how many lines there are in the file. 
        if e.dxf.layer == "IV_BEND_DOWN" or e.dxf.layer == "IV_BEND":
            bends += 1

        #This part will read the layers, and calculate the perimeter of those respective layers.
        #If it is not a defined bend line, it will measured into the perimeter
        else: #e.dxf.layer == "OUTER" or e.dxf.layer == "IV_ARC_CENTERS" or e.dxf.layer == "IV_INTERIOR_PROFILES" or e.dxf.layer == "IV_TOOL_CENTER" or e.dxf.layer == "IV_TOOL_CENTER_DOWN" or e.dxf.layer == "IV_FEATURE_PROFILES" or e.dxf.layer == "AM_KONT" or e.dxf.layer == "Visible (ISO)" or e.dxf.layer == "0": 

            #this part server the pupose to read how many holes there are in the part
            if e.dxf.layer == "IV_INTERIOR_PROFILES":
                interior_profiles += 1
            
            #This part reads the distance "dl" of lines 
            if e.dxftype() == "LINE":
                
                dl = abs(math.sqrt((e.dxf.start[0] - e.dxf.end[0])**2 + (e.dxf.start[1] - e.dxf.end[1])**2)) 
                total_perimeter += dl
            
            #This part reads the distance "dc" of circles
            elif e.dxftype() == "CIRCLE": 

                dc = 2*math.pi*e.dxf.radius
                total_perimeter += dc
            
            #This part reads the distance "da" of arcs
            elif e.dxftype() == "ARC":
                
                end = e.dxf.end_angle
                start = e.dxf.start_angle
                angle = abs(start - end)

                if end == 0:
                    end = 360 
                    angle = abs(start - end)
                
                da = abs(2*math.pi*e.dxf.radius*angle/360)
                total_perimeter += da

            #This part read the distance "ds" of splines - an aproximation since splines are curved.
            elif e.dxftype() == "SPLINE":

                points = e.control_points
                
                for i in range(len(points) - 1):
                    ds = math.sqrt((points[i][0] - points[i + 1][0])**2 + (points[i][1] - points[i + 1][1])**2)
                    total_perimeter += ds

            #This part reads the distance "dp" of polylines - an aproximation since polylines can be curved 
            elif e.dxftype() == "POLYLINE":
                
                poly_points = e.points()

                polyline_vectors = []
                for k in poly_points:
                    polyline_vectors.append(ezdxf.math.Vec3(k))
                
                for f in range(len(polyline_vectors) - 1):
                    dp = math.sqrt((polyline_vectors[f][0] - polyline_vectors[f + 1][0])**2 + (polyline_vectors[f][1] - polyline_vectors[f + 1][1])**2)
                    total_perimeter += dp*1.05
                
                total_perimeter += (math.sqrt((polyline_vectors[0][0] - polyline_vectors[len(polyline_vectors) - 1][0])**2 + (polyline_vectors[0][1] - polyline_vectors[len(polyline_vectors) - 1][1])**2))*1.05

    #Where the time of the cut is calculated
    total_perimeter = round(total_perimeter)  

                
    if material == "Carbono":
        cut_time = round((total_perimeter/1000)*time_cut_carb[sheet_thickness.index(thickness)] + time_hole_carb[sheet_thickness.index(thickness)]*(interior_profiles + 1), 2)

    elif material == "Inox":
        cut_time = round((total_perimeter/1000)*time_cut_stain[sheet_thickness.index(thickness)] + time_hole_stain[sheet_thickness.index(thickness)]*(interior_profiles + 1), 2)
        
    #The function returns the time it takes to cut the part and, the number of bends; information useful in the future.

    return(cut_time, bends)


def Path_Finder(search_name):

    #where the search will happen
    chiffre = search_name[0:4]

    #creating the searching directory 
    search_dir = r"\\ctbn33\AVOR\__Desenhos_Windchill" + "/" + chiffre

    #all the files in the directory
    try:
        dir_files = os.listdir(search_dir)
    except:
        return(0)
    match_files = []

    #Search for the files that match
    for i in range(len(dir_files)):
        if dir_files[i].endswith(".dxf") and dir_files[i].startswith(search_name[0:14]):
            match_files.append(dir_files[i])
    
    #On the matchfiles get the latest version, if multiple match files exist. 
    if len(match_files) > 1:
        try:
            version = 0
            index = 0

            for i in range(len(match_files)):
                n_version = int(match_files[i][-6:-4])
                if n_version >= version:
                    index = i
                    version = n_version
            
            latest_version_dir = search_dir + "/" + match_files[index]
        except:
            latest_version_dir = search_dir + "/" + match_files[0]
    
    else:
        latest_version_dir = search_dir + "/" + match_files[0]
    
    if len(match_files) == 0:
        return(0)
    
    return(latest_version_dir)

def achaEspessura(nome):
    espessurasTodas = ["1", "1.5", "2", "2.5", "3", "4", "5", "6", "8", "10", "12", "15"]

    k = -1
    espessuraVar = ""
    espessuraNova = ""
    while (nome[k] != "x") and (nome[k] != "X") and (k >= -6):
        espessuraVar += nome[k]

        k -= 1
    
    for i in range(len(espessuraVar)):
        espessuraNova += espessuraVar[-i - 1]

    #se a espessura não for uma das padrões arredonda a espessra para cima, para dar o tempo de corte.
    if not(espessuraNova in espessurasTodas):
        espessuraNova = str(round(float(espessuraNova)))

        if not(espessuraNova is espessurasTodas):
            espessuraNova = False

    #Acha o material da chapa
    if ("1.0" in nome) or ("DD11" in nome):
        material = "Carbono"
    
    elif "1.4" in nome:
        material = "Inox"
    
    else:
        material = False
    
    return(espessuraNova, material)


def pathFinderExcel(nome, diretorio):

    try:
        dir_files = os.listdir(diretorio)
    except:
        return(0)
    match_files = []

    for i in range(len(dir_files)):
        if dir_files[i].endswith(".dxf") and dir_files[i].startswith(nome[0:14]):
            match_files.append(dir_files[i])

    if len(match_files) > 0:

        return(diretorio + match_files[0])   

    else:
        return(0)
    

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
            
                espessura, material = achaEspessura(Material_var[j])

                if (material != False) and (espessura != False):
                
                    EspessuraMaterial.append(espessura)
                    MaterialMaterial.append(material)

                else:
                    EspessuraMaterial.append("Erro de material")
                    MaterialMaterial.append("Erro de material")

            else:
                EspessuraMaterial.append("Erro de material")
                MaterialMaterial.append("Erro de material")

        diretorioFiles = path[:-32]
        LaserTempos = []
        DobraNumero = []
        DobraTempos = []
        Area = []

        #Faz o cálculo de laser, dobra, area
        for k in range(len(EspessuraMaterial)):

            if not(EspessuraMaterial[k] == "Erro de material"):

                diretorio = pathFinderExcel(Nome[k], diretorioFiles)
                
                if diretorio != False:
                    if EspessuraMaterial[k] != "Erro de material" and MaterialMaterial != "Erro de material":
                        laser, dobra = laser_calculator(diretorio, EspessuraMaterial[k], MaterialMaterial[k])
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
            
            elif (EspessuraMaterial[k] == "Erro de material"):
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

def checagem_codigo(codigo):
    alfabeto = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", " "]
    numeros = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9"]

    condicao = True

    for i in range(len(codigo)):
        
        if len(codigo) == 14:
            
            if (i <= 3) and (codigo[i] in alfabeto) and (condicao == True):
                condicao = True
            
            elif i == 4 and codigo[i] == "-" and condicao == True:
                condicao = True
            
            elif i == 10 and codigo[i] == "-" and condicao == True:
                condicao = True
            
            elif (i > 4 and i < 10) and (codigo[i] in numeros) and condicao == True:
                condicao = True

            elif (i > 10 and i <= 13) and (codigo[i] in numeros) and condicao == True:
                condicao = True

            else:
                condicao = False
        else:
            condicao = False
        
    return(condicao)

def checkbox_carbono():
    Inox_Checkbox.deselect()

def checkbox_inox():
    Carbono_Checkbox.deselect()

#função que irá roda após apertar o botão
def TempoIndividual():

    CodigoUser = SelectCode.get()

    #pega a espessura dada pelo usuário
    espessura_user = OpcaoEspessura.get()

    #pega o material fornecido pelo usuário
    if Carbono_Checkbox.get() == 1:
        material_user = "Carbono"
    
    else: 
        material_user = "Inox"

    if material_user == "Inox" and espessura_user == "15":
        CTkMessagebox(title = "Erro", message = "Espessura conflitante com o material, selecione uma espessura inferior a 15mm para Inox!", icon = "warning", option_1 = "OK")

    #pega o código fornecido pelo usuário. 
    #teste para ver se o código fornecdio esta correto. 
    # AAAA - 00000 - 000

    condicao = checagem_codigo(CodigoUser[0:14])

    if condicao == True:
        LocalArquivo = Path_Finder(CodigoUser)

        if not(LocalArquivo == 0):

            tempo, dobra = laser_calculator(LocalArquivo, espessura_user, material_user)

            popup = tkinter.Toplevel(aplicativo)
            popup.title("Informações individuais")
            popup.config(bg = "black")
            popup.geometry("450x100")

            TextLaser = "Tempo de corte para o item " + CodigoUser[0:14] + " é de " + str(tempo) + " minutos."
            customtkinter.CTkLabel(popup, text = TextLaser, text_color = "light green").pack()

            if int(dobra) > 0: 
                TextDobras = "Número de dobras é de " + str(dobra)
                customtkinter.CTkLabel(popup, text = TextDobras, text_color = "red").pack()

        else: 

            CTkMessagebox(title = "Erro", message = "Arqvuivo não encontrado na pasta Desenhos Windchill", icon = "warning", option_1 = "OK")
    
    else:

        CTkMessagebox(title = "Erro", message = "Código não atende as especificações. Revisar o código.", icon = "warning", option_1 = "OK")

caminho = [0]

def SelecionaDiretorio():
    caminho[0] = (filedialog.askopenfilename(title="Select a File", filetype=(('All files','*.xlsm'),('All files','.xlsx'))))

    return(caminho)

def CalculoExcel():

    calculo = ReadWrite(caminho[0])

    if calculo == True:
        CTkMessagebox(title = "Calculado com suecesso", message = "Planilha salva com sucesso", icon = "check", option_1 = "OK")
    
    else:
        CTkMessagebox(title = "Erro", message = "Erro ao gerar a planilha", icon = "cancel", option_1 = "OK")

    return()


aplicativo.grid_columnconfigure((0), weight=1)

nameLabel2 = customtkinter.CTkLabel(aplicativo, text = "BOM Large", text_color = "light green", font = ("Arial", 25), justify = "center", width = 250)
nameLabel2.grid(row = 0, column = 0, padx = 10, pady = 10)

button_directory = customtkinter.CTkButton(aplicativo, text = "Selecione a planilha", command = SelecionaDiretorio)
button_directory.grid(row = 1, column = 0, padx = 10, pady = 10)

ButtonCalc = customtkinter.CTkButton(aplicativo, text = "Calcular", command = CalculoExcel)
ButtonCalc.grid(row = 2, column = 0, pady = 20, padx = 20)

nameLabel = customtkinter.CTkLabel(aplicativo, text = "Peças individuais", text_color = "light green", font = ("Arial", 25), justify = "center", width = 250)
nameLabel.grid(row = 0, column = 1, padx = 10, pady = 10, columnspan = 2)

SelectCode = customtkinter.CTkEntry(aplicativo, placeholder_text = "Entre com o material", font = ("Arial", 17), justify = "center", width = 160, height = 30)
SelectCode.grid(row = 1, column = 1, padx = 10, pady = 10, columnspan = 2)

LabelMaterial = customtkinter.CTkLabel(aplicativo, text = "Selecione a matéria prima", text_color = "light green", font = ("Arial", 17), justify = "center", width = 250)
LabelMaterial.grid(row = 2, column = 1, padx = 10, pady = 10, columnspan = 2)

Carbono_checkbox_Value = customtkinter.BooleanVar(value = True)
Carbono_Checkbox = customtkinter.CTkCheckBox(aplicativo, text = "Carbono", command = checkbox_carbono, width = 10, font = ("Arial", 12), variable = Carbono_checkbox_Value)
Carbono_Checkbox.grid(row = 3, column = 1, padx = 10, pady = 15)

Inox_Checkbox = customtkinter.CTkCheckBox(aplicativo, text = "Inox", command = checkbox_inox, width = 10, font = ("Arial", 12))
Inox_Checkbox.grid(row = 3, column = 2, padx = 0, pady = 15)

LabelEspessura = customtkinter.CTkLabel(aplicativo, text = "Selecione a espessura", text_color = "light green", font = ("Arial", 15), justify = "left")
LabelEspessura.grid(row = 4, column = 1, padx = 0, pady = 10)

OpcaoEspessura = customtkinter.CTkOptionMenu(aplicativo, values = ["1","1.5","2","2.5","3","4","5","6","8","10","12","15"], width = 60)
OpcaoEspessura.grid(row = 4, column = 2, padx = 10, pady = 10)

BotaoStartIndividual = customtkinter.CTkButton(aplicativo, text = "Calcular tempo individual", width = 250, font = ("Arial", 15), command = TempoIndividual)
BotaoStartIndividual.grid(row = 5, column = 1, padx = 10, pady = 20, columnspan  = 2)

aplicativo.mainloop()