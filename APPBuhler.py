import customtkinter
import Laser_Time
from CTkMessagebox import CTkMessagebox
import ChecaCodigo
import AchaDiretorio
import tkinter
from tkinter import ttk
from tkinter import filedialog
import ExcelLeGrava
import xlsxwriter
import pandas
import PIL
import ezdxf
import os
import sys

aplicativo = customtkinter.CTk()
customtkinter.set_default_color_theme("green")
customtkinter.set_appearance_mode("dark")
aplicativo.title("APP Bühler")
aplicativo.geometry("500x500")

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

    condicao = ChecaCodigo.checagem_codigo(CodigoUser[0:14])

    if condicao == True:
        LocalArquivo = AchaDiretorio.Path_Finder(CodigoUser)

        if not(LocalArquivo == 0):

            tempo, dobra = Laser_Time.laser_calculator(LocalArquivo, espessura_user, material_user)

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

    calculo = ExcelLeGrava.ReadWrite(caminho[0])

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