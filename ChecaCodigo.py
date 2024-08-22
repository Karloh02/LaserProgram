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