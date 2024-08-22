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

