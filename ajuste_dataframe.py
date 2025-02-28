import pandas as pd
import numpy as np
from unidecode import unidecode

def ajuste_excel(separados):
    #Ajusta o CPF do cliente para que atenda os requisitos do smart
    separados["CPF/CNPJ"] = separados["CPF/CNPJ"].str.replace(".","", regex=False)
    separados["CPF/CNPJ"] = separados["CPF/CNPJ"].str.replace("-","", regex=False)
    separados["CPF/CNPJ"] = separados["CPF/CNPJ"].str.replace("/","", regex=False)
    
    #Ajusta o telefone do cliente para que atenda os requisitos do smart
    separados["Fone"] = separados["Fone"].str.replace(")","", regex=False)
    separados["Fone"] = separados["Fone"].str.replace("(","", regex=False)
    separados["Fone"] = separados["Fone"].str.replace("-","", regex=False)
    separados.loc[separados["Fone"].str.contains("x"), "Fone"] = 99999999999
    
    #Ajusta os códigos para que fiquem sem o 99 na frente
    sem_ajuste = ["996760", "995066", "994409"]
    
    separados["Código (SKU)"] = separados["Código (SKU)"].astype(str)
    for i in range(len(separados["Código (SKU)"])):
        if separados["Código (SKU)"][i] in sem_ajuste:
            print(i)
        else:
            if separados["Código (SKU)"][i].startswith('99'):
                separados["Código (SKU)"][i] = separados["Código (SKU)"][i].replace('99','',1)
                
    separados["Nome do contato"] = separados["Nome do contato"].apply(unidecode)
    
    #Acrescenta a descrição de manipulado nos itens que estão com tipo interno em branco
    separados["tipoInterno"] = separados["tipoInterno"].replace(np.nan,"manipulado")