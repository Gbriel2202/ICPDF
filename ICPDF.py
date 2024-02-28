import pdfplumber
import re
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
import os
from tkinter import *
from tkinter.filedialog import askopenfilename

#Declaração de variaveis
excelEu = 'TabelaEu.xlsx' #Caminho da tabela do Eu gerada
excelCo = 'TabelaCo.xlsx' #Caminho da tabela do Co gerada

#Funções
def extractPDF(pdfFile): #Função para conversão de PDF para texto
    with pdfplumber.open(pdfFile) as pdf:
        text = ''
        for page in pdf.pages:
            text += page.extract_text()
    return text

def is_number(s): #Função para checar se o valor pode ser tratado como um float em vez de String
    try:
        float()
        return True
    except ValueError:
        return False 

def filterDFEu(df,min1,max1,min2,max2,min3,max3): #Função de filtro para aquisição do valor
    filteredDF = df[df.iloc[:,0].between(min1,max1)]
    row1 = df[df.iloc[:,0].between(min2,max2)]
    row2 = df[df.iloc[:,0].between(min3,max3)]
    filteredDF = pd.concat([filteredDF, row1, row2], ignore_index = True)
    return filteredDF

def filterDFCo(df,min1,max1,min2,max2): #Função de filtro para aquisição do valor
    filteredDF = df[df.iloc[:,0].between(min1,max1)]
    row1 = df[df.iloc[:,0].between(min2,max2)]
    filteredDF = pd.concat([filteredDF, row1], ignore_index = True)
    return filteredDF   
    
def CoSelect(): #Função de Selecão do botão Cobalto
    minCo1 = 1331.794 #Valor Minimo de filtro 1
    maxCo1 = 1333.194 #Valor Maximo de filtro 1
    minCo2 = 121.36065 #Valor Minimo de filtro 2
    maxCo2 = 122.76065 #Valor Maximo de filtro 2
    fCo = filterDFCo(df,minCo1,maxCo1,minCo2,maxCo2)
    fCo.insert(0,"Hora",hora)
    fCo.insert(0,"Data",dia)
    fCo.insert(8, "Deadtime", deadTime)
    fCo = fCo.drop('BG', axis = 1)
    print(fCo)
    if os.path.isfile(excelCo): #Checagem de se o arquivo ja exista e se ja existir, adicionar o conteudo a tabela
        existingDF = pd.read_excel(excelCo)
        updatedDF = pd.concat([existingDF, fCo], ignore_index=True) 
        updatedDF.to_excel(excelCo, index = False)
    else: #Caso o arquivo não exista, um novo é criado
        fCo.to_excel(excelCo, index = False)
    top.destroy()

def EuSelect(): #Função de seleção do botão Europio
    minEu1 = 1407.313 #Valor Minimo de filtro 1
    maxEu1 = 1408.713 #Valor Maximo de filtro 1
    minEu2 = 121.08 #Valor Minimo de filtro 2
    maxEu2 = 122.48 #Valor Maximo de filtro 2
    minEu3 = 778.2 #Valor Minimo de filtro 3
    maxEu3 = 779.6 #Valor Maximo de filtro 3
    fEu = filterDFEu(df,minEu1,maxEu1,minEu2,maxEu2,minEu3,maxEu3)
    fEu.insert(0,"Hora",hora)
    fEu.insert(0,"Data",dia)
    fEu.insert(8, "Deadtime", deadTime)
    fEu = fEu.drop('BG', axis=1)
    if os.path.isfile(excelEu): #Checagem de se o arquivo ja exista e se ja existir, adicionar o conteudo a tabela
        existingDF = pd.read_excel(excelEu)
        updatedDF = pd.concat([existingDF, fEu], ignore_index=True) 
        updatedDF.to_excel(excelEu, index = False)
    else: #Caso o arquivo não exista, um novo é criado
        fEu.to_excel(excelEu, index = False)
    top.destroy()

#Codigo Principal
file_name = askopenfilename() #Seleção de arquivo

pdf_text = extractPDF(file_name) #Conversão do PDF para texto
pattern = r'\d+\.\d+\s\d+\.\d+\s\d+\.\d+\s\d+\s\d+\.\d+\s\d+\.\d+' #Padrão de organização dos numeros 
matches = [] #Inicialização de variavel para a organização dos numeros encontrados em lista
count = 0 #Inicialização de contador
for line in pdf_text.split('\n'): #Para cada linha no PDF
    count += 1 #Incrementar contador
    if count == 17: #Linha onde se encontra data e hora de aquisição
        data_hora = line #Guarda a linha
        data_hora = data_hora.split(' : ') #Faz a divisão da linha baseada em uma especificação
        dia = data_hora[1].split(' ')[0]
        hora = data_hora[1].split(' ')[1]
    if count == 20: #Linha on de se encontra o Deadtime
        deadTime = line #Guarda a linha
        deadTime = deadTime.split(' : ')[1] #Faz a divisão da linha baseada em uma especificação
    if count <= 34: #Pular até a linha 35
        continue
    numbers = re.findall(pattern, line) #encontrar padrão de numeros na linha
    if numbers:
        matches.extend(numbers) #Se o padrão é encontrado os numeros são salvos na lista

df = pd.DataFrame([num.split() for num in matches], columns = ['Energia', 'Resol', 'Canal','BG','CPS', '1s']) #Organização dos números encontrados em um DataFrame

for col in df.columns: #Substitui o tipo dos valores conferidos na funcão anterior para Float 
    if df[col].apply(is_number).all():
        df[col] = df[col].astype(float)

top = Tk() #Geração da Janela de seleção
top.geometry("300x150") #Tamanho da Janela
radio = IntVar() #Variavel dos botões
lbl = Label(text= "Selecione o tipo de amostra:") #Texto da janela
lbl.pack()
RB1 = Radiobutton(top, text = "Co", variable = radio, value = 1, command = CoSelect) #Botão Cobalto chama função para Cobalto e fecha o programa após execução
RB1.pack(anchor=CENTER)
RB2 = Radiobutton(top, text = "Eu", variable = radio, value = 2, command = EuSelect) #Botão Europio chama função para Europio e fecha o programa após execução
RB2.pack(anchor=CENTER)
label = Label(top)
label.pack()
top.attributes('-topmost', True)
top.update()
top.attributes('-topmost', False)
top.mainloop()