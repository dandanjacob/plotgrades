#Reading the xlsx document
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import openpyxl
arq = openpyxl.load_workbook('planilhas.xlsx') 



#ploting a bars graph
matricula_aluno = int(67106);
nome_materia = [];
nota_materia = [];


#for each planilha in the xlsx document, I'll see if the matricula_aluno is in the planilha
for i in range(1, len(arq.sheetnames)):
        planilha = pd.read_excel('planilhas.xlsx', sheet_name=i)
        #find matricula_aluno in planilha
        for j in range(0, len(planilha)-1): #forline in planilha
            #read colunm A line j in the xlsx document
            if(planilha.iloc[j, 0] == matricula_aluno): #if matricula_aluno is in the planilha
                nome_materia.append(arq.sheetnames[i]); #add the name of the materia in the list
                #for each avaliation in planilha
            nota_final = 0;
            soma_pesos = 0;
            for k in range(2, (len(planilha.columns))): #for each avaliation in planilha
                    if (k % 2 == 0): 
                        soma_pesos += soma_pesos + planilha.iloc[j, k]; #sum the weights
                        nota_final = nota_final + planilha.iloc[j,k]*planilha.iloc[j,k+1]; #add the nota of the avaliation in the list
            nota_materia.append(nota_final*10/soma_pesos); #add the nota of the materia in the list
            break;
################################
 #We have the name of the materia and the nota of the materia,
 # so height of the bar and how many bar there is
################################

#PESOS DAS MATERIAS
#defining the weights of the bars
weights = [];
planilha = pd.read_excel('planilhas.xlsx', sheet_name='Geral');
for i in range(0, len(nome_materia)):
    for j in range(0, len(planilha)):
        if(planilha.iloc[j, 0] == nome_materia[i]):
            weights.append(planilha.iloc[j, 2]);
  


#CR DO PERÍODO
#media ponderada de weighs e nota_materia
media_ponderada = 0;
for i in range(0, len(weights)):
    media_ponderada += weights[i]*nota_materia[i];
media_ponderada = media_ponderada/sum(weights);

#printweights and nota_materia



#define space between bars
space = [];
acumulado = 0;
space_around = 0.1;
for i in range(0, len(weights)):
    if(i==0):
        space.append(0.5*weights[i] + space_around);
        acumulado = acumulado + weights[i] + space_around;
    else:
        space.append(acumulado + weights[i]*0.5 + space_around);
        acumulado = acumulado + weights[i]+ space_around;

#PERÍODO
#extracting numbers from a nome_materia
periodo = '';
p = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '.'];
for i in range(0, len(nome_materia[0])):
    if(nome_materia[0][i] in p):
        periodo = periodo + nome_materia[0][i];


#NOME DA MATÉRIA
#deleting numbers and symbols in nome_materia
for i in range(0, len(nome_materia)):
    nome_materia[i] = nome_materia[i].replace("1", "");
    nome_materia[i] = nome_materia[i].replace("2", "");
    nome_materia[i] = nome_materia[i].replace("3", "");
    nome_materia[i] = nome_materia[i].replace("4", "");
    nome_materia[i] = nome_materia[i].replace("5", "");
    nome_materia[i] = nome_materia[i].replace("6", "");
    nome_materia[i] = nome_materia[i].replace("7", "");
    nome_materia[i] = nome_materia[i].replace("8", "");
    nome_materia[i] = nome_materia[i].replace("9", "");
    nome_materia[i] = nome_materia[i].replace("0", "");
    nome_materia[i] = nome_materia[i].replace("-", "");
    nome_materia[i] = nome_materia[i].replace(".", "");


print(weights);
print(space);
print(nota_materia);


# plot a histogram
plt.bar(space, nota_materia, weights);
plt.xticks(space, nome_materia);
plt.yticks([]);
plt.title(periodo)
#add a vertical line in the mean of the notas

#LINHA NOTA MÍNIMA 
plt.axhline(6, color='b', linestyle='dashed', linewidth=1)
plt.axhline(media_ponderada, color='g', linestyle='dashed', linewidth=1)
#plot media_ponderada in the y axis
plt.text(0, media_ponderada, 'CR: ' + str(round(media_ponderada, 2)), fontsize=6, color='g')
plt.text(0, 6, str(6), fontsize=6, color='b')
plt.show()


# 