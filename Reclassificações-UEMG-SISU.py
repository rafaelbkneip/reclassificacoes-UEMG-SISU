import xlsxwriter  
from time import sleep
from datetime import date
from PyPDF2 import PdfReader, PdfWriter

readpdf = PdfReader()


geral=[]
geral_aux = []

for m in range(len(readpdf.pages)):

    text = readpdf.pages[m].extract_text()

    listas = text.split("\n")

    print(m)

    aux = []

    linhas = []


    if (m==0):
        primeira = 29
        ultima=3

    elif(m==len(readpdf.pages)):
        primeira = 0
        ultima = 15

    else:
        primeira = 0
        ultima = 1

    for i in range(primeira, len(listas)-ultima):

        if(listas[i] == 'Suplente'):
            
            aux.append(listas[i])
            linhas.append(aux)
            aux = []

        elif(listas[i] == 'Classificado'):
            
            aux.append(listas[i])
            linhas.append(aux)
            aux = []

        else:
            
            aux.append(listas[i])

    
            
    print("\n")

 
    cont = 0


    nome = []
    nome_aux = ""
    unidade=""
    curso=""

    
    for r in range(len(linhas)):
        
        geral_aux.append(linhas[r][-1])
        geral_aux.append(linhas[r][-2])
        geral_aux.append(linhas[r][-3])

        for k in range(len(linhas[r])):
            if(linhas[r][k].isnumeric()):
                geral_aux.append(linhas[r][k])
                break
                
        for j in range(k+1, len(linhas[r])-3):
            
            nome_aux =  nome_aux  + " " + linhas[r][j] 

        geral_aux.append(nome_aux)
        nome_aux = ""

        if((linhas[r][k-2])=='Área Básica de'):
            
            
            geral_aux.append('Área Básica de Ensino (ABI)')
            geral_aux.append(linhas[r][k-3])
            k=k-1
        
        else:
            geral_aux.append(linhas[r][k-1])
            geral_aux.append(linhas[r][k-2])


        for l in range(len(linhas[r])):
            if(linhas[r][l]== 'ABAETÉ' or linhas[r][l]== 'DIVINÓPOLIS' or linhas[r][l]== 'BARBACENA' or linhas[r][l]== 'CAMPANHA' or linhas[r][l]== 'CARANGOLA' or linhas[r][l]== 'CLÁUDIO' or linhas[r][l]== 'DIAMANTINA'  or linhas[r][l]== 'GUIGNARD'  or linhas[r][l]== 'EDUCAÇÃO'  or linhas[r][l]== 'NEGÓCIOS' or  linhas[r][l]== 'FRUTAL' or linhas[r][l]== 'ITUIUTABA' or linhas[r][l]==  'MONLEVADE' or linhas[r][l]== 'LEOPOLDINA' or linhas[r][l]=='PASSOS' or linhas[r][l]=='CALDAS' or linhas[r][l]=='UBÁ' or linhas[r][l]=='CATAGUASES' or linhas[r][l]=='DESIGN'):
                break
            
           

            elif(linhas[r][l]== 'IBIRITÉ'):
                break

        for j in range(l+1):
            unidade =  unidade + linhas[r][j]  + " "
            
        geral_aux.append(unidade)
        unidade = ""

        for j in range(l+1,k-2):
            curso =  curso + linhas[r][j]  + " "
        

        geral_aux.append(curso)
        curso = ""
        

        geral.append(geral_aux)
        geral_aux = []



for p in range(len(geral)):
    
    geral[p][0], geral[p][2], geral[p][-1], geral[p][4] = geral[p][4],  geral[p][-1], geral[p][2], geral[p][0]


workbook = xlsxwriter.Workbook('UNIMONTES - 2.xlsx')
sheet = workbook.add_worksheet()   

for i in range(len(geral)):

    

    if (  len(geral[i][7].split(" ")) > 6 or len(geral[i][7].split(" "))==2  or len(geral[i][7].split(" "))==3  ):

        for j in range(len(geral[i])):
            geral[i][j] = "Impossível ler. Acesse o pdf diretamente no site para esse curso e colocação"
            
for i in range(len(geral)):
    sheet.write(i+1, 0, geral[i][0])
    sheet.write(i+1, 1, geral[i][1])
    sheet.write(i+1, 2, geral[i][2])
    sheet.write(i+1, 3, geral[i][3])
    sheet.write(i+1, 4, geral[i][4])
    sheet.write(i+1, 5, geral[i][5])
    sheet.write(i+1, 6, geral[i][6])
    sheet.write(i+1, 6, geral[i][7])
    sheet.write(i+1, 8, geral[i][8])
    
workbook.close()