import os
import pandas as pd
import openpyxl

#Definição de caminho

#Pasta das planilhas
caminho_planilhas = 'C:\\Users\\joao.pereira\\Elementus SA\\Elementus SA - Planejamento\\Comercial\\Propostas\\3_Aprovadas\\Input BI\\'
termo_inicial=len(caminho_planilhas)
#Pasta referente ao log
caminho_log ='C:\\Users\\joao.pereira\\Elementus SA\\Elementus SA - Planejamento\\Comercial\\Propostas\\3_Aprovadas\\Controle_log_PBI'


#Definição de variáveis globais
arquivo_log = open(f'{caminho_log}\\log.txt', 'r')
log = int(arquivo_log.read())
arquivo_log.close()
path = f'{caminho_planilhas}'
listagem = []
data = []


if log == 0: 

    #lista de arquivos por ordem de modificação
    arquivos = os.listdir(path)
    listagem_1 = [(os.path.join(path, arquivo), os.path.getmtime(os.path.join(path, arquivo))) for arquivo in arquivos]
    listagem_1 = sorted(listagem_1, key=lambda arquivo: arquivo[1], reverse=True)
    listagem_1 = pd.DataFrame(listagem_1)
   
    listagem_1.to_csv(f'{caminho_log}\\historico.txt', sep=';', index=False)
   
    #leitura de arquivos na pasta
    for filename in os.listdir(path):
        if filename.endswith('.xlsx'):
            
            df = pd.read_excel(os.path.join(path, filename),sheet_name="PBI",usecols=[0,1,2,3,4,5,6])
            
            for v in range(len(filename)):
                if filename[v] == "_":
                    termo_final=v
                    break
            projeto = str(filename)[:termo_final]

            novos_termos = pd.Series([projeto] * len(df))
            df['PR'] = novos_termos

            data.append(df)

    df_PBI = pd.concat(data, ignore_index=True)
    plan_PBI = pd.DataFrame(df_PBI)


    #covertendo data para evitar erros
    plan_PBI['DATA'] = pd.to_datetime(plan_PBI['DATA'], errors='coerce').dt.strftime('%d/%m/%Y')


    #Limpeza de datas
    for i in range(len(plan_PBI["DATA"])):

        linha_data = plan_PBI["DATA"].iloc[i]

        if str(linha_data)[:3] == "MÊS":

            plan_PBI["DATA"] = plan_PBI["DATA"].apply(lambda x: str(x).replace(str(linha_data), '') if str(x).startswith(str(linha_data)) else x)
    plan_PBI = plan_PBI.dropna(subset=['DATA'])
    PBI_result = pd.DataFrame(plan_PBI)



    #Filtrando #REF nas Descrição
    for j in range(len(plan_PBI["DESCRIÇÃO"])):

        linha_desc = plan_PBI["DESCRIÇÃO"].iloc[j]
        
        if str(linha_desc)[:5] == "#REF!":

            plan_PBI["DESCRIÇÃO"] = plan_PBI["DESCRIÇÃO"].apply(lambda x: str(x).replace(str(linha_desc), '') if str(x).startswith(str(linha_desc)) else x)
    plan_PBI = plan_PBI.dropna(subset=['DATA'])
    plan_PBI = plan_PBI.dropna(subset=['DESCRIÇÃO'])
    PBI_result = pd.DataFrame(plan_PBI)

    
    PBI_result.to_csv(f'{caminho_log}\\framework_PBI.txt', sep=';', index=False)

    arquivo_log = open(f'{caminho_log}\\log.txt', 'w')
    arquivo_log.write("1")
    arquivo_log.close()
   
    
    
if log != 0:

    #lista de arquivos por ordem de modificação
    arquivos = os.listdir(path)
    listagem_2 = [(os.path.join(path, arquivo), os.path.getmtime(os.path.join(path, arquivo))) for arquivo in arquivos]
    listagem_2 = sorted(listagem_2, key=lambda arquivo: arquivo[1], reverse=True)
    listagem_2 = pd.DataFrame(listagem_2)

    listagem_2.to_csv(f'{caminho_log}\\historico2.txt', sep=';', index=False)
    
    historico1 = pd.read_csv(f'{caminho_log}\\historico.txt', delimiter=';')
    lista1 = historico1["0"].tolist()
    historico2 = pd.read_csv(f'{caminho_log}\\historico2.txt',delimiter=';')
    lista2 = historico2["0"].tolist()
    
    framework_PBI = pd.read_csv(f'{caminho_log}\\framework_PBI.txt',delimiter=';')
    framework_PBI = pd.DataFrame(framework_PBI)
    
    #lista arquivos com diferença entre os dois historicos
    for item in lista1:
        if item not in lista2:
            listagem.append(item)

    for item in lista2:
        if item not in lista1:
            listagem.append(item)
    
        
    #caso não haja mudança
    if len(lista1)==len(lista2):
        print("1")
        PBI_result = pd.DataFrame(framework_PBI)
        #historico2.to_csv(f'{caminho_log}\\historico.txt', sep=';', index=False)
            
            
    #caso haja a remoção de arquivo da pasta
    if len(lista1)>len(lista2):
        print("2")
        
        for filename in listagem:
            
            for v in range(termo_inicial,len(filename)):
                if filename[v] == "_":
                    termo_final=v
                    break
            
            projeto = str(filename)[termo_inicial:termo_final]
                    
            framework_PBI = framework_PBI.drop(framework_PBI[framework_PBI['PR'] == projeto].index)
            framework_PBI = pd.DataFrame(framework_PBI)
        PBI_result=pd.DataFrame(framework_PBI)     
        PBI_result.to_csv(f'{caminho_log}\\framework_PBI.txt', sep=';', index=False)
        historico2.to_csv(f'{caminho_log}\\historico.txt', sep=';', index=False)
        
    
    #caso haja novo arquivo 
    if len(lista2)>len(lista1):
        print("3")
      
        for filename in listagem:
        
            df = pd.read_excel(filename,sheet_name="PBI",usecols=[0,1,2,3,4,5,6])
            
            for v in range(termo_inicial,len(filename)):
                if filename[v] == "_":
                    termo_final=v
                    break
            
            projeto = str(filename)[termo_inicial:termo_final]

            novos_termos = pd.Series([projeto] * len(df))
            df['PR'] = novos_termos

            data.append(df)

        df_PBI = pd.concat(data, ignore_index=True)
        plan_PBI = pd.DataFrame(df_PBI)
        

        if len(listagem)==0:
            None
        else:
            plan_PBI['DATA'] = pd.to_datetime(plan_PBI['DATA'], errors='coerce').dt.strftime('%d/%m/%Y')



            #Limpeza de datas
            for i in range(len(plan_PBI["DATA"])):

                linha_data = plan_PBI["DATA"].iloc[i]

                if str(linha_data)[:3] == "MÊS":

                    plan_PBI["DATA"] = plan_PBI["DATA"].apply(lambda x: str(x).replace(str(linha_data), '') if str(x).startswith(str(linha_data)) else x)
            plan_PBI = plan_PBI.dropna(subset=['DATA'])
            
            plan_PBI = pd.DataFrame(plan_PBI)

            #Filtrando #REF nas Descrição
            for j in range(len(plan_PBI["DESCRIÇÃO"])):

                linha_desc = plan_PBI["DESCRIÇÃO"].iloc[j]
                
                if str(linha_desc)[:5] == "#REF!":
                    
                    plan_PBI["DESCRIÇÃO"] = plan_PBI["DESCRIÇÃO"].apply(lambda x: str(x).replace(str(linha_desc), '') if str(x).startswith(str(linha_desc)) else x)
            plan_PBI = plan_PBI.dropna(subset=['DATA'])
            plan_PBI = plan_PBI.dropna(subset=['DESCRIÇÃO'])

            PBI_result = pd.concat([framework_PBI,plan_PBI])
            PBI_result = pd.DataFrame(PBI_result)
            
            PBI_result.to_csv(f'{caminho_log}\\framework_PBI.txt', sep=';', index=False)
            historico2.to_csv(f'{caminho_log}\\historico.txt', sep=';', index=False)


PBI_result = PBI_result.dropna(subset=['DATA'])
PBI_result = PBI_result.dropna(subset=['DESCRIÇÃO'])
PBI_result = pd.DataFrame(PBI_result)
PBI_result = PBI_result.drop_duplicates()
PBI_result = pd.DataFrame(PBI_result)