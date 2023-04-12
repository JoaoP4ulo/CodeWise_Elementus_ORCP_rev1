import os
import pandas as pd
import openpyxl
import datetime

dias_da_semana = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado', 'Domingo']
hoje = datetime.date.today()
dia_da_semana = hoje.weekday()
now = datetime.datetime.now()

dia_da_semana_str = dias_da_semana[dia_da_semana]
data_hora = now.strftime("%d/%m/%Y %H:%M:%S")
hora = now.strftime("%H:%M:%S")

arquivo_log = open('log.txt', 'r+')
log = int(arquivo_log.read())
arquivo_log.close()

path = 'C:\\Users\\jppel\Desktop\\ELEMENTUS\\ORCP_Python\\planilhas'

listagem=[]
data = []



if log == 0: 

    #lista de arquivos por ordem de modificação
    arquivos = os.listdir(path)
    listagem_1 = [(os.path.join(path, arquivo), os.path.getmtime(os.path.join(path, arquivo))) for arquivo in arquivos]
    listagem_1 = sorted(listagem_1, key=lambda arquivo: arquivo[1], reverse=True)
    listagem_1 = pd.DataFrame(listagem_1)
    listagem_1["2"] = data_hora
    listagem_1.to_csv('historico.txt', sep=';', index=False)
   
    


    for filename in os.listdir(path):
        if filename.endswith('.xlsx'):
            
            df = pd.read_excel(os.path.join(path, filename),sheet_name="PBI",usecols=[0,1,2,3,4,5,6])
            projeto = str(filename)[:5]

            novos_termos = pd.Series([projeto] * len(df))
            df['PR'] = novos_termos

            data.append(df)

    df_PBI = pd.concat(data, ignore_index=True)
    plan_PBI = pd.DataFrame(df_PBI)



    plan_PBI['DATA'] = pd.to_datetime(plan_PBI['DATA'], errors='coerce').dt.strftime('%d/%m/%Y')


    #Limpeza de datas
    for i in range(len(plan_PBI["DATA"])):

        linha_data = plan_PBI["DATA"].iloc[i]

        if str(linha_data)[:3] == "MÊS":

            plan_PBI["DATA"] = plan_PBI["DATA"].apply(lambda x: str(x).replace(str(linha_data), '') if str(x).startswith(str(linha_data)) else x)
    plan_PBI = plan_PBI.drop(plan_PBI[plan_PBI['DATA'] == ''].index)


    #Filtrando #REF nas Descrição
    for j in range(len(plan_PBI["DESCRIÇÃO"])):

        linha_desc = plan_PBI["DESCRIÇÃO"].iloc[j]
        
        if str(linha_desc)[:5] == "#REF!":

            plan_PBI["DESCRIÇÃO"] = plan_PBI["DESCRIÇÃO"].apply(lambda x: str(x).replace(str(linha_desc), '') if str(x).startswith(str(linha_desc)) else x)
    plan_PBI = plan_PBI.drop(plan_PBI[plan_PBI['DESCRIÇÃO'] == ''].index)   

    plan_PBI.to_csv('framework_PBI.txt', sep=';', index=False)





    arquivo_log = open('log.txt', 'w')
    arquivo_log.write("1")
else:

    #lista de arquivos por ordem de modificação
    arquivos = os.listdir(path)
    listagem_2 = [(os.path.join(path, arquivo), os.path.getmtime(os.path.join(path, arquivo))) for arquivo in arquivos]
    listagem_2 = sorted(listagem_2, key=lambda arquivo: arquivo[1], reverse=True)
    listagem_2 = pd.DataFrame(listagem_2)
    listagem_2["2"] = data_hora
    listagem_2.to_csv('historico2.txt', sep=';', index=False)
    
    historico1 = pd.read_csv('historico.txt', delimiter=';')
    lista1 = historico1["0"].tolist()
    historico2 = pd.read_csv('historico2.txt',delimiter=';')
    lista2 = historico2["0"].tolist()
    
    for item in lista1:
        if item not in lista2:
            listagem.append(item)

    for item in lista2:
        if item not in lista1:
            listagem.append(item)



    for filename in listagem:
      
        df = pd.read_excel(filename,sheet_name="PBI",usecols=[0,1,2,3,4,5,6])
        projeto = str(filename)[55:60]

        novos_termos = pd.Series([projeto] * len(df))
        df['PR'] = novos_termos

        data.append(df)

    df_PBI = pd.concat(data, ignore_index=True)
    plan_PBI = pd.DataFrame(df_PBI)

    plan_PBI['DATA'] = pd.to_datetime(plan_PBI['DATA'], errors='coerce').dt.strftime('%d/%m/%Y')



    #Limpeza de datas
    for i in range(len(plan_PBI["DATA"])):

        linha_data = plan_PBI["DATA"].iloc[i]

        if str(linha_data)[:3] == "MÊS":

            plan_PBI["DATA"] = plan_PBI["DATA"].apply(lambda x: str(x).replace(str(linha_data), '') if str(x).startswith(str(linha_data)) else x)
    plan_PBI = plan_PBI.drop(plan_PBI[plan_PBI['DATA'] == ''].index)


    #Filtrando #REF nas Descrição
    for j in range(len(plan_PBI["DESCRIÇÃO"])):

        linha_desc = plan_PBI["DESCRIÇÃO"].iloc[j]
        
        if str(linha_desc)[:5] == "#REF!":

            plan_PBI["DESCRIÇÃO"] = plan_PBI["DESCRIÇÃO"].apply(lambda x: str(x).replace(str(linha_desc), '') if str(x).startswith(str(linha_desc)) else x)
    plan_PBI = plan_PBI.drop(plan_PBI[plan_PBI['DESCRIÇÃO'] == ''].index)

    framework_PBI = pd.read_csv('framework_PBI.txt',delimiter=';')
    framework_PBI = pd.DataFrame(framework_PBI)

    PBI_result = pd.concat([framework_PBI,plan_PBI])
    print(PBI_result)

    
    if dia_da_semana_str == 'Sexta-feira' and hora =="12:00:00":
        PBI_result.to_csv('framework_PBI.txt', sep=';', index=False)
