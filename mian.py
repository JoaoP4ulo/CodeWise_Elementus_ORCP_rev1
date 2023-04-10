import os
import pandas as pd
import openpyxl


path = 'C:\\Users\\jppel\\Desktop\\ELEMENTUS\\ORCP_Python'
data = []


for filename in os.listdir(path):
    if filename.endswith('.xlsx'):
        
        df = pd.read_excel(os.path.join(path, filename),sheet_name="PBI",usecols=[0,1,2,3,4,5,6])
        projeto = str(filename)[:5]

        novos_termos = pd.Series([projeto] * len(df))
        df['PR'] = novos_termos

        data.append(df)

plan_PBI = pd.concat(data, ignore_index=True)
caminho = "PR392_ORCPP2112203.xlsx"


#Limpeza de datas
plan_PBI['DATA'] = pd.to_datetime(plan_PBI['DATA'], errors='coerce').dt.strftime('%d/%m/%Y')
for i in range(len(plan_PBI["DATA"])):

    linha_data = plan_PBI["DATA"].iloc[i]
    linha_fam = plan_PBI["FAM"].iloc[i]
    linha_item = plan_PBI["ITEM"].iloc[i]
    linha_desc = plan_PBI["DESCRIÇÃO"].iloc[i]
    linha_cat = plan_PBI["CATEGORIA"].iloc[i]
    linha_sub = plan_PBI["SUBCATEGORIA"].iloc[i]
    linha_val = plan_PBI["VALOR"].iloc[i]
    linha_pr = plan_PBI["PR"].iloc[i]


    if str(linha_data)[:3] == "MÊS":

        plan_PBI["DATA"] = plan_PBI["DATA"].apply(lambda x: str(x).replace(str(linha_data), '') if str(x).startswith(str(linha_data)) else x)
        plan_PBI["FAM"] = plan_PBI["FAM"].apply(lambda x: str(x).replace(str(linha_fam), '') if str(x).startswith(str(linha_fam)) else x)
        plan_PBI["ITEM"] = plan_PBI["ITEM"].apply(lambda x: str(x).replace(str(linha_item), '') if str(x).startswith(str(linha_item)) else x)
        plan_PBI["DESCRIÇÃO"] = plan_PBI["DESCRIÇÃO"].apply(lambda x: str(x).replace(str(linha_desc), '') if str(x).startswith(str(linha_desc)) else x)
        plan_PBI["CATEGORIA"] = plan_PBI["CATEGORIA"].apply(lambda x: str(x).replace(str(linha_cat), '') if str(x).startswith(str(linha_cat)) else x)
        plan_PBI["SUBCATEGORIA"] = plan_PBI["SUBCATEGORIA"].apply(lambda x: str(x).replace(str(linha_sub), '') if str(x).startswith(str(linha_sub)) else x)
        plan_PBI["VALOR"] = plan_PBI["VALOR"].apply(lambda x: str(x).replace(str(linha_val), '') if str(x).startswith(str(linha_val)) else x)
        plan_PBI["PR"] = plan_PBI["PR"].apply(lambda x: str(x).replace(str(linha_pr), '') if str(x).startswith(str(linha_pr)) else x)


#Filtrando #REF nas Descrição
for j in range(len(plan_PBI["DESCRIÇÃO"])):
    linha_data = plan_PBI["DATA"].iloc[j]
    linha_fam = plan_PBI["FAM"].iloc[j]
    linha_item = plan_PBI["ITEM"].iloc[j]
    linha_desc = plan_PBI["DESCRIÇÃO"].iloc[j]
    linha_cat = plan_PBI["CATEGORIA"].iloc[j]
    linha_sub = plan_PBI["SUBCATEGORIA"].iloc[j]
    linha_val = plan_PBI["VALOR"].iloc[j]
    linha_pr = plan_PBI["PR"].iloc[j]


    
    if str(linha_val)[:4] == "#REF!":

        plan_PBI["DATA"] = plan_PBI["DATA"].apply(lambda x: str(x).replace(str(linha_data), '') if str(x).startswith(str(linha_data)) else x)
        plan_PBI["FAM"] = plan_PBI["FAM"].apply(lambda x: str(x).replace(str(linha_fam), '') if str(x).startswith(str(linha_fam)) else x)
        plan_PBI["ITEM"] = plan_PBI["ITEM"].apply(lambda x: str(x).replace(str(linha_item), '') if str(x).startswith(str(linha_item)) else x)
        plan_PBI["DESCRIÇÃO"] = plan_PBI["DESCRIÇÃO"].apply(lambda x: str(x).replace(str(linha_desc), '') if str(x).startswith(str(linha_desc)) else x)
        plan_PBI["CATEGORIA"] = plan_PBI["CATEGORIA"].apply(lambda x: str(x).replace(str(linha_cat), '') if str(x).startswith(str(linha_cat)) else x)
        plan_PBI["SUBCATEGORIA"] = plan_PBI["SUBCATEGORIA"].apply(lambda x: str(x).replace(str(linha_sub), '') if str(x).startswith(str(linha_sub)) else x)
        plan_PBI["VALOR"] = plan_PBI["VALOR"].apply(lambda x: str(x).replace(str(linha_val), '') if str(x).startswith(str(linha_val)) else x)
        plan_PBI["PR"] = plan_PBI["PR"].apply(lambda x: str(x).replace(str(linha_pr), '') if str(x).startswith(str(linha_pr)) else x)



#Somando valor da Disp. Direta
# filtro = plan_PBI.loc[plan_PBI["SUBCATEGORIA"] == "Despesa Direta"]
# valor_dd = round((filtro["VALOR"].sum()),2)

