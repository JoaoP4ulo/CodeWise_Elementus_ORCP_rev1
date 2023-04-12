import datetime
import pandas as pd
import os

hoje = datetime.date.today()
dia_da_semana = hoje.weekday()
now = datetime.datetime.now()
dias_da_semana = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado', 'Domingo']
dia_da_semana_str = dias_da_semana[dia_da_semana]
data_hora = now.strftime("%d/%m/%Y %H:%M:%S")



path = 'C:\\Users\\jppel\Desktop\\ELEMENTUS\\ORCP_Python\\planilhas'


arquivos = os.listdir(path)
listagem_1 = [(os.path.join(path, arquivo), os.path.getmtime(os.path.join(path, arquivo))) for arquivo in arquivos]
listagem_1 = sorted(listagem_1, key=lambda arquivo: arquivo[1], reverse=True)
listagem_1 = pd.DataFrame(listagem_1)
listagem_1["2"] = data_hora
listagem_1.to_csv('historico.txt', sep=';', index=False)
