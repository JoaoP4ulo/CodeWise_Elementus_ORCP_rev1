import pandas as pd


frame = pd.read_csv('framework_PBI.txt', delimiter=';')
frame = pd.DataFrame(frame)

valor_dd = frame[frame['SUBCATEGORIA'] == 'Despesa Direta']

pr = valor_dd[valor_dd["PR"]=="PR396"]
pr = pd.DataFrame(pr[pr['VALOR']!=0])
soma = pr["VALOR"].sum()
print(soma)
