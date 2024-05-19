from datetime import datetime as dt
import pandas as pd
import sys

df_atual = pd.read_csv(sys.argv[1], sep=';',encoding='iso-8859-1')
ano =  sys.argv[2]
mes = sys.argv[3]

print('Processando a planilha do SimplesVet')

df_filter = df_atual.drop_duplicates(subset=['Animal'], keep="first")
df_filter = df_filter.drop(df_filter[df_filter['Responsável'] == 'Catland - Tratamento via Catland'].index)

total = len(df_filter['Animal'])
total_abrigo = len(df_filter[(df_filter['Responsável'] == 'Catland - Sede') | (df_filter['Responsável'] == 'Hospital Veterinário 4 Patas') | (df_filter['Responsável'] == 'Hospital Veterinário DC Clínica')])

print("Totais:\n")
print(df_filter['Responsável'].value_counts())
print('================================================================================')
print("Total de gatinhos: " + str(total))

print("Total de gatinhos em Abrigo: " + str(total_abrigo))
print("Total de gatinhos em LT: " +  str(total - total_abrigo))
