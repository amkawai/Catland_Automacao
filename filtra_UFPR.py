import pandas as pd
import sys

df_atual = pd.read_csv(sys.argv[1], sep=';',encoding='iso-8859-1')

print('Processando')

df_filter = df_atual[(df_atual['Animal - Status'] == 'Ativo') & (
(df_atual['Cliente - Nome'] == 'Catland - Cafofinho') | 
(df_atual['Cliente - Nome'] == 'Catland - LT Eterno') | 
(df_atual['Cliente - Nome'] == 'Catland - Quarentena Externa') | 
(df_atual['Cliente - Nome'] == 'Catland - Ronron Cat Café') | 
(df_atual['Cliente - Nome'] == 'Catland - Sede') | 
(df_atual['Cliente - Nome'] == 'Hospital Veterinário 4 Patas') | 
(df_atual['Cliente - Nome'] == 'Hospital Veterinário DC Clínica') | 
(df_atual['Cliente - Nome'] == 'Petz Alto da Boa Vista'))
]

# ordena pelo Cod SimplesVet
df_filter = df_filter.sort_values(by=['Animal - Código'])

# exclui a linha com AAAexemplo
df_filter = df_filter.drop(df_filter[df_filter['Animal - Nome'] == 'AAAexemplo'].index)

df_filter.to_csv('UFPR_filtrada.csv', sep=';')

total = len(df_filter)
total_abrigo = len(df_filter[(df_filter['Cliente - Nome'] == 'Catland - Sede') | (df_filter['Cliente - Nome'] == 'Hospital Veterinário 4 Patas') | (df_filter['Cliente - Nome'] == 'Hospital Veterinário DC Clínica')])

print("Totais:\n")
print(df_filter["Cliente - Nome"].value_counts())

print("Total de gatinhos: " + str(total))

print("Total de gatinhos em Abrigo: " + str(total_abrigo))
print("Total de gatinhos em LT: " +  str(total - total_abrigo))
