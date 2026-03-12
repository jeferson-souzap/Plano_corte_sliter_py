import pandas as pd
import os
import platform

SO = platform.system()

if SO == 'Windows':
    BASE_INPUT = r'C:\Users\marce\OneDrive\Documentos\GitHub\plano_corte\input'
    BASE_OUTPUT = r'C:\Users\marce\OneDrive\Documentos\GitHub\plano_corte\output'
elif SO == 'Linux':
    BASE_INPUT = r'/home/stark/Documentos/Dev/Plano_corte_py/files/input'
    BASE_OUTPUT = r'/home/stark/Documentos/Dev/Plano_corte_py/files/output'
    

else:
    raise Exception('Sistema Operacional não suportado')


# Importando base do plano de corte
db_plano_corte = pd.read_excel(os.path.join(BASE_INPUT, 'db_plano_corte.xlsx'))

# Variaveis para os Calculos

# Coluna "Soma dos cortes (mm)" = somatório(coluna "Desenvolvimento" * coluna "Número de corte")

# plano_corte = soma dos valores encontrados na coluna "Soma dos cortes (mm)"

# perda_milimetros = Largura da bobina - plano_corte


# máximo 1,7%
#perda_percentual = perda_milimetros / Largura da bobina

#Regra de negócio.

'''
Usuário escolhe os parametros:
espessura, tipo de material e uma matriz

o script precisa retornar todas as combinações de matrizes da mesma espessura e tipo, que consiga manter o percentual de perda entre (0,67% e 1,7%)


'''

