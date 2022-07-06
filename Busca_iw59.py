from datetime import datetime
import pandas as pd
import time
import openpyxl

nome_extração = ['IW59']
tempo_inicio = []
tempo_final = []
duracao = []

tempo_1 = datetime.now()

cont_1 = 5
while True:
    if cont_1 == 0:
        break
    time.sleep(1)
    cont_1 = cont_1 - 1

tempo_2 = datetime.now()

tempo_inicio.append(str(tempo_1))
print(tempo_inicio)
tempo_final.append(str(tempo_2))
print(tempo_final)

nome_extração.append("IW59")

duracao_time = str(tempo_2 - tempo_1)

print(duracao)

plan = openpyxl.load_workbook(r"C:\Users\**\PycharmProjects\pythonProject1\Relatório.xlsx")
plan_extraçoes = plan['Extrações']

for item in nome_extração:
    plan_extraçoes.append([nome_extração,tempo_inicio,tempo_final,duracao])

plan.save(r"C:\Users\**\PycharmProjects\pythonProject1\Relatório.xlsx")
