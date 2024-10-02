import csv
import openpyxl
import os
import pandas as pd
from datetime import datetime, timedelta
from loginAutomatico import SapGui

wb = openpyxl.Workbook()
ws = wb.active

wb2 = openpyxl.Workbook()
ws2 = wb2.active

primeiro_dia_mes_atual = datetime.today().replace(day=1).date()

def last_day_of_month(any_day):
    # The day 28 exists in every month. 4 days later, it's always next month
    next_month = any_day.replace(day=28) + timedelta(days=4)
    # subtracting the number of the current day brings us back one month
    return next_month - timedelta(days=next_month.day)

primeira_data_mes_seguinte = last_day_of_month(primeiro_dia_mes_atual)+timedelta(days=1)

primeira_data_antecipacao = last_day_of_month(primeira_data_mes_seguinte)+timedelta(days=1)

SAP = SapGui()

SAP.SapLogin()

SAP.SapGetBAR()

with open(r"C:\Users\s-Lucas.SAraujo\ZBWINPL007_GESTÃO_DOS_ESTOQUES_20220623.csv") as f:
    reader = csv.reader(f, delimiter='\t')
    for row in reader:
        ws.append(row)

wb.save(r"C:\Users\s-Lucas.SAraujo\ZBWINPL007_GESTÃO_DOS_ESTOQUES_20220623.xlsx")

bar_df = pd.read_excel(r"C:\Users\s-Lucas.SAraujo\ZBWINPL007_GESTÃO_DOS_ESTOQUES_20220623.xlsx", engine='openpyxl')

bar_df['kg'] = bar_df['kg'].str.replace(',','.')
bar_df['cq'] = bar_df['cq'].str.replace(',','.')
bar_df['devolucao'] = bar_df['devolucao'].str.replace(',','.')
bar_df['producao'] = bar_df['producao'].str.replace(',','.')
bar_df['livre'] = bar_df['livre'].str.replace(',','.')

bar_df['kg'] = pd.to_numeric(bar_df['kg'])
bar_df['cq'] = pd.to_numeric(bar_df['cq'])
bar_df['devolucao'] = pd.to_numeric(bar_df['devolucao'])
bar_df['producao'] = pd.to_numeric(bar_df['producao'])
bar_df['livre'] = pd.to_numeric(bar_df['livre'])

bar_df['data_desejada_cliente'] = pd.to_datetime(bar_df['data_desejada_cliente'])

bar_df['data_desejada_cliente'] = bar_df['data_desejada_cliente'].dt.date

def condicao(valor):
    if valor < primeiro_dia_mes_atual:
        return 'Lay backlog'
    elif valor >= primeiro_dia_mes_atual and valor <= last_day_of_month(primeiro_dia_mes_atual):
        return 'Mês corrente'
    elif valor >= primeira_data_mes_seguinte and valor <= last_day_of_month(primeira_data_mes_seguinte):
        return 'Mês seguinte'
    elif valor >= primeira_data_antecipacao:
        return 'Antecipação'
    else:
        return 'Sobras'

# Adicionando uma coluna com dados condicionais usando apply
bar_df['Classificacao'] = bar_df['data_desejada_cliente'].apply(condicao)

bar_df_2 = pd.pivot_table(bar_df, values=['kg','cq','livre','producao','devolucao'],index='Classificacao',aggfunc='sum')

ordem_das_colunas = ['kg', 'cq', 'livre', 'producao', 'devolucao']

bar_df_2 = bar_df_2[ordem_das_colunas]

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(r"C:\Users\s-Lucas.SAraujo\Produto Final.xlsx", engine='xlsxwriter')

# Write each dataframe to a different worksheet. you could write different string like above if you want
bar_df.to_excel(writer, sheet_name='Barreiro - Base')
bar_df_2.to_excel(writer, sheet_name='Barreiro - Resumo')

os.remove(r"C:\Users\s-Lucas.SAraujo\ZBWINPL007_GESTÃO_DOS_ESTOQUES_20220623.csv")

os.remove(r"C:\Users\s-Lucas.SAraujo\ZBWINPL007_GESTÃO_DOS_ESTOQUES_20220623.xlsx")

SAP.SapGetJCB()

with open(r"C:\Users\s-Lucas.SAraujo\ZBWINPL007_GESTÃO_DOS_ESTOQUES_20220623.csv") as f:
    reader = csv.reader(f, delimiter='\t')
    for row in reader:
        ws2.append(row)

wb2.save(r"C:\Users\s-Lucas.SAraujo\ZBWINPL007_GESTÃO_DOS_ESTOQUES_20220623.xlsx")

jcb_df = pd.read_excel(r"C:\Users\s-Lucas.SAraujo\ZBWINPL007_GESTÃO_DOS_ESTOQUES_20220623.xlsx", engine='openpyxl')

jcb_df['kg'] = jcb_df['kg'].str.replace(',','.')
jcb_df['cq'] = jcb_df['cq'].str.replace(',','.')
jcb_df['devolucao'] = jcb_df['devolucao'].str.replace(',','.')
jcb_df['producao'] = jcb_df['producao'].str.replace(',','.')
jcb_df['livre'] = jcb_df['livre'].str.replace(',','.')

jcb_df['kg'] = pd.to_numeric(jcb_df['kg'])
jcb_df['cq'] = pd.to_numeric(jcb_df['cq'])
jcb_df['devolucao'] = pd.to_numeric(jcb_df['devolucao'])
jcb_df['producao'] = pd.to_numeric(jcb_df['producao'])
jcb_df['livre'] = pd.to_numeric(jcb_df['livre'])

jcb_df['data_desejada_cliente'] = pd.to_datetime(jcb_df['data_desejada_cliente'])

jcb_df['data_desejada_cliente'] = jcb_df['data_desejada_cliente'].dt.date

def condicao(valor):
    if valor < primeiro_dia_mes_atual:
        return 'Lay backlog'
    elif valor >= primeiro_dia_mes_atual and valor <= last_day_of_month(primeiro_dia_mes_atual):
        return 'Mês corrente'
    elif valor >= primeira_data_mes_seguinte and valor <= last_day_of_month(primeira_data_mes_seguinte):
        return 'Mês seguinte'
    elif valor >= primeira_data_antecipacao:
        return 'Antecipação'
    else:
        return 'Sobras'

# Adicionando uma coluna com dados condicionais usando apply
jcb_df['Classificacao'] = jcb_df['data_desejada_cliente'].apply(condicao)

jcb_df_2 = pd.pivot_table(jcb_df, values=['kg','cq','livre','producao','devolucao'],index='Classificacao',aggfunc='sum')

ordem_das_colunas = ['kg', 'cq', 'livre', 'producao', 'devolucao']

jcb_df_2 = jcb_df_2[ordem_das_colunas]

# Write each dataframe to a different worksheet. you could write different string like above if you want
jcb_df.to_excel(writer, sheet_name='Jeceaba - Base')
jcb_df_2.to_excel(writer, sheet_name='Jeceaba - Resumo')

# Close the Pandas Excel writer and output the Excel file.
writer.save()

os.remove(r"C:\Users\s-Lucas.SAraujo\ZBWINPL007_GESTÃO_DOS_ESTOQUES_20220623.csv")
os.remove(r"C:\Users\s-Lucas.SAraujo\ZBWINPL007_GESTÃO_DOS_ESTOQUES_20220623.xlsx")

SAP.SAPQuit()