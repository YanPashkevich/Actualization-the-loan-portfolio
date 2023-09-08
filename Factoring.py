#!/usr/bin/env python
# coding: utf-8

# In[2]:


import pandas as pd
import sys
import numpy as np
from datetime import datetime, timedelta, date, time
import openpyxl
from xml.etree import ElementTree


### Берем название элемента из командной строки
element = ''
date_sys = ''

### Пробуем получить название файла переданной из командной строки
try:
    element = sys.argv[1]
    date_sys = sys.argv[2]
except ValueError:
    element = element
    date_sys = date_sys
element = 'Белмедпрепараты_КЗ_12_05_2022_14_15_07'
date_sys = '11_05_2022_14_44_34'

### Преобразуем полученную из командной строки дату
date_sys_day = int(date_sys[0:2])
date_sys_month = int(date_sys[3:5])
date_sys_year = int(date_sys[6:10])
date_sys_main = date(date_sys_year,date_sys_month,date_sys_day)

name_of_excel_file = "{}-факторинг.xlsx".format(element) 
path = "Z:/Credit_report/answers/{}.xml".format(element)
tree = ElementTree.parse(path)
root = tree.getroot()


### Таблицы DF ###
df_all = pd.DataFrame(columns=['Кредитор', 'Сумма по договору, тыс', 'Сумма текущей задолженности, тыс', 'Валюта', 'Дата погашения', 'Окончание выборки', 'stopcontract'])


contract_number = 0
contract_name = 0

perfomance_on_date = []
sum_list = []
name_currency = []

### Осаток задолженности
sum_debt_currency = []
name_debt_currency = []

### для stopcontract
stop_contract = []

### для просроченной задолженности (если она есть)
sum_arrears = []
name_arrears =[]

for element in root.iter('FactoringSellerContractList'):
    if str(element.tag) == 'FactoringSellerContractList':
        for child in element:
            for contract in child:


                ### Nahodim nomer conracta
                if str(contract.tag) == 'contractnumber':
                    contract_number = contract.text

                ### Nahodim tip credita    
                if str(contract.tag) == 'credittype':
                    for credittype in contract:
                        if str(credittype.tag) == 'nametype':
                            contract_name = credittype.text

                ### Nahodim datu pogasheniya
                if str(contract.tag) == 'PerformanceDateOnDate':
                    for child in contract:
                        if str(child.tag) == 'performancedate':
                            date_str = child.text
                            day = int(date_str[0:2])
                            month = int(date_str[3:5])
                            year = int(date_str[6:10])
                            date_s = date(year,month,day)
                            perfomance_on_date.append(date_s)

                ### Nahodim summy po dogovoru i imya valuty
                if str(contract.tag) == 'AmountOnDate':
                    for child in contract:      
                        if str(child.tag) == 'amount':
                            for amount in child:
                                ### nahodim summu
                                if str(amount.tag) == 'sum':
                                    summa_po_dogovoru = float(amount.text) / 1000
                                    summa_po_dogovoru = "%.2f" % summa_po_dogovoru
                                    summa_po_dogovoru = float(summa_po_dogovoru)
                                    sum_list.append(summa_po_dogovoru)
                                ### nagodim nazvanie valuty
                                if str(amount.tag) == 'currency':
                                    for currency in amount:
                                        if str(currency.tag) == 'namecurrency':
                                            name_currency.append(currency.text[0:3])
                                            
                                            
                                            
                                            
                                            
                ## Nahodim summy ZADOLZHENNOSTI i imya valuty ЕСЛИ ЕСТЬ ДРУГОЙ ТЭГ
                if str(contract.tag) == 'FactoringSellerTransaction':
                    for child in contract:      
                        if str(child.tag) == 'remainingdebt-Undisclosed':
                            for amount in child:
                                ### nahodim summu
                                if str(amount.tag) == 'sum':
                                    summa_po_dogovoru = float(amount.text) / 1000
                                    summa_po_dogovoru = "%.2f" % summa_po_dogovoru
                                    summa_po_dogovoru = float(summa_po_dogovoru)
                                    sum_debt_currency.append(summa_po_dogovoru)
                                ### nagodim nazvanie valuty
                                if str(amount.tag) == 'currency':
                                    for currency in amount:
                                        if str(currency.tag) == 'namecurrency':
                                            name_debt_currency.append(currency.text[0:3])
                                            
                                     
                                    
                ### Проверяем есть ли ПРОСРОЧКА
                if str(contract.tag) == 'FactoringSellerTransaction':
                    for child in contract:      
                        if str(child.tag) == 'LateFactoringSeller-Recourse':
                            for arrears in child:
                                ### nahodim summu
                                if str(arrears.tag) == 'rest':
                                    summa_po_dogovoru = float(arrears.text) / 1000
                                    summa_po_dogovoru = "%.2f" % summa_po_dogovoru
                                    summa_po_dogovoru = float(summa_po_dogovoru)
                                    sum_arrears.append(summa_po_dogovoru)
                                ### nagodim nazvanie valuty
                                if str(arrears.tag) == 'currency':
                                    for currency in arrears:
                                        if str(currency.tag) == 'namecurrency':
                                            name_arrears.append(currency.text[0:3])
                                            

                
                ### !!! Проверяем подходящий ли нам контракт !!!
                if str(contract.tag) == 'stopcontract':
                    for child in contract:
                        if str(child.tag) == 'stopdate':
                            print('!!! STOP_CONTRACT !!!', child.text)
                            date_str = child.text
                            day = int(date_str[0:2])
                            month = int(date_str[3:5])
                            year = int(date_str[6:10])
                            date_s = date(year,month,day)
                            stop_cont = date_s
                            stop_contract.append(stop_cont)

                    
            ### !!! ВТОРАЯ ПРОВЕРКА подходит ли нам контракт !!! 
            try:
                if perfomance_on_date[0] < date_sys_main and len(sum_arrears) == 0:
                #if perfomance_on_date[0] < date_sys_main:
                    contract_number = 0
                    contract_name = 0
            except IndexError:
                #print('contract_number -----', contract_number)
                contract_number = 0
                contract_name = 0
            
##################################################### Заносим в df ###########################################################    

            #if contract_number != 0 and contract_name != 0:
            if contract_number != 0 and contract_name != 0:

                ### Заносим в действующую таблицу ###
                # Находим длину получившегося df
                length_df = len(df_all)
                # Создаем еще одну строку с nan
                df_all = df_all.append({'Кредитор': np.nan}, ignore_index=True)
                
                df_all['Кредитор'].loc[0] = contract_number + ' ' + contract_name
                df_all['Сумма по договору, тыс'].loc[0] = sum_list[0]
                
                try:
                    
                    if len(sum_arrears) != 0:
                        promezhutochn = 'срочная задолженность - {} просроченная задолженность - {}'.format(sum_debt_currency[0], sum_arrears[0])
                        df_all['Сумма текущей задолженности, тыс'].loc[0] = promezhutochn
                        
                except IndexError:
                    df_all['Сумма текущей задолженности, тыс'].loc[0] = np.nan
                    
                df_all['Валюта'].loc[0] = name_currency[0]
                df_all['Дата погашения'].loc[0] = perfomance_on_date[0].strftime("%d/%m/%Y")
                df_all['Окончание выборки'].loc[0] = np.nan
                try:
                    df_all['stopcontract'].loc[length_df] = stop_contract[0].strftime("%d/%m/%Y")
                except IndexError:
                    pass


                print('номер контракта: ', contract_number)
                print('тип контракта: ', contract_name)
                print('дата погашения: ', perfomance_on_date[0].strftime("%d/%m/%Y"))
                print('сумма по договору: ', sum_list[0])
                print('валюта по договору: ', name_currency[0])
                print('задолженность: ', sum_debt_currency)
                print('название валют задолженности: ', name_debt_currency)
                print()

                perfomance_on_date = []
                sum_list = []
                name_currency = []
                stop_contract = []                                            
                sum_debt_currency = []
                name_debt_currency = []
                sum_arrears = []
                name_arrears =[]
                
            else:
                perfomance_on_date = []
                sum_list = []
                name_currency = []
                stop_contract = []
                sum_debt_currency = []
                name_debt_currency = []
                sum_arrears = []
                name_arrears =[]
    
    
#################################################### Cохраняем данные в Excel ###################################################
if df_all.empty:
    print('empty')
else:
    ### Убираем 0 на np.nan в сумме договора
    ind = 0
    for i in df_all['Сумма по договору, тыс']:
        if str(i) != 'nan':
            if int(i) == 0:
                df_all['Сумма по договору, тыс'].loc[ind] = np.nan
        ind += 1
        
    path_name = "Z:/Credit_report/answers/{}".format(name_of_excel_file)
    # создаем Excel writer чтобы использовать XlsxWriter как движок
    writer = pd.ExcelWriter(path_name, engine='xlsxwriter')
    #### Конвертируем DF месяца, дня как Excel объекты
    df_all.to_excel(writer, sheet_name='Лист1', index=False)
    #### Сохраняем
    writer.save()

