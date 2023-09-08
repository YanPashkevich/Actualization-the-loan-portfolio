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
#element = '7_1_Кредитная_заявка_БелОМО_1'
#date_sys = '31/12/2021_17:37:23'

### Преобразуем полученную из командной строки дату
date_sys_day = int(date_sys[0:2])
date_sys_month = int(date_sys[3:5])
date_sys_year = int(date_sys[6:10])
date_sys_main = date(date_sys_year,date_sys_month,date_sys_day)

name_of_excel_file = "{}-поручительства.xlsx".format(element) 
path = "Z:/Credit_report/answers/{}.xml".format(element)
tree = ElementTree.parse(path)
root = tree.getroot()



### Таблицы DF ###
df_all = pd.DataFrame(columns=['Кредитор', 'Сумма по договору, тыс', 'Сумма текущей задолженности, тыс', 'Валюта', 'Дата погашения', 'Окончание выборки', 'stopcontract'])


### Переменная для хранения номера контракта
contract_number = 0
### Переменная для хранения типа контракта
contract_name = 0

### Список для хранения даты погашения контракта
perfomance_on_date = []
### Список для хранения суммы контракта
sum_list = []
### Список для хранения валюты контракта
name_currency = []

### для stopcontract
stop_contract = []

for element in root.iter('SuretyContractList'):
    if str(element.tag) == 'SuretyContractList':
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
                            
                ### !!! Проверяем подходящий ли нам контракт !!!
                if str(contract.tag) == 'stopdate':
                        date_str = contract.text
                        day = int(date_str[0:2])
                        month = int(date_str[3:5])
                        year = int(date_str[6:10])
                        date_s = date(year,month,day)
                        stop_cont = date_s
                        stop_contract.append(stop_cont)
                            
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
                            

            #sum_list = []
            #perfomance_on_date = []
            #name_currency = []
            
##################################################### Заносим в df ###########################################################    

            if contract_number != 0 and contract_name != 0:
        
                print(contract_number)
                print(contract_name)
                print(perfomance_on_date[0])
                print(sum_list[0])
                print(name_currency[0])
                print()
                
                
                ### Заносим в действующую таблицу ###
                # Находим длину получившегося df
                length_df = len(df_all)
                # Создаем еще одну строку с nan
                df_all = df_all.append({'Кредитор': np.nan}, ignore_index=True)
                
                df_all['Кредитор'].loc[length_df] = contract_number + ' ' + contract_name
                df_all['Сумма по договору, тыс'].loc[length_df] = sum_list[0]
                df_all['Сумма текущей задолженности, тыс'].loc[length_df] = sum_list[0]
                df_all['Валюта'].loc[length_df] = name_currency[0]
                df_all['Дата погашения'].loc[length_df] = perfomance_on_date[0].strftime("%d/%m/%Y")
                df_all['Окончание выборки'].loc[length_df] = np.nan
                
                try:
                    df_all['stopcontract'].loc[length_df] = stop_contract[0].strftime("%d/%m/%Y")
                except IndexError:
                    pass
                
            sum_list = []
            perfomance_on_date = []
            name_currency = []
            stop_contract = []
        
        
#################################################### Cохраняем данные в Excel ###################################################
if df_all.empty:
    print('empty')
else:
    path_name = "Z:/Credit_report/answers/{}".format(name_of_excel_file)
    # создаем Excel writer чтобы использовать XlsxWriter как движок
    writer = pd.ExcelWriter(path_name, engine='xlsxwriter')
    #### Конвертируем DF месяца, дня как Excel объекты
    df_all.to_excel(writer, sheet_name='Лист1', index=False)
    #### Сохраняем
    writer.save() 


# In[ ]:


df_all


# In[ ]:




