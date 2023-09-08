#!/usr/bin/env python
# coding: utf-8

# In[4]:


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
#element = 'Водоканал_7_1_Кредитная_заявка_v7_4_27_04_2022_6_00_04'
#date_sys = '29_04_2022_08_00_36'

### Преобразуем полученную из командной строки дату
date_sys_day = int(date_sys[0:2])
date_sys_month = int(date_sys[3:5])
date_sys_year = int(date_sys[6:10])
date_sys_main = date(date_sys_year,date_sys_month,date_sys_day)

name_of_excel_file = "{}-лизинг.xlsx".format(element) 
path = "Z:/Credit_report/answers/{}.xml".format(element)
tree = ElementTree.parse(path)
root = tree.getroot()


### Таблицы DF ###
df_all = pd.DataFrame(columns=['Кредитор', 'Сумма по договору, тыс', 'Сумма текущей задолженности, тыс', 'Валюта', 'Дата погашения', 'Окончание выборки', 'stopcontract'])


### Переменная для хранения номера конраткта
contract_number = 0
### Переменная для хранения типа конраткта
contract_name = 0


### Список для хранения даты погашения
perfomance_on_date = []
### Список для хранения суммы по договору
summa_po_dogovoru = []
### Список для хранения названий валют по договору
name_currency_summa_po_dogovoru = []
### Список для хранения сумм задолженностей по договору
sum_list_debt = []
### Список для хранения валют задолженностей по договору
name_currency_debt = []


### LATESUM СУММЫ в конкретном договоре
latesum_sum_list = []
### LATESUM НАЗВАНИЕ ВАЛЮТЫ в конкретном договоре
latesum_name_currency = []
### LATESUM_ALL СУММЫ в конкретном договоре
latesum_sum_list_ALL = []
### LATESUM_ALL НАЗВАНИЕ ВАЛЮТЫ в конкретном договоре
latesum_name_currency_ALL = []

### для stopcontract
stop_contract = []

### flag говорящий, что есть строка "нет"
flag_no = 0



for element in root.iter('LeasingContractList'):
    if str(element.tag) == 'LeasingContractList':
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
                
                            
                ### Nahodim summu po dogovory
                if str(contract.tag) == 'AmountOnDate':
                    for child in contract:
                        if str(child.tag) == 'amount':
                            for child1 in child: 
                                if str(child1.tag) == 'sum':
                                    summa_po_dogovoru1 = float(child1.text) / 1000
                                    summa_po_dogovoru1 = "%.2f" % summa_po_dogovoru1
                                    summa_po_dogovoru.append(float(summa_po_dogovoru1))
                                for currency in child1:
                                    if str(currency.tag) == 'namecurrency':
                                        name_currency_summa_po_dogovoru.append(currency.text[0:3])
                                        
                ### Nahodim summu po dogovory если Cost
                if str(contract.tag) == 'Cost' and 'EUR' not in name_currency_summa_po_dogovoru and 'USD' not in name_currency_summa_po_dogovoru and 'RUB' not in name_currency_summa_po_dogovoru:
                    flag_no = 1
                    #summa_po_dogovoru = []
                    #name_currency_summa_po_dogovoru = []
                    #for child in contract:
                    #    if str(child.tag) == 'sum':
                    #        summa_po_dogovoru1 = float(child.text) / 1000
                    #        summa_po_dogovoru1 = "%.2f" % summa_po_dogovoru1
                    #        summa_po_dogovoru.append(float(summa_po_dogovoru1))
                    #        
                    #    if str(child.tag) == 'currency':
                    #        
                    #        for currency in child:
                    #            if str(currency.tag) == 'namecurrency':
                    #                name_currency_summa_po_dogovoru.append(currency.text[0:3])
                                        
                ### Nahodim summu tekuschey zadolzhennosti
                if str(contract.tag) == 'LeasingTransaction':
                    for child in contract:
                        if str(child.tag) == 'remainingdebt':
                            for remainingdebt in child:
                                ### Zanosim summu v spisok
                                if str(remainingdebt.tag) == 'sum':
                                    x = float(remainingdebt.text) / 1000
                                    sum_dot2 = "%.2f" % x
                                    summa_debt = float(sum_dot2)
                                    sum_list_debt.append(summa_debt)
                                ### Zanosim nazvanie valuty v spisok
                                if str(remainingdebt.tag) == 'currency':
                                    for currency in remainingdebt:
                                        if str(currency.tag) == 'namecurrency':
                                            name_currency_debt.append(currency.text[0:3])
                                            
                        ### ВНОСИМ LATESUM !!! ###
                        if str(child.tag) == 'LateLeasingSum':
                            for latesum in child:
                                ### nahodim summu
                                if str(latesum.tag) == 'rest':
                                    sum_list_latesum = float(latesum.text) / 1000
                                    sum_list_latesum = "%.2f" % sum_list_latesum
                                    sum_list_latesum = float(sum_list_latesum)
                                    #print(sum_list_latesum)
                                    latesum_sum_list.append(sum_list_latesum)
                                ### nagodim nazvanie valuty
                                if str(latesum.tag) == 'currency':
                                    for currency in latesum:
                                        if str(currency.tag) == 'namecurrency':
                                            #print(currency.text[0:3])
                                            latesum_name_currency.append(currency.text[0:3])
                            
                            
                ### !!! Проверяем подходящий ли нам контракт !!!
                if str(contract.tag) == 'stopcontract':
                    for child in contract:
                        if str(child.tag) == 'stopdate':
                            #print('!!! STOP_CONTRACT !!!', child.text)
                            date_str = child.text
                            day = int(date_str[0:2])
                            month = int(date_str[3:5])
                            year = int(date_str[6:10])
                            date_s = date(year,month,day)
                            stop_cont = date_s
                            stop_contract.append(stop_cont)
                    
                                                                
            ### !!! ВТОРАЯ ПРОВЕРКА подходит ли нам контракт !!! 
            try:
                if perfomance_on_date[0] < date_sys_main and len(latesum_sum_list) == 0:
                    contract_number = 0
                    contract_name = 0
            except IndexError:
                #print('contract_number -----', contract_number)
                contract_number = 0
                contract_name = 0
                
            ### Создаем список всех валют и сумм по LATESUM
            a = 0
            for name_c in latesum_name_currency: 
                if name_c in latesum_name_currency_ALL:
                    a = a+1
                else:
                    latesum_sum_list_ALL.append(latesum_sum_list[a])
                    latesum_name_currency_ALL.append(latesum_name_currency[a])
                    a = a+1
                    
                    
########################################## Проверяем подходящий ли это нам контракт ############################################                
            
            if contract_number != 0 and contract_name != 0:
                

                
                ### Заносим в df
                length_df = len(df_all)
                df_all = df_all.append({'Кредитор': np.nan}, ignore_index=True)
                                
                df_all['Кредитор'].loc[length_df] = contract_number + ' ' + contract_name
                print(contract_number)
                print(summa_po_dogovoru)
                print(name_currency_summa_po_dogovoru)
                print('задолженность: ', sum_list_debt)
                print('задолженность ВАЛЮТЫ: ', name_currency_debt)
                print('flag_no = ', flag_no)
                print()
                
                ### Если есть flag_no = 1, то берем по сумму по договору с индексо 0 в списке, иначе с инд. 1
                if flag_no == 1:
                    df_all['Сумма по договору, тыс'].loc[length_df] = float(summa_po_dogovoru[0])
                if flag_no == 0:
                    try:
                        df_all['Сумма по договору, тыс'].loc[length_df] = float(summa_po_dogovoru[1])
                    except IndexError:
                        df_all['Сумма по договору, тыс'].loc[length_df] = float(summa_po_dogovoru[0])
                
                ### Если пустота в поле "Остаток задолженности" - ставим 0 
                try:
                    promezhutochn = str(sum_list_debt[0])
                except IndexError:
                    promezhutochn = '0.0'
                    
                
                
                if len(latesum_sum_list) != 0:
                    if len(latesum_name_currency_ALL) == 1 and latesum_name_currency_ALL[0] != 'BYN':
                        pass
                    else:
                        index_latesum = latesum_name_currency_ALL.index(name_currency_summa_po_dogovoru[0])
                        if latesum_sum_list[index_latesum] != 0:
                            promezhutochn = 'срочная задолженность - ' + promezhutochn
                            try:
                                prosroch = ' просроченная задолженность - {}'.format(latesum_sum_list[index_latesum])
                                promezhutochn = promezhutochn + prosroch
                            except IndexError:
                                ### просроченная задолженность == 0
                                pass
                                #prosroch = ' просроченная задолженность - 0'
                                #promezhutochn = promezhutochn + prosroch
                                            
                                            
                                            
                df_all['Сумма текущей задолженности, тыс'].loc[length_df] = promezhutochn
                valute = ''
                
                ### Если есть flag_no = 1, то берем валюту по договору с индексом 0 в списке, иначе с инд. 1
                if flag_no == 1:
                    valute = name_currency_summa_po_dogovoru[0]
                if flag_no == 0:
                    try:
                        valute = name_currency_summa_po_dogovoru[1]
                    except IndexError:
                        valute = name_currency_summa_po_dogovoru[0]
                
                df_all['Валюта'].loc[length_df] = valute
                    
                df_all['Дата погашения'].loc[length_df] = perfomance_on_date[0].strftime("%d/%m/%Y")
                df_all['Окончание выборки'].loc[length_df] = "-"
                try:
                    df_all['stopcontract'].loc[length_df] = stop_contract[0].strftime("%d/%m/%Y")
                except IndexError:
                    pass
                
                perfomance_on_date = []
                summa_po_dogovoru = []
                name_currency_summa_po_dogovoru = []
                sum_list_debt = []
                name_currency_debt = []
                latesum_sum_list = []
                latesum_name_currency = []
                latesum_sum_list_ALL = []
                latesum_name_currency_ALL = []
                stop_contract = []
                flag_no = 0
            else:
                perfomance_on_date = []
                summa_po_dogovoru = []
                name_currency_summa_po_dogovoru = []
                sum_list_debt = []
                name_currency_debt = []
                latesum_sum_list = []
                latesum_name_currency = []
                latesum_sum_list_ALL = []
                latesum_name_currency_ALL = []
                stop_contract = []
                flag_no = 0
                
                
#################################################### Cохраняем данные в Excel ###################################################
if df_all.empty:
    print('empty')
else:
    ### Убираем 0 на np.nan в сумме договора
    ind = 0
    for i in df_all['Сумма по договору, тыс']:
    ### Убираем 0 на np.nan в сумме договора
        if str(i) != 'nan':
            if float(i) == 0:
                df_all['Сумма по договору, тыс'].loc[ind] = np.nan
        ind += 1
        
    path_name = r"Z:/Credit_report/answers/{}".format(name_of_excel_file)
    # создаем Excel writer чтобы использовать XlsxWriter как движок
    writer = pd.ExcelWriter(path_name, engine='xlsxwriter')
    #### Конвертируем DF месяца, дня как Excel объекты
    df_all.to_excel(writer, sheet_name='Лист1', index=False)
    #### Сохраняем
    writer.save()    


# In[5]:


df_all


# In[ ]:





# In[ ]:




