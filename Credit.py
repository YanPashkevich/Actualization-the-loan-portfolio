#!/usr/bin/env python
# coding: utf-8

# In[71]:


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
#date_sys = '31_12_2021_17_37_23'

### Преобразуем полученную из командной строки дату
date_sys_day = int(date_sys[0:2])
date_sys_month = int(date_sys[3:5])
date_sys_year = int(date_sys[6:10])
date_sys_main = date(date_sys_year,date_sys_month,date_sys_day)


name_of_excel_file = "{}-кредиты.xlsx".format(element) 
path = "Z:/Credit_report/answers/{}.xml".format(element)
#name_of_excel_file = "answer-кредиты.xlsx".format(element) 
#path = "Z:/Credit_report/answers/answer.xml".format(element)
tree = ElementTree.parse(path)
root = tree.getroot()


### Таблицы DF ###
df_all = pd.DataFrame(columns=['Кредитор', 'Сумма по договору, тыс', 'Сумма текущей задолженности, тыс', 'Валюта', 'Дата погашения', 'Окончание выборки', 'stopcontract'])



### Список номеров контракта в теге <contract>
contract_number = []
### Список типов контракта в теге <contract>
contract_name = []


########################### Промежуточные списки для валют по договору и задолженностей ########################################################
sum_list_dop = []
name_currency_dop = []

sum_debt_list_dop = []
name_debt_currency_dop = []
#############################################################################################################################


### Список сумм по договору в теге <contract>
sum_list = []
### Список валют по договору в теге <contract>
name_currency = []
### Список сумм ЗАДОЛЖЕННОСТИ по договору в теге <contract>
sum_debt_list = []
### Список валют ЗАДОЛЖЕННОСТИ по договору в теге <contract>
name_debt_currency = []
### Список сумм ЗАДОЛЖЕННОСТИ по договору в теге <contract>
sum_debt_list = []
### Список валют ЗАДОЛЖЕННОСТИ по договору в теге <contract>
name_debt_currency = []
### ВЕСЬ СПИСОК ВАЛЮТ в конкретном договоре
all_name_currency = []
### !!! Лист погашения (новый введенный элемент) !!! ###
pogash_list = []

### LATESUM СУММЫ в конкретном договоре
latesum_sum_list = []
### LATESUM НАЗВАНИЕ ВАЛЮТЫ в конкретном договоре
latesum_name_currency = []
### LATESUM_ALL СУММЫ в конкретном договоре
latesum_sum_list_ALL = []
### LATESUM_ALL НАЗВАНИЕ ВАЛЮТЫ в конкретном договоре
latesum_name_currency_ALL = []


### LATEPERCENT СУММЫ в конкретном договоре
latepercent_sum_list = []
### LATEPERCENT НАЗВАНИЕ ВАЛЮТЫ в конкретном договоре
latepercent_name_currency = []
### LATEPERCENT_ALL СУММЫ в конкретном договоре
latepercent_sum_list_ALL = []
### LATEPERCENT_ALL НАЗВАНИЕ ВАЛЮТЫ в конкретном договоре
latepercent_name_currency_ALL = []


### Дата окончания в теге <contract>
end_date = 0
### Дата погашения в теге <contract>
pogash_date = 0

### для stopcontract
stop_contract = []







### flags for сумма и проценты задолженности
flag_a = 0
flag_b = 0



for element in root.iter('CreditGroup'):
    if str(element.tag) == 'CreditGroup':
        for child in element:
            for child1 in child:
                for contract in child1:
                    
                    
                    
                    ### Nahodim nomer conracta
                    if str(contract.tag) == 'contractnumber':
                        contract_number.append(contract.text)
                    
                        
                    ### Nahodim tip credita    
                    if str(contract.tag) == 'credittype':
                        for credittype in contract:
                            if str(credittype.tag) == 'nametype':
                                contract_name.append(credittype.text)
                                
                    ### Nahodim datu okonchanie vyborki
                    if str(contract.tag) == 'CreditGrantingLastDate':
                        date_str = contract.text
                        day = int(date_str[0:2])
                        month = int(date_str[3:5])
                        year = int(date_str[6:10])
                        date_s = date(year,month,day)
                        end_date = date_s

                    ### Nahodim datu pogasheniya
                    if str(contract.tag) == 'PerformanceDateOnDate':
                        for child in contract:
                            if str(child.tag) == 'performancedate':  
                                date_str = child.text
                                day = int(date_str[0:2])
                                month = int(date_str[3:5])
                                year = int(date_str[6:10])
                                date_s = date(year,month,day)
                                pogash_date = date_s
                                pogash_list.append(pogash_date)
                                
                                
                    ### !!! Проверяем подходящий ли нам контракт !!!
                    if str(contract.tag) == 'stopcontract':
                        for child in contract:
                            if str(child.tag) == 'stopdate':
                                date_str = child.text
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
                                        sum_list_dop.append(summa_po_dogovoru)
                                    ### nagodim nazvanie valuty
                                    if str(amount.tag) == 'currency':
                                        for currency in amount:
                                            if str(currency.tag) == 'namecurrency':
                                                name_currency_dop.append(currency.text[0:3])                                                
                                                
                    ### Nahodim summy tekuschey zadolzhennosti
                    if str(contract.tag) == 'credittransaction':
                        for child in contract:
                            if str(child.tag) == 'remainingdebt':
                                for remainingdebt in child:
                                    ### nahodim summu
                                    if str(remainingdebt.tag) == 'sum':
                                        sum_list_debt1 = float(remainingdebt.text) / 1000
                                        sum_list_debt1 = "%.2f" % sum_list_debt1
                                        sum_list_debt1 = float(sum_list_debt1)
                                        sum_debt_list_dop.append(sum_list_debt1)
                                    ### nagodim nazvanie valuty
                                    if str(remainingdebt.tag) == 'currency':
                                        for currency in remainingdebt:
                                            if str(currency.tag) == 'namecurrency':
                                                name_debt_currency_dop.append(currency.text[0:3])
                                                
                                                
                                                
                                                
                                                
                                                
                                                
                    
                    ### Nahodim LATESUM
                    if str(contract.tag) == 'credittransaction':
                        for child in contract:
                            if str(child.tag) == 'latesum':
                                for latesum in child:
                                    ### nahodim summu
                                    if str(latesum.tag) == 'rest':
                                        sum_list_latesum = float(latesum.text) / 1000
                                        sum_list_latesum = "%.2f" % sum_list_latesum
                                        sum_list_latesum = float(sum_list_latesum)
                                        #if sum_list_latesum != 0:   
                                            #latesum_sum_list.append(sum_list_latesum)
                                            #flag_a = 1
                                        latesum_sum_list.append(sum_list_latesum)
                                    ### nagodim nazvanie valuty
                                    if str(latesum.tag) == 'currency':
                                        for currency in latesum:
                                            #if str(currency.tag) == 'namecurrency' and flag_a == 1:
                                            if str(currency.tag) == 'namecurrency' and flag_a == 1:
                                                latesum_name_currency.append(currency.text[0:3])
                                                #flag_a = 0
                                                
                     ### Nahodim LATEPERCENT
                    if str(contract.tag) == 'credittransaction':
                        for child in contract:
                            if str(child.tag) == 'latepercent':
                                for latepercent in child:
                                    ### nahodim summu
                                    if str(latepercent.tag) == 'rest':
                                        sum_list_latepercent = float(latepercent.text) / 1000
                                        sum_list_latepercent = "%.2f" % sum_list_latepercent
                                        sum_list_latepercent = float(sum_list_latepercent)
                                        if sum_list_latepercent != 0:
                                            latepercent_sum_list.append(sum_list_latepercent)
                                            flag_b = 1
                                    ### nagodim nazvanie valuty
                                    if str(latepercent.tag) == 'currency':
                                        for currency in latepercent:
                                            if str(currency.tag) == 'namecurrency' and flag_b == 1:
                                                latepercent_name_currency.append(currency.text[0:3])
                                                flag_b = 0
                                                
                                                
                                        ### !!! ВТОРАЯ ПРОВЕРКА подходит ли нам контракт !!!
                
                try:
                    pogash_date = pogash_list[0]
                    if pogash_date < date_sys_main and len(latesum_sum_list) == 0:
                        contract_number = 0
                        contract_name = 0
                except IndexError:
                    contract_number = 0
                    contract_name = 0

                                 
                                
                               

                                
                                
                                                
                                                
                ### Создаем список всех валют и сумм по договору
                a = 0
                for name_c in name_currency_dop:
                    
                    if name_c in name_currency:
                        a = a+1
                    else:
                        sum_list.append(sum_list_dop[a])
                        name_currency.append(name_currency_dop[a])
                        a = a+1
                        
                        
                ### Создаем список всех валют и сумм по договору
                a = 0
                for name_c in name_debt_currency_dop:
                    
                    if name_c in name_debt_currency:
                        a = a+1
                    else:
                        sum_debt_list.append(sum_debt_list_dop[a])
                        name_debt_currency.append(name_debt_currency_dop[a])
                        a = a+1
                        
                        
                        
                        
                        
                        
                        
                ### Создаем список всех валют и сумм по LATESUM
                a = 0
                for name_c in latesum_name_currency:
                    
                    if name_c in latesum_name_currency_ALL:
                        a = a+1
                    else:
                        latesum_sum_list_ALL.append(latesum_sum_list[a])
                        latesum_name_currency_ALL.append(latesum_name_currency[a])
                        a = a+1
                        
                        
                ### Создаем список всех валют и сумм по LATEPERCENT
                a = 0
                for name_c in latepercent_name_currency:
                    
                    if name_c in latepercent_name_currency_ALL:
                        a = a+1
                    else:
                        latepercent_sum_list_ALL.append(latepercent_sum_list[a])
                        latepercent_name_currency_ALL.append(latepercent_name_currency[a])
                        a = a+1     
                        

########################################## Проверяем подходящий ли это нам контракт ############################################                

                if contract_number != 0 and contract_name != 0:
                    
                    
                    ### Заносим в список все виды ВАЛЮТ в договоре
                    for name in name_currency:
                        if name in all_name_currency:
                            pass
                        else:
                            all_name_currency.append(name)
                    for name in name_debt_currency:
                        if name in all_name_currency:
                            pass
                        else:
                            all_name_currency.append(name)
                            
                    print('номер контракта - ', contract_number) 
                    print('задолженнонсти:', sum_debt_list)
                    print('задолженнонсти:',name_debt_currency)
                    print('все валюты:', all_name_currency)
                    
                            
                            
                            
                            
                                          
                    ### Вносим валюты и суммы       
                    for name in all_name_currency:
                        if name in name_currency:
                            
                            index = name_currency.index(name)
                            ### Если есть ТАКАЯ валюта в ЗАДОЛЖЕННОСТЯХ
                            if name in name_debt_currency:
                                
                                
                                index_debt = name_debt_currency.index(name)
                                length_df = len(df_all)
                                df_all = df_all.append({'Кредитор': np.nan}, ignore_index=True)
                                print('валюта есть в задолженностях', name, index_debt)
                                
                                df_all['Кредитор'].loc[length_df] = contract_number[0] + ' ' + contract_name[0]
                                df_all['Сумма по договору, тыс'].loc[length_df] = sum_list[index]
                                
                                
                                promezhutochn = str(sum_debt_list[index_debt])
                                print('промежуточн', promezhutochn)
                                ### Вносим по-новому данные о задолженностях
                                if name in latesum_name_currency:
                                    if len(latesum_sum_list) != 0:
                                        index_latesum = latesum_name_currency_ALL.index(name)
                                        index_percent = latepercent_name_currency_ALL.index(name)
                                        promezhutochn = 'срочная задолженность - ' + promezhutochn
                                        try:
                                            prosroch = ' просроченная задолженность - {} + {}'.format(latesum_sum_list[index_latesum], latepercent_sum_list[index_percent])
                                            promezhutochn = promezhutochn + prosroch
                                        except IndexError:
                                            prosroch = ' просроченная задолженность - 0'
                                            promezhutochn = promezhutochn + prosroch
                                
                                    
                                
                                
                                if str(promezhutochn) == '0.0':
                                    promezhutochn = 0
                                df_all['Сумма текущей задолженности, тыс'].loc[length_df] = promezhutochn
                                df_all['Валюта'].loc[length_df] = name
                                df_all['Дата погашения'].loc[length_df] = pogash_date.strftime("%d/%m/%Y")
                                df_all['Окончание выборки'].loc[length_df] = end_date.strftime("%d/%m/%Y")
                                try:
                                    df_all['stopcontract'].loc[length_df] = stop_contract[0].strftime("%d/%m/%Y")
                                except IndexError:
                                    pass
                            ### Если такой валюты нет
                            else:
                                length_df = len(df_all)
                                df_all = df_all.append({'Кредитор': np.nan}, ignore_index=True)
                                
                                df_all['Кредитор'].loc[length_df] = contract_number[0] + ' ' + contract_name[0]
                                df_all['Сумма по договору, тыс'].loc[length_df] = sum_list[index]
                                df_all['Сумма текущей задолженности, тыс'].loc[length_df] = '0'
                                
                                
                                
                                
                                ### Вносим по-новому данные о задолженностях
                                if name in latesum_name_currency:
                                    if len(latesum_sum_list) != 0:
                                        
                                        index_latesum = latesum_name_currency_ALL.index(name)
                                        index_percent = latepercent_name_currency_ALL.index(name)
                                        promezhutochn = 'срочная задолженность - ' + promezhutochn
                                        try:
                                            prosroch = ' просроченная задолженность - {} + {}'.format(latesum_sum_list[index_latesum], latepercent_sum_list[index_percent])
                                            promezhutochn = promezhutochn + prosroch
                                        except IndexError:
                                            prosroch = ' просроченная задолженность - 0'
                                            promezhutochn = promezhutochn + prosroch
                                    if str(promezhutochn) == '0.0':
                                        promezhutochn = 0
                                    df_all['Сумма текущей задолженности, тыс'].loc[length_df] = promezhutochn
                                
                                

                                df_all['Валюта'].loc[length_df] = name
                                df_all['Дата погашения'].loc[length_df] = pogash_date.strftime("%d/%m/%Y")
                                df_all['Окончание выборки'].loc[length_df] = end_date.strftime("%d/%m/%Y")
                                try:
                                    df_all['stopcontract'].loc[length_df] = stop_contract[0].strftime("%d/%m/%Y")
                                except IndexError:
                                    pass
                                
                        ### Если нет такой валюты по ДОГОВРУ начинаем смотреть валюту в ЗАДОЛЖЕННОСТЯХ
                        elif name in name_debt_currency:
                            index_debt = name_debt_currency.index(name)
                            print('индекс валюты - ', name, 'в задолженностях - ',index_debt)
                            
                            length_df = len(df_all)
                            df_all = df_all.append({'Кредитор': np.nan}, ignore_index=True)
                                
                            df_all['Кредитор'].loc[length_df] = contract_number[0] + ' ' + contract_name[0]
                            df_all['Сумма по договору, тыс'].loc[length_df] = np.nan
                            promezhutochn = sum_debt_list[index_debt]
                            print('')
                            
                                
                            ### Вносим по-новому данные о задолженностях
                            if name in latesum_name_currency:
                                
                                if len(latesum_sum_list) != 0:
                                    index_latesum = latesum_name_currency_ALL.index(name)
                                    index_percent = latepercent_name_currency_ALL.index(name)
                                    promezhutochn = 'срочная задолженность - ' + promezhutochn
                                    try:
                                        prosroch = ' просроченная задолженность - {} + {}'.format(latesum_sum_list[index_latesum], latepercent_sum_list[index_percent])
                                        promezhutochn = promezhutochn + prosroch
                                    except IndexError:
                                        prosroch = ' просроченная задолженность - 0'
                                        promezhutochn = promezhutochn + prosroch
                            
                            if str(promezhutochn) == '0.0':
                                    promezhutochn = 0
                            df_all['Сумма текущей задолженности, тыс'].loc[length_df] = promezhutochn
                            df_all['Валюта'].loc[length_df] = name
                            df_all['Дата погашения'].loc[length_df] = pogash_date.strftime("%d/%m/%Y")
                            df_all['Окончание выборки'].loc[length_df] = end_date.strftime("%d/%m/%Y")
                            try:
                                df_all['stopcontract'].loc[length_df] = stop_contract[0].strftime("%d/%m/%Y")
                            except IndexError:
                                pass
                            
                            
                            
                            
                            
                            
                     

                    
                    contract_number = []
                    contract_name = []
                    sum_list = []
                    name_currency = []
                    sum_list_dop = []
                    name_currency_dop = []
                    sum_debt_list_dop = []
                    name_debt_currency_dop = []
                    sum_debt_list = []
                    name_debt_currency = []
                    all_name_currency = []
                    latesum_sum_list = []
                    latesum_name_currency = []
                    latepercent_sum_list = []
                    latepercent_name_currency = []
                    pogash_list = []
                    
                    latesum_sum_list_ALL = []
                    latesum_name_currency_ALL = []
                    latepercent_sum_list_ALL = []
                    latepercent_name_currency_ALL = []
                    stop_contract = []

                else:
                    contract_number = []
                    contract_name = []
                    sum_list = []
                    name_currency = []
                    sum_list_dop = []
                    name_currency_dop = []
                    sum_debt_list_dop = []
                    name_debt_currency_dop = []
                    sum_debt_list = []
                    name_debt_currency = []
                    all_name_currency = []
                    latesum_sum_list = []
                    latesum_name_currency = []
                    latepercent_sum_list = []
                    latepercent_name_currency = []
                    pogash_list = []
                    
                    latesum_sum_list_ALL = []
                    latesum_name_currency_ALL = []
                    latepercent_sum_list_ALL = []
                    latepercent_name_currency_ALL = []
                    stop_contract = []
                    
                print()
#################################################### Cохраняем данные в Excel ###################################################
if df_all.empty:
    print('empty')
else:
    ### Убираем 0 на np.nan в сумме договора
    ind = 0
    for i in df_all['Сумма по договору, тыс']:
    ### Убираем 0 на np.nan в сумме договора
        if str(i) != 'nan':
            if int(i) == 0:
                df_all['Сумма по договору, тыс'].loc[ind] = np.nan
        ind += 1
        
    path_name = "Z:/Credit_report/answers/{}".format(name_of_excel_file)
    # создаем Excel writer чтобы использовать XlsxWriter как движок
    writer = pd.ExcelWriter(path_name, engine='xlsxwriter')
    ### Конвертируем DF месяца, дня как Excel объекты
    df_all.to_excel(writer, sheet_name='Лист1', index=False)
    ### Сохраняем
    writer.save()#


# In[72]:


df_all


# In[ ]:




