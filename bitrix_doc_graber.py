from fast_bitrix24 import Bitrix
import json
import datetime
from datetime import datetime
import re
from dock_info_second import *
import pandas as pd
# import logging

# logging.getLogger('fast_bitrix24').addHandler(logging.StreamHandler())
webhook = webhook_1
b = Bitrix(webhook)

# Получение сделки по id
def get_deal_id(n):
    deal = b.get_by_ID('crm.deal.get',[n])[f'{n}']
    return deal

# Получение контакта по id
def get_deal_contact(n):
    if n == None:
        return 'Без Имени'
    else:
        contact = b.get_by_ID('crm.contact.list',[n])[0]['NAME']
        return contact

# Получение всех сделок
def get_all_deals():
    deals = b.get_all('crm.deal.list')
    for i in deals:
        deal_title = i['TITLE']
        deal_sum = i['OPPORTUNITY']
        deal_name = get_deal_contact(i['CONTACT_ID'])
        print(f'\nНазвание: {deal_title}\nСумма сделки: {deal_sum}\nИмя: {deal_name}')

def site_get_all_deals():
    deals = b.get_all('crm.deal.list')
    arr_deadl = []
    for i in deals:
        frag = {}
        deal_title = i['TITLE']
        deal_sum = i['OPPORTUNITY']
        deal_name = get_deal_contact(i['CONTACT_ID'])
        frag = {
            'TITLE': deal_title,
            'SUMM': deal_sum,
            'CONTACT': deal_name
        }
        arr_deadl.append(frag)
    return arr_deadl

# Получение документов сделки по id
def incorrect_all_doc_deal(n):
    doc_info = b.get_by_ID('crm.documentgenerator.document.list',[n])                       # Запрос списка документов по сделке
    deal_title = get_deal_id(n)                                                             # Запрос названия сделки
    deal_title = deal_title['TITLE']
    print(f'\nРеестр документов для сделки \"{deal_title}\"\nКол-во документов: ',end='')
    # print(doc_info)
    doc_info = doc_info['6']['documents']                                                   # Вход в массив для корректной переборки
    print(len(doc_info))
    for doc in doc_info:                                                                    # Переборка документов для вывода
        # print(doc)
        doc_title = doc['title']
        doc_create = doc['createTime']
        doc_create = re.split('[- .:T+]+', doc_create)                                      # Разделение даты для последующей сборки в привычный вид
        doc_create = (f'{doc_create[2]}.{doc_create[1]}.{doc_create[0]} {doc_create[3]}:{doc_create[4]}')
        doc_url = doc['downloadUrl']
        print(f'\nНазвание документа: {doc_title}\nДата создания: {doc_create}\nСсылка на скачивание: {doc_url}')

# Получение документов сделки по id
def get_doc_deal(n):
    doc_info = b.get_by_ID('crm.documentgenerator.document.list', {'ID_list':[n]}, params={'filter': {'entityTypeId': '2','entityId': [n]}})
    deal_title = get_deal_id(n)
    # print(deal_title)
    deal_title = deal_title['TITLE']
    print(f'\nРеестр документов для сделки \"{deal_title}\"\nКол-во документов: ',end='')
    # print(doc_info)
    doc_info = doc_info['ID_list']['documents']
    print(len(doc_info))
    for doc in doc_info:
        doc_title = doc['title']
        doc_create = doc['createTime']
        doc_create = re.split('[- .:T+]+', doc_create)
        doc_create = (f'{doc_create[2]}.{doc_create[1]}.{doc_create[0]} {doc_create[3]}:{doc_create[4]}')
        doc_url = doc['downloadUrl']
        print(f'\nНазвание документа: {doc_title}\nДата создания: {doc_create}\nСсылка на скачивание: {doc_url}')
        
        xlsx_deal_id.append(n)
        xlsx_deal_title.append(deal_title)
        xlsx_doc_title.append(doc_title)
        xlsx_doc_date.append(doc_create)
        xlsx_url.append(doc_url)


def get_all_doc_deal():
    # Получение массива id сделок
    deals = b.get_all('crm.deal.list')
    arr_deal_id = []
    for deal in deals:
        deal_id = deal['ID']
        arr_deal_id.append(deal_id)

    # Получение док-ов по каждой сделке
    for i in arr_deal_id:
        get_doc_deal(i)

    #Создание xlsx
    df = pd.DataFrame({'ID Сделки': xlsx_deal_id,
                       'Название сделки': xlsx_deal_title, 
                       'Название документа': xlsx_doc_title, 
                       'Дата создания': xlsx_doc_date, 
                       'Ссылка на документ': xlsx_url                  })
    df.to_excel(f'C:/localhost/test_projects/bitrix_doc_graber/Реестр документов.xlsx', sheet_name='Документы', index=False)
    print('xlsx создался')

# def testttt():
#     with open('C:/localhost/test_projects/bitrix_doc_graber/log.json','a') as file:
#         file.write('test1')
    # data_json = {
    #         'ID Сделки': {n},
    #         'Название сделки': {deal_title},
    #         'Название документа': {doc_title},
    #         'Дата создания': {doc_create},
    #         'Ссылка на скачивание': {doc_url}
    #     }
    #     with open('C:/localhost/test_projects/bitrix_doc_graber/log.txt','a') as file:
    #         json.dump(data_json,file)

if __name__ == "__main__":
    xlsx_deal_id = []
    xlsx_deal_title = []
    xlsx_doc_title = []
    xlsx_doc_date = []
    xlsx_url = []
    # y = json.dumps(client, indent = 4)
    # y = json.loads(y)

    # get_doc_deal(6)
    get_all_doc_deal()

    # testttt()

    # r = site_get_all_deals()
    # print(r)
