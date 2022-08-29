import pandas as pd
from fast_bitrix24 import Bitrix
import json
import time

### Work with data

data = pd.read_excel (r'C:\Users\Admin\Desktop\Analitics\Season\Test.xlsx')
data = data.values.tolist()
deadline = '07/25/2022'
title = 'Летнее колесо фортуны!'


### Work with bitrix

webhook = (webhook)
b = Bitrix(webhook)

for i in data: 
    
    ### Create a tasks

    count = str(i[2])
    date_last_deal = str(i[3])
    company_title = i[4]
    amount_obj = str(i[5])
    description = 'Создана задача на прозвон летнего колеса фортуны. \n\n' + 'Большой шанс на заключение сделки с этой компанией в летний период. \n\n' + 'Название компании: ' + company_title + '\nДата последней сделки: ' + date_last_deal + '\nКоличество сделок летом: ' + count + '\n Количество объектов: ' + amount_obj

    method = 'tasks.task.add'
    params = {
        'fields': {
            'TITLE': title,
            'DESCRIPTION': description,
            'CREATED_BY': 2166,
            'RESPONSIBLE_ID': i[1],
            'DEADLINE': deadline,
            'UF_CRM_TASK': i[0]        
        }
    }
    res = b.call(method,params)
  
    # print('Задача на ' + i[1] + ' создана' )

    res = json.dumps(res)
    res = json.loads(res)
    task_id = res['task']['id']
    resp_id = res['task']['responsibleId']

    ### Update a task

    method = 'tasks.task.update'
    params = {
        'taskId': task_id, 
        'fields[UF_CRM_TASK][]': i[0]
        }
    
    b.call(method,params)

    # print('Задача прикреплена к компании ' + i[0])

    ### Send a message

    url = ' https://portal.stavtrack.ru/company/personal/user/0/tasks/task/view/'
    message = 'Создана задача "Летнее Колесо Фортуны", все подробности в задаче. \n\n' + 'Ссылка на задачу: ' + url + task_id + "/"

    method = 'im.message.add'
    params = {
            'DIALOG_ID': resp_id,
            'MESSAGE': message
        }
        
    b.call(method,params)
    
    # print('Сообщение с ссылкой на задачу ответственному отправлено!')

    time.sleep(0.3)
