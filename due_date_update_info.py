# -*- coding: utf-8 -*-


import os
from datetime import datetime
from pprint import pprint


import shotgun_api3 as sa
import pandas as pd


SG = sa.Shotgun(
    'https://west.shotgunstudio.com',
    api_key='Nrsxkpgjjiprj~rz4hygmfkzb',
    script_name = 'ww_accessor'
)


DEV = 0


NOW_PATH = os.path.dirname( os.path.abspath( __file__ ) )

EXCEL_FILE = os.path.join(NOW_PATH, 'due_date_update', 'due_date_blank_241118_v02.xlsx')


def _get_excel_date( excel_file ):
    df = pd.read_excel(excel_file)

    data_list = []

    for _, row in df.iterrows():
        update_dict = {}

        update_dict['task_assignees'] = row['Assigned To'].strip() if row['Assigned To'] else ''
        update_dict['id'] = int(str(row['Id']).strip()) if row['Id'] else 0
        update_dict['entity'] = row['Link'].strip() if row['Link'] else ''
        update_dict['content'] = row['Task Name'].strip() if row['Task Name'] else ''
        update_dict['step'] = row['Pipeline Step'].strip() if row['Pipeline Step'] else ''
        update_dict['due_date'] = formatting_date(row['Due Date']) if row['Due Date'] else ''
        update_dict['project'] = row['Project'].strip() if row['Project'] else ''

        data_list.append(update_dict)

    return data_list


def formatting_date( date ):
    date_convert_str = datetime.strptime( date.strip(), "%Y.%m.%d" ).strftime( "%Y-%m-%d" )
    return date_convert_str


def find_and_duedate_update():
    if DEV:
        data_list = [{
            'id': 191142,
            'entity': 'S100_0010',
            'content': 'mm',
            'step': 'mm',
            'due_date': '2024-08-08',
            'project': 'RND'
        }]
    else:
        data_list = _get_excel_date( EXCEL_FILE )
    

    update_count = 0
    for data in data_list:
        filters = [
            ['content', 'is', data['content']],
            ['step.Step.code', 'is', data['step']],
            ['entity.Shot.code', 'is', data['entity']],
            ['id', 'is', data['id']],
            ['project.Project.name', 'is', data['project']]
        ]

        fields = ['content', 'step', 'entity', 'id', 'project', 'due_date', 'task_assignees']

        task = SG.find_one('Task', filters, fields)

        print(task['project']['name'], task['entity']['name'], f"{task['due_date']} --> {data['due_date']}")
        update_data = {'due_date': data['due_date']}

        if int(task['id']) == int(data['id']):
            update_task = SG.update('Task', task['id'], update_data)
            if update_task:
                update_count += 1
                print(f'update ::: {update_count}')
            else:
                print('failed update due_date')
                with open(os.path.join(NOW_PATH, 'due_date_error_shot', 'due_date_update_error.log'), 'a') as file:
                    file.write('=' * 20 + '\n')
                    file.write('failed update due_date' + '\n')
                    file.write('=' * 20 + '\n')
                    file.write(f'Project : {task["project"]["name"]}' + '\n')
                    file.write(f'Task : {task["content"]}' + '\n')
                    file.write(f'Id : {task["id"]}' + '\n')
                    file.write(f'Assigned Entity : {task["entity"]["name"]}' + '\n')
                    file.write('=' * 20 + '\n')
        else:
            print(f'Task Id : {task["id"]} // Excel Id : {data["id"]} are different.')
            with open(os.path.join(NOW_PATH, 'due_date_error_shot', 'due_date_update_error.log'), 'a') as file:
                file.write('=' * 20 + '\n')
                file.write('Task ID and Excel ID are different' + '\n')
                file.write('=' * 20 + '\n')
                file.write(f'Project : {task["project"]["name"]}' + '\n')
                file.write(f'Task : {task["content"]}' + '\n')
                file.write(f'Task Id : {task["id"]} // Excel Id : {data["id"]}' + '\n')
                file.write(f'Assigned Entity : {task["entity"]["name"]}' + '\n')
                file.write('=' * 20 + '\n')
                
            break




if __name__ == '__main__':
    find_and_duedate_update()
