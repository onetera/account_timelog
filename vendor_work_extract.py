# -*- coding: utf-8 -*-


import os
from datetime import datetime
from collections import defaultdict
from pprint import pprint


import shotgun_api3 as sa
import pandas as pd


DEV = 1


NOW_PATH = os.path.abspath( os.path.dirname( __file__ ) )


SG = sa.Shotgun(
    'https://west.shotgunstudio.com',
    api_key='Nrsxkpgjjiprj~rz4hygmfkzb',
    script_name = 'ww_accessor'
)


def work_to_vendor():
    projects = _get_projects_info()
    vendors = _get_vendors_info()

    all_data = []

    # 1~10
    months = [f"2024-{str(month).zfill(2)}" for month in range(1, 11)]

    for project in projects:
        print(f'Project name :: {project["name"]}')

        organized_data = []
        task_dict = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))

        tasks = _get_tasks_info( project, vendors )
        print(f'project tasks all count :: {len(tasks)}')
        
        task_count = 0
        if not tasks:
            for month in months: 
                organized_data.append({
                    'Vendor': '-',
                    'Month': month,
                    'Count': 0
                })
        else:
            for task in tasks:
                task_count +=1

                due_date = task['due_date']
                due_date_month = datetime.strptime(due_date, '%Y-%m-%d').strftime('%Y-%m')

                assignees = task['task_assignees']
                filtered_vendors = [assignee['name'] for assignee in assignees if assignee in vendors]

                for vendor in filtered_vendors:
                    if not task_dict[task['entity']['name']]:
                        task_dict[task['entity']['name']][vendor][due_date_month] = 1 
                
                print(f'save :: {task_count}')

            for shot_code, vendor_data in task_dict.items():
                for vendor, month_data in vendor_data.items():
                    for month, count in month_data.items():
                        organized_data.append({
                            'Vendor': vendor,
                            'Month': month,
                            'Count': count
                        })
            

        df = pd.DataFrame(organized_data)
        pivot_df = df.pivot_table(index='Vendor', columns='Month', values='Count', aggfunc='sum').reindex(columns=months, fill_value=0)

        pivot_df['Total'] = pivot_df[months].sum(axis=1)
        pivot_df = pivot_df.fillna(0).astype(int)
        pivot_df.insert(0, 'Project', project["name"])

        all_data.append(pivot_df)

    final_df = pd.concat(all_data)
    final_df.set_index('Project', append=True, inplace=True)
    final_df = final_df.reorder_levels(['Project', 'Vendor'])

    return final_df


def write_excel():
    df = work_to_vendor()

    if DEV:
        excel_file = os.path.join(NOW_PATH, 'vendor_report', 'DEV.xlsx')
    else:
        excel_file = os.path.join(NOW_PATH, 'vendor_report', 'vendor_work_2024_01_to_10.xlsx')
    
    writer = pd.ExcelWriter(excel_file, engine = "xlsxwriter")

    df = df.sort_values(by='Project')
    df.to_excel(writer, sheet_name="Vendor Work")

    worksheet = writer.sheets["Vendor Work"]
    worksheet.set_column('A:A', 20)
    worksheet.set_column('B:B', 30)

    writer.close()


def _get_projects_info():
    project_filter = [
        ['sg_status', 'in', ['Active', 'Finished']],
        ['name', 'not_in', ['_Timelog']]
    ]
    project_field = ['id', 'name', 'sg_status']

    if DEV:
        # project_filter.append(['name', 'is', 'queen'])
        # project_filter.append(['name', 'in', ['queen', 'date']])
        # project_filter.append(['name', 'in', ['queen', 'date', 'sweethome']])
        # project_filter.append(['name', 'in', ['son']])
        project_filter.append(['name', 'in', ['date']])
        # project_filter.append(['name', 'in', ['sweethome', 'son']])
        # project_filter.append(['name', 'is', 'sweethome'])

    return SG.find('Project', project_filter, project_field)


def _get_vendors_info():
    vendor_dept = SG.find_one('Department', [['id', 'is', 74]], ['users'])

    return vendor_dept['users']


def _get_tasks_info( project, vendors ):
    status_collect = [
        'wip',
        'wtg',
        'omt',
        'hld'
    ]

    task_filters = [
        ['project', 'is', project],
        ['sg_status_list', 'not_in', status_collect]
    ]

    if DEV:
        task_filters.append(['task_assignees.HumanUser.name', 'is', 'supervic supervic'])
        # task_filters.append(['due_date', 'greater_than', datetime(2024, 7, 31).strftime('%Y-%m-%d')])
        # task_filters.append(['due_date', 'less_than', datetime(2024, 9, 1).strftime('%Y-%m-%d')])
        task_filters.append(['due_date', 'less_than', datetime(2024, 11, 1).strftime('%Y-%m-%d')])
        task_filters.append(['due_date', 'greater_than', datetime(2023, 12, 31).strftime('%Y-%m-%d')])
    else:
        task_filters.append(['task_assignees', 'in', vendors])
        task_filters.append(['due_date', 'less_than', datetime(2024, 11, 1).strftime('%Y-%m-%d')])
        task_filters.append(['due_date', 'greater_than', datetime(2023, 12, 31).strftime('%Y-%m-%d')])

    task_fileds = ['id', 'sg_status_list', 'due_date', 
                   'content', 'task_assignees', 'entity']
    
    order = [{'field_name': 'due_date', 'direction': 'desc'}]

    tasks = SG.find('Task', task_filters, task_fileds, order=order)
    tasks = [task for task in tasks if task.get('entity') and task['entity'].get('type') == 'Shot']

    return tasks




if __name__ == "__main__":
    write_excel()
    # work_to_vendor()