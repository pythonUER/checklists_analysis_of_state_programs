#!/usr/bin/env python
# coding: utf-8

# <h1>Table of Contents<span class="tocSkip"></span></h1>
# <div class="toc"><ul class="toc-item"><li><span><a href="#Загрузка-данных-из-свода-ЧЕК-ЛИСТОВ" data-toc-modified-id="Загрузка-данных-из-свода-ЧЕК-ЛИСТОВ-1"><span class="toc-item-num">1&nbsp;&nbsp;</span>Загрузка данных из свода ЧЕК-ЛИСТОВ</a></span></li><li><span><a href="#Графики-рассчетные" data-toc-modified-id="Графики-рассчетные-2"><span class="toc-item-num">2&nbsp;&nbsp;</span>Графики рассчетные</a></span></li><li><span><a href="#Сводные-таблицы" data-toc-modified-id="Сводные-таблицы-3"><span class="toc-item-num">3&nbsp;&nbsp;</span>Сводные таблицы</a></span><ul class="toc-item"><li><span><a href="#По-отделам-УЭР" data-toc-modified-id="По-отделам-УЭР-3.1"><span class="toc-item-num">3.1&nbsp;&nbsp;</span>По отделам УЭР</a></span></li><li><span><a href="#По-сотрудникам-УЭР" data-toc-modified-id="По-сотрудникам-УЭР-3.2"><span class="toc-item-num">3.2&nbsp;&nbsp;</span>По сотрудникам УЭР</a></span></li><li><span><a href="#По-ИОГВ" data-toc-modified-id="По-ИОГВ-3.3"><span class="toc-item-num">3.3&nbsp;&nbsp;</span>По ИОГВ</a></span></li></ul></li><li><span><a href="#Анализ-за-определенный-период-времени-(за-апрель-2021,-за-последнюю-неделю)" data-toc-modified-id="Анализ-за-определенный-период-времени-(за-апрель-2021,-за-последнюю-неделю)-4"><span class="toc-item-num">4&nbsp;&nbsp;</span>Анализ за определенный период времени (за апрель 2021, за последнюю неделю)</a></span></li><li><span><a href="#Для-dashboard" data-toc-modified-id="Для-dashboard-5"><span class="toc-item-num">5&nbsp;&nbsp;</span>Для dashboard</a></span><ul class="toc-item"><li><span><a href="#Анализ-госпрограмм/-планов-реализации/-порядков-предоставления-субсидий-в-разрезе-ИОГВ." data-toc-modified-id="Анализ-госпрограмм/-планов-реализации/-порядков-предоставления-субсидий-в-разрезе-ИОГВ.-5.1"><span class="toc-item-num">5.1&nbsp;&nbsp;</span>Анализ госпрограмм/ планов реализации/ порядков предоставления субсидий в разрезе ИОГВ.</a></span></li><li><span><a href="#Анализ-госпрограмм/-планов-реализации/-порядков-предоставления-субсидий-по-месяцам." data-toc-modified-id="Анализ-госпрограмм/-планов-реализации/-порядков-предоставления-субсидий-по-месяцам.-5.2"><span class="toc-item-num">5.2&nbsp;&nbsp;</span>Анализ госпрограмм/ планов реализации/ порядков предоставления субсидий по месяцам.</a></span></li></ul></li></ul></div>

# In[1]:


import warnings
warnings.filterwarnings('ignore')

import pandas as pd
from datetime import date, datetime, timedelta

import plotly.graph_objects as go
from plotly.subplots import make_subplots
import branca
from math import log


# In[2]:


def get_key(d, value):
    """
    Функция для восстановления ключа по значению
    :param d: словарь dict
    :param value: значение, для которого ищем ключ
    :return: ключ из словаря
    """
    for keys, values in d.items():
        if value in values:
            return keys
    else:
        return 'ФИО исполнителя нет в списке сотрудников'
    
def cut_string(string, limit):
    words = string.split(' ')
    for word in words:
        if len(word) > limit:
            limit = len(word)
    lines = []
    for word in words:
        if len(lines) == 0:
            lines.append(word)
        else:
            if len(lines[-1]) + 1 + len(word) > limit:
                lines.append(word)
            else:
                lines[-1] += f' {word}'
    return '<br>'.join(lines)  

def make_beautiful_date(some_date):
    return f'{month_dict[some_date.month]} {some_date.year}'

def make_beautiful_period(start_date, end_date):
    return f'{start_date.day:02}.{start_date.month:02}-{end_date.day:02}.{end_date.month:02}.{end_date.year}'

def str_to_date(some_str_date):
    _ = some_str_date.split('.')
    return date(_[2], _[1], _[0])   


# In[3]:


month_dict = dict(zip([x for x in range(1,13)],
                      ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 
                       'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']))

task_list = ['Анализ госпрограмм', 'Анализ планов реализации', 'Анализ порядков предоставления субсидий']

employees_dict = {'Отдел аграрно-продовольственной политики, природопользования и финансового сектора':                  ['Шибина Л.В.', 'Басинских М.В.', 'Татьянина М.И.', 'Брик И.А.', 'Царик А.М.', 'Агеева Д.А.'],
                  'Отдел развития реального сектора экономики': ['Пушилин В.А.', 'Макарова Ю.О.', 'Зубова Е.И.', 
                                                                 'Преснякова Д.С.'],
                  'Отдел развития социальной сферы': ['Евсина М.Н.', 'Колошин Д.И.', 'Петрова А.Б.', 'Слюняева Т.Б.', 
                                                      'Шавина А.В.', 'Вязникова А.В.', 'Болдырева Е.Ю.']}

colors = ('#E7D776', '#93B45C', '#5FA0A6', '#EDB793', 
          '#EDDF88', '#8FCB68', '#7EB7BC', '#E7A376', 
         )          
my_colors = []
my_colors.extend(colors*10)
my_colormap=[[0.0, '#A083C8'],[0.25, '#5FA0A6'],[0.5, '#93B45C'],[0.75, '#E7D776'], [1.0, '#EDB793']]

my_colors_dict = {'Не согласовано': '#EDB793', 
                  'Согласовано с замечаниями': '#5FA0A6',
                  'Согласовано': '#8FCB68', 
                  'В работе': '#E7D776'}


# # Загрузка данных из свода ЧЕК-ЛИСТОВ

# In[4]:


# path = r'U:\~ 09_Машинное обучение_Прогноз показателей СЭР\ЧЕК-ЛИСТЫ и DATA-SHOP\ВЫГРУЗКА ЧЕК-ЛИСТОВ/'
path = r'U:\ЧЕК-ЛИСТЫ и DATA-SHOP\ВЫГРУЗКА ЧЕК-ЛИСТОВ/'
file_name = f'{path}RESULTS - Чек-листы.xlsx'
Sheets_names = pd.ExcelFile(file_name).sheet_names[:]


# In[5]:


df = pd.DataFrame()
for ind, sheet in enumerate(Sheets_names[-3:]):
    print(sheet)
    temp = pd.read_excel(file_name, sheet_name=sheet)
    temp['task'] = task_list[ind]
    temp['Отдел УЭР'] = [get_key(employees_dict, employee) for employee in temp['Исполнитель'].values]     
    df = pd.concat([df, temp])
df.index = range(df.shape[0])

for i in df.index:
    try:
#         print(df.loc[i, 'task'], df.loc[i, 'Исполнитель'], df.loc[i, 'Номер и дата заключения в СЭД "Дело"'])
        _ = list(map(int, df.loc[i, 'Дата заключения'].split('.')))
        df.loc[i, 'date'] = date(_[-1], _[-2], 1)        
        df.loc[i, 'date_conclusion'] = date(_[-1], _[-2], _[0])
    except AttributeError:
        print(f"\033[94mДата заключения не заполнена или в формате '%Y-%b-%d %H:%m:%s'\nЗадача: {df.loc[i, 'task']}, Исполнитель: {df.loc[i, 'Исполнитель']}, некорректная дата: {df.loc[i, 'Дата заключения']}\033[0m")
        print(df.loc[i, 'file'])
        df.loc[i, 'date'] = date(1900,1,1) 
    except IndexError:
        print(f"\033[96mДата заключения в неправильном формате\nЗадача: {df.loc[i, 'task']}, Исполнитель: {df.loc[i, 'Исполнитель']}, некорректная дата: {df.loc[i, 'Дата заключения']}\033[0m")
        print(df.loc[i, 'file'])
        df.loc[i, 'date'] = date(1900,2,1)    
    except ValueError:
        print(f"\033[92mВ дате заключения содержится текст\nЗадача: {df.loc[i, 'task']}, Исполнитель: {df.loc[i, 'Исполнитель']}, некорректная дата: {df.loc[i, 'Дата заключения']}\033[0m")
        print(df.loc[i, 'file'])
        df.loc[i, 'date'] = date(1900,3,1)
    
        
df = df[['task', 'date_conclusion', 'date', 'Месяц заключения', 'Отдел УЭР', 'Исполнитель', 
         'Ответственное отраслевое управление',
         'Тип согласования', 'Статус заключения', 'Номер и дата заключения в СЭД "Дело"',         
         'Наименование государственной программы', 'Номер государственной программы',
         'Постановления об утверждении ГП', 'Номер и дата Постановления об утверждении ГП',
         'Плановый год госпрограммы',
         'Номер и дата версии в АИС Проект-СМАРТ Про', 'Дата в АИС Проект-СМАРТ Про',
         'Номер и дата РКПД в СЭД Дело', 'Наименование приказа', 'file', ]]
df['Ответственное отраслевое управление'] = df['Ответственное отраслевое управление'].map("Управление {}".format)
df.sort_values(by=['task', 'date', 'Отдел УЭР', 'Исполнитель', 'Ответственное отраслевое управление'], inplace=True)
df.index = range(df.shape[0])


# In[6]:


for ind, col in enumerate(df.columns):
    print(ind, col) 
    
df_dict = {col: sorted([val for val in df[col].unique()]) for col in df.columns[[0, 2, 4, 5, 6, 8]]}


# In[7]:


departments = df_dict['Отдел УЭР']
iogvs = df_dict['Ответственное отраслевое управление']
dates = sorted(df_dict['date'])


# In[8]:


category_columns = ['task', 'Отдел УЭР', 'Исполнитель', 'Ответственное отраслевое управление',
                    'Тип согласования', 'Статус заключения', 'Наименование государственной программы',]

df[category_columns] = df[category_columns].astype("category")


# # Графики рассчетные

# In[ ]:


"""График ежемесячной нагрузки по Отделам УЭР"""
df_department = pd.DataFrame(columns=['task', 'date', 'department', 'value', 'text'])

for task in df_dict['task']:
    for date_ in  df_dict['date']:
        for department in df_dict['Отдел УЭР']:
            temp = df[df['task'].isin([task]) & df['date'].isin([date_]) & df['Отдел УЭР'].isin([department])]
            df_department.loc[len(df_department)] = {'task': task,
                                                     'date': date_,
                                                     'department': department,
                                                     'value': len(temp),
                                                     'text': f'<b>{department}</b><br>Период: {make_beautiful_date(date_)}<br>Количество заключений: {len(temp)}',
                                                    }
            
fig = make_subplots(rows=3, cols=1, 
                    subplot_titles=('Анализ госпрограмм',
                                    'Анализ планов реализации', 
                                    'Анализ порядков предоставления субсидий', ),
                    specs=[[{}], 
                           [{}],
                           [{}]],
                    )

###########################################################################################################################
for task_ind, task in enumerate(task_list):
###########################################################################################################################    
    for ind, department in enumerate(departments):        
        temp=df_department[df_department['department'].isin([department]) & df_department['task'].isin([task])]
        
        fig.add_trace(go.Bar(x=[make_beautiful_date(date_) for date_ in temp['date'].unique()],
                             y=temp['value'],
                             name=cut_string(department, 35),
                             text = temp['text'].unique(),
                             hoverinfo = 'text',
                             marker=dict(color=my_colors[ind], line_color='grey'),
                             legendgroup=department,
                             showlegend=True if task_ind==0 else False
                             ),
                      row=task_ind+1, col=1
                     )
        fig.update_xaxes(type='category', 
                         categoryarray=[make_beautiful_date(date_) for date_ in df_department['date'].unique()],
                         visible=True)
        fig.update_layout(barmode='group',
                          legend=dict(yanchor='top',
                                      font=dict(size=10)),
                          template='simple_white',
                         )
        
fig.update_layout(title = dict(text = f'Деятельность отраслевых отделов за период {make_beautiful_date(dates[0])} - {make_beautiful_date(dates[-1])}', 
                               font = dict(size = 16), 
                               xanchor = 'left', yanchor = 'top'),
                 )
###########################################################################################################################
fig.write_html('По отделам УЭР.html')
get_ipython().system('По отделам УЭР.html')


# In[ ]:


"""График количество согласований в разрезе ИОГВ за все время"""
df_iogv = pd.DataFrame(columns=['task', 'iogv', 'value', 'text'])

for task in df_dict['task']:
    for iogv in df_dict['Ответственное отраслевое управление']:
        temp = df[df['task'].isin([task]) & df['Ответственное отраслевое управление'].isin([iogv])]
        df_iogv.loc[len(df_iogv)] = {'task': task,
                                           'iogv': iogv,
                                           'value': len(temp),
                                           'text': f'<b>{iogv}</b><br>{task}<br>Количество согласований: {len(temp)}',
                                          }
        

fig=go.Figure()
###########################################################################################################################
for task_ind, task in enumerate(task_list):
###########################################################################################################################    
    temp=df_iogv[df_iogv['task'].isin([task])]
    fig.add_trace(go.Bar(x=temp['iogv'],
                         y=temp['value'],
                         text = temp['text'].unique(),
                         name=task,
                         hoverinfo = 'text',
                         marker=dict(color=my_colors[task_ind],
#                                      color=temp['value'].values,
#                                      colorscale=my_colormap, 
                                     line_color='grey'),
                         legendgroup=iogv,
                         showlegend=True
                        ),
#                   row=task_ind+1, col=1
                 )
###########################################################################################################################        
fig.update_xaxes(type='category', 
                 visible=False)
fig.update_layout(title = dict(text = f'Количество согласований в разрезе ИОГВ за период {make_beautiful_date(dates[0])} - {make_beautiful_date(dates[-1])}', 
                               font = dict(size = 16), 
                               xanchor = 'left', yanchor = 'top'),
                  legend=dict(yanchor='top',
                              font=dict(size=10), ),
                  barmode='stack',
                  template='simple_white',
                 )
###########################################################################################################################
fig.write_html('По ИОГВ.html')
get_ipython().system('По ИОГВ.html')


# In[ ]:


"""График по сотрудникам: количество согласований по месяцам"""
df_employee = pd.DataFrame(columns=['task', 'date', 'employee', 'value', 'text'])

for task in df_dict['task']:
    for date_ in  df_dict['date']:
        for employee in df_dict['Исполнитель']:
            temp = df[df['task'].isin([task]) & df['date'].isin([date_]) & df['Исполнитель'].isin([employee])]
            df_employee.loc[len(df_employee)] = {'task': task,
                                                 'date': date_,
                                                 'employee': employee,
                                                 'value': len(temp),
                                                 'text': f'<b>{employee}</b><br>Период: {make_beautiful_date(date_)}<br>Количество заключений: {len(temp)}',
                                                }
            
fig = make_subplots(rows=3, cols=1, 
                    subplot_titles=('Анализ госпрограмм',
                                    'Анализ планов реализации', 
                                    'Анализ порядков предоставления субсидий', ),
                    vertical_spacing=0.1                    
                   )
###########################################################################################################################
for task_ind, task in enumerate(task_list):
###########################################################################################################################    
    for ind, employee in enumerate(df_dict['Исполнитель']):        
        temp=df_employee[df_employee['employee'].isin([employee]) & df_employee['task'].isin([task])]
        
        fig.add_trace(go.Bar(x=[make_beautiful_date(date_) for date_ in temp['date'].unique()],
                             y=temp['value'],
                             name=cut_string(employee, 35),
                             text = temp['text'].unique(),
                             hoverinfo = 'text',
                             marker=dict(color=my_colors[ind], line_color='grey'),
                             legendgroup=employee,
                             showlegend=True if task_ind==0 else False
                             ),
                      row=task_ind+1, col=1
                     )
        
fig.update_xaxes(type='category', 
                 categoryarray=[make_beautiful_date(date_) for date_ in df_employee['date'].unique()],
                 visible=True)
        
fig.update_layout(title = dict(text = f'Деятельность отраслевых отделов за период {make_beautiful_date(dates[0])} - {make_beautiful_date(dates[-1])}', 
                               font = dict(size = 16), 
                               xanchor = 'left', yanchor = 'top'),
                  barmode='group',
#                   barmode='stack',
                  legend=dict(yanchor='top',
                              font=dict(size=10)),
                  template='simple_white',
                 )
###########################################################################################################################
fig.write_html('По сотрудникам УЭР.html')
get_ipython().system('По сотрудникам УЭР.html')


# In[ ]:


"""График по сотрудникам: количество согласований за все время"""
df_employee = pd.DataFrame(columns=['task', 'employee', 'value', 'text'])

for task in df_dict['task']:
    for employee in df_dict['Исполнитель']:
        temp = df[df['task'].isin([task]) & df['Исполнитель'].isin([employee])]
        df_employee.loc[len(df_employee)] = {'task': task,
                                             'employee': employee,
                                             'value': len(temp),
                                             'text': f'<b>{employee}</b><br>{task}<br>Количество заключений: {len(temp)}',
                                            }
        
fig = go.Figure()
###########################################################################################################################
for task_ind, task in enumerate(task_list):
###########################################################################################################################    
    for ind, employee in enumerate(df_dict['Исполнитель']):        
        temp=df_employee[df_employee['employee'].isin([employee]) & df_employee['task'].isin([task])]
        
        fig.add_trace(go.Bar(x=temp['employee'],
                             y=temp['value'],
#                              name=cut_string(employee, 35),
                             name=task,
                             text = temp['text'].unique(),
                             hoverinfo = 'text',
                             marker=dict(color=my_colors[task_ind], line_color='grey'),
                             legendgroup=employee,
                             showlegend=True if ind==0 else False
#                              showlegend=False
                             ),
                     )
fig.update_xaxes(visible=True)        
fig.update_layout(title = dict(text = f'Деятельность отраслевых отделов за период {make_beautiful_date(dates[0])} - {make_beautiful_date(dates[-1])}', 
                               font = dict(size = 16), 
                               xanchor = 'left', yanchor = 'top'),
                  barmode='stack',
                  legend=dict(yanchor='top',
                              font=dict(size=10),),
                  template='simple_white',
                 )
###########################################################################################################################
fig.write_html('По сотрудникам УЭР-0.html')
get_ipython().system('По сотрудникам УЭР-0.html')


# # Сводные таблицы

# ## По отделам УЭР

# In[ ]:


table=pd.pivot_table(df, values=['task'], index=['Отдел УЭР'], columns=['Статус заключения'], 
                     aggfunc='count', 
                     fill_value=0.0, margins=True, dropna=True, 
                     margins_name='Итого', observed=False)
table.sort_values(by=[('task', 'Итого')], ascending=False, inplace=True)

fig = go.Figure()
for ind, index_name in enumerate(table.index[1:]):
    for col_ind, col in enumerate(table.columns[:-1]):
        fig.add_trace(go.Bar(x=[cut_string(index_name, 20)],
                             y=[table.loc[index_name, col]],
                             name=col[-1],
                             text = f"<b>{cut_string(index_name, 20)}</b><br>{col[-1]}: {int(table.loc[index_name, col])}",
                             hoverinfo = 'text',
                             marker=dict(color=my_colors_dict[col[-1]], line_color='grey'),
                             legendgroup=col[-1],
                             showlegend=True if ind==0 else False
                            ),
                     )        
fig.update_layout(title = dict(text = f'Заключения отраслевых отделов УЭР<br>за период {make_beautiful_date(dates[0])} - {make_beautiful_date(dates[-1])}', 
                               font = dict(size = 16), 
                               xanchor = 'left', yanchor = 'top'),
                  barmode='stack',
                  legend=dict(yanchor='top',
                              font=dict(size=10),),
                  template='simple_white',
                 )
fig.show()
fig.write_html('Сводный_по отделам УЭР.html')
# !Сводный_по отделам УЭР.html
table


# ## По сотрудникам УЭР

# In[ ]:


table=pd.pivot_table(df, values=['task'], index=['Исполнитель'], columns=['Статус заключения'], 
                     aggfunc='count', 
                     fill_value=0.0, margins=True, dropna=True, 
                     margins_name='Итого', observed=False)
table.sort_values(by=[('task', 'Итого')], ascending=False, inplace=True)

fig = go.Figure()
for ind, index_name in enumerate(table.index[1:]):
    for col_ind, col in enumerate(table.columns[:-1]):
        fig.add_trace(go.Bar(x=[cut_string(index_name, 20)],
                             y=[table.loc[index_name, col]],
                             name=col[-1],
                             text = f"<b>{cut_string(index_name, 20)}</b><br>{col[-1]}: {int(table.loc[index_name, col])}",
                             hoverinfo = 'text',
                             marker=dict(color=my_colors_dict[col[-1]], line_color='grey'),
                             legendgroup=col[-1],
                             showlegend=True if ind==0 else False
                            ),
                     )        
fig.update_layout(title = dict(text = f'Заключения сотрудников УЭР<br>за период {make_beautiful_date(dates[0])} - {make_beautiful_date(dates[-1])}', 
                               font = dict(size = 16), 
                               xanchor = 'left', yanchor = 'top'),
                  barmode='stack',
                  legend=dict(yanchor='top',
                              font=dict(size=10),),
                  template='simple_white',
                 )
fig.show()
fig.write_html('Сводный_по сотрудникам.html')
# !Сводный_по сотрудникам.html
table


# In[ ]:


table=pd.pivot_table(df, values=['task'], index=['Исполнитель'], columns=['date'], 
                     aggfunc='count', 
                     fill_value=0.0, margins=True, dropna=True, 
                     margins_name='Итого', observed=False)
table.sort_values(by=[('task', 'Итого')], ascending=False, inplace=True)

fig = go.Figure()
for ind, index_name in enumerate(table.index[1:]):
    for col_ind, col in enumerate(table.columns[:-1]):
        fig.add_trace(go.Bar(x=[make_beautiful_date(col[-1])],
                             y=[table.loc[index_name, col]],
#                              x=[cut_string(index_name, 20)],
#                              y=[table.loc[index_name, col]],
#                              name=str(col[-1]),
                             name=index_name,
                             text = f"<b>{cut_string(index_name, 20)}</b><br>{make_beautiful_date(col[-1])}: {int(table.loc[index_name, col])}",
                             hoverinfo = 'text',
                             marker=dict(color=my_colors[ind], line_color='grey'),
                             legendgroup=index_name,
                             width=0.05,
                             showlegend=True if col_ind==0 else False
#                              showlegend=False
                            ),
                     )    
fig.update_xaxes(type='category', 
#                  categoryarray=[make_beautiful_date(date_) for date_ in dates],
                 visible=True)
fig.update_layout(title = dict(text = f'Заключения сотрудников УЭР<br>за период {make_beautiful_date(table.columns[0][-1])} - {make_beautiful_date(table.columns[-2][-1])}<br>(разбивка по месяцам)', 
                               font = dict(size = 16), 
                               xanchor = 'left', yanchor = 'top'),
                  barmode='group',
                  legend=dict(yanchor='top',
                              font=dict(size=10),),
                  template='simple_white',
                 )
fig.show()
fig.write_html('Сводный_по сотрудникам_по месяцам-1.html')
# !Сводный_по сотрудникам_по месяцам-1.html
table


# In[ ]:


"""ЭТОТ ГРАФИК НЕ НУЖЕН"""
table=pd.pivot_table(df, values=['task'], index=['Исполнитель'], columns=['date'], 
                     aggfunc='count', 
                     fill_value=0.0, margins=True, dropna=True, 
                     margins_name='Итого', observed=False)
table.sort_values(by=[('task', 'Итого')], ascending=False, inplace=True)

fig = go.Figure()
for ind, index_name in enumerate(table.index[1:]):
    for col_ind, col in enumerate(table.columns[:-1]):
        fig.add_trace(go.Bar(x=[cut_string(index_name, 20)],
                             y=[table.loc[index_name, col]],
                             name=str(col[-1]),
                             text = f"<b>{cut_string(index_name, 20)}</b><br>{make_beautiful_date(col[-1])}: {int(table.loc[index_name, col])}",
                             hoverinfo = 'text',
                             marker=dict(color=my_colors[ind], line_color='grey'),
                             legendgroup=str(col[-1]),
                             showlegend=True if ind==0 else False
#                              showlegend=False
                            ),
                     )        
fig.update_layout(title = dict(text = f'Заключения сотрудников УЭР<br>за период {make_beautiful_date(table.columns[0][-1])} - {make_beautiful_date(table.columns[-2][-1])}', 
                               font = dict(size = 16), 
                               xanchor = 'left', yanchor = 'top'),
                  barmode='stack',
                  legend=dict(yanchor='top',
                              font=dict(size=10),),
                  template='simple_white',
                 )
fig.show()
fig.write_html('Сводный_по сотрудникам_по месяцам.html')
# !Сводный_по сотрудникам_по месяцам.html
table


# ## По ИОГВ

# In[ ]:


table=pd.pivot_table(df, values=['task'], index=['Ответственное отраслевое управление'], columns=['Статус заключения'], 
                     aggfunc='count', 
                     fill_value=0.0, margins=True, dropna=True, 
                     margins_name='Итого', observed=False)
table.sort_values(by=[('task', 'Итого')], ascending=False, inplace=True)

fig = go.Figure()
for ind, index_name in enumerate(table.index[1:]):
    for col_ind, col in enumerate(table.columns[:-1]):
        fig.add_trace(go.Bar(y=[cut_string(index_name, 30)],
                             x=[table.loc[index_name, col]],
                             name=cut_string(col[-1], 30),
                             text = f"<b>{cut_string(index_name, 30)}</b><br>{col[-1]}: {int(table.loc[index_name, col])}",
                             hoverinfo = 'text',
                             marker=dict(color=my_colors_dict[col[-1]], line_color='grey'),
                             legendgroup=col[-1],
                             showlegend=True if ind==0 else False,
                             orientation='h'
                            ),
                     )  
fig.update_xaxes(visible=True) 
fig.update_layout(title = dict(text = f'Количество заключений в разрезе ответственных ИОГВ<br>за период {make_beautiful_date(dates[0])} - {make_beautiful_date(dates[-1])}', 
                               font = dict(size = 16), 
                               xanchor = 'left', yanchor = 'top'),
                  barmode='stack',
                  legend=dict(yanchor='top',
                              font=dict(size=10),),
                  template='simple_white',
                 )
fig.show()
fig.write_html('Сводный_по ИОГВ.html')
# !Сводный_по ИОГВ.html
table


# In[ ]:


final = 'Не согласовано'
table=pd.pivot_table(df[df['Статус заключения'].isin([final])], 
                     values=['task'], index=['Ответственное отраслевое управление'], columns=['Исполнитель'], 
                     aggfunc='count', 
                     fill_value=0.0, margins=True, dropna=True, 
                     margins_name='Итого', observed=False)
table.sort_values(by=[('task', 'Итого')], ascending=False, inplace=True)

fig = go.Figure()
for ind, index_name in enumerate(table.index[1:]):
    for col_ind, col in enumerate(table.columns[:-1]):
        fig.add_trace(go.Bar(y=[cut_string(index_name, 30)],
                             x=[table.loc[index_name, col]],
                             name=cut_string(col[-1], 30),
                             text = f"<b>{cut_string(index_name, 30)}</b><br>{final}<br>{col[-1]}: {int(table.loc[index_name, col])}",
                             hoverinfo = 'text',
                             marker=dict(color=my_colors_dict[final], line_color='grey'),
                             legendgroup=col[-1],
                             showlegend=True if ind==0 else False,
#                              showlegend=True,
                             orientation='h'
                            ),
                     )  
        
final = 'Согласовано с замечаниями'
table=pd.pivot_table(df[df['Статус заключения'].isin([final])], 
                     values=['task'], index=['Ответственное отраслевое управление'], columns=['Исполнитель'], 
                     aggfunc='count', 
                     fill_value=0.0, margins=True, dropna=True, 
                     margins_name='Итого', observed=False)
table.sort_values(by=[('task', 'Итого')], ascending=False, inplace=True)

for ind, index_name in enumerate(table.index[1:]):
    for col_ind, col in enumerate(table.columns[:-1]):
        fig.add_trace(go.Bar(y=[cut_string(index_name, 30)],
                             x=[table.loc[index_name, col]],
                             name=cut_string(col[-1], 30),
                             text = f"<b>{cut_string(index_name, 30)}</b><br>{final}<br>{col[-1]}: {int(table.loc[index_name, col])}",
                             hoverinfo = 'text',
                             marker=dict(color=my_colors_dict[final], line_color='grey'),
                             legendgroup=col[-1],
                             showlegend=False,
                             orientation='h'
                            ),
                     )          
        
final = 'Согласовано'
table=pd.pivot_table(df[df['Статус заключения'].isin([final])], 
                     values=['task'], index=['Ответственное отраслевое управление'], columns=['Исполнитель'], 
                     aggfunc='count', 
                     fill_value=0.0, margins=True, dropna=True, 
                     margins_name='Итого', observed=False)
table.sort_values(by=[('task', 'Итого')], ascending=False, inplace=True)

for ind, index_name in enumerate(table.index[1:]):
    for col_ind, col in enumerate(table.columns[:-1]):
        fig.add_trace(go.Bar(y=[cut_string(index_name, 30)],
                             x=[table.loc[index_name, col]],
                             name=cut_string(col[-1], 30),
                             text = f"<b>{cut_string(index_name, 30)}</b><br>{final}<br>{col[-1]}: {int(table.loc[index_name, col])}",
                             hoverinfo = 'text',
                             marker=dict(color=my_colors_dict[final], line_color='grey'),
                             legendgroup=col[-1],
                             showlegend=False,
                             orientation='h'
                            ),
                     )  
        
fig.update_xaxes(visible=True) 
fig.update_layout(title = dict(text = f'Количество заключений в разрезе ответственных ИОГВ и сотрудников УЭР<br>за период {make_beautiful_date(dates[0])} - {make_beautiful_date(dates[-1])}', 
                               font = dict(size = 16), 
                               xanchor = 'left', yanchor = 'top'),
                  barmode='stack',
                  legend=dict(yanchor='top',
                              font=dict(size=10),),
                  template='simple_white',
                 )
fig.show()
fig.write_html(f'Сводный график по ИОГВ_по сотрудникам.html')
# !Сводный график по ИОГВ_по сотрудникам.html


# # Анализ за определенный период времени (за апрель 2021, за последнюю неделю)

# In[ ]:


"""ЗА ПОСЛЕДНИЙ МЕСЯЦ"""
table=pd.pivot_table(df[df['date'].isin([dates[-1]])], 
                     values=['file'], 
                     index=['task', 'Наименование государственной программы', 'Исполнитель'], 
                     columns=['Статус заключения'], 
                     aggfunc='count', 
                     fill_value=0.0, margins=True, dropna=True, 
                     margins_name='Итого', observed=False)
table


# In[ ]:


"""ЗА ПОСЛЕДНЮЮ НЕДЕЛЮ"""
try:
    table=pd.pivot_table(df[df['date_conclusion'] > (date.today() + timedelta(days=-7))], 
                         values=['file'], 
                         index=['task', 'Наименование государственной программы', 'Исполнитель'], 
                         columns=['Статус заключения'], 
                         aggfunc='count', 
                         fill_value=0.0, margins=True, dropna=True, 
                         margins_name='Итого', observed=False)
    table
except Exception as e:
    print(e)
table


# In[ ]:


"""В РАБОТЕ"""
try:
    table=pd.pivot_table(df[df['Статус заключения'].isin(['В работе'])], 
                         values=['file'], 
                         index=['task', 'Наименование государственной программы', 'Исполнитель', ], 
                         columns=['Статус заключения'], 
                         aggfunc='count', 
                         fill_value=0.0, margins=True, dropna=True, 
                         margins_name='Итого', observed=False)
    table
except Exception as e:
    print(e)


# In[ ]:


df.columns


# # Для dashboard

# In[9]:


df_dict


# ## Анализ госпрограмм/ планов реализации/ порядков предоставления субсидий в разрезе ИОГВ.

# In[10]:


for i in range(len(df_dict['task'])):
    task = df_dict['task'][i]
    # здесь устанавливаем срез - за какие месяца хотим строить графики, для примера взят срез "за последние три месяца"
    dates_ = df_dict['date'][-3:]
    table=pd.pivot_table(df[df['task'].isin([task]) & df['date'].isin(dates_)], 
                         values=['task'], index=['Ответственное отраслевое управление'], columns=['Статус заключения'], 
                         aggfunc='count', 
                         fill_value=0.0, margins=True, dropna=True, 
                         margins_name='Итого', observed=False)
    table.sort_values(by=[('task', 'Итого')], ascending=False, inplace=True)
    

    fig = go.Figure()
    for ind, index_name in enumerate(table.index[1:]):
        for col_ind, col in enumerate(table.columns[:-1]):
            fig.add_trace(go.Bar(y=[cut_string(index_name, 30)],
                                 x=[table.loc[index_name, col]],
                                 name=cut_string(col[-1], 30),
                                 text = f"<b>{cut_string(index_name, 30)}</b><br>{col[-1]}: {int(table.loc[index_name, col])}",
                                 hoverinfo = 'text',
                                 marker=dict(color=my_colors_dict[col[-1]], line_color='grey'),
                                 legendgroup=col[-1],
                                 showlegend=True if ind==0 else False,
                                 orientation='h'
                                ),
                         )  
    fig.update_xaxes(visible=True) 
    fig.update_layout(title = dict(text = f'<b>{task}</b>    <br>Количество заключений в разрезе ответственных ИОГВ    <br>за период {make_beautiful_date(dates_[0])} - {make_beautiful_date(dates_[-1])}', 
                                   font = dict(size = 16), 
                                   xanchor = 'left', yanchor = 'top'),
                      barmode='stack',
                      legend=dict(yanchor='top', xanchor='right',
                                  y=1, x=1,
                                  font=dict(size=10),
                                  orientation = 'h'),
                      template='simple_white',
                     )
    fig.show()
    fig.write_html(f'Сводный_по ИОГВ_{task}.html')    


# ## Анализ госпрограмм/ планов реализации/ порядков предоставления субсидий по месяцам.

# In[11]:


for i in range(len(df_dict['task'])):
    task = df_dict['task'][i]
    # здесь устанавливаем срез - за какие месяца хотим строить графики, для примера взят срез "за последние три месяца"
    dates_ = df_dict['date'][:]
    table=pd.pivot_table(df[df['task'].isin([task]) & df['date'].isin(dates_)], 
                         values=['task'], index=['date'], columns=['Статус заключения'], 
                         aggfunc='count', 
                         fill_value=0.0, margins=True, dropna=True, 
                         margins_name='Итого', observed=False)
    table.sort_values(by=[('task', 'Итого')], ascending=False, inplace=True)
    

    fig = go.Figure()
    for ind, index_name in enumerate(table.index[1:]):
        for col_ind, col in enumerate(table.columns[:-1]):
            fig.add_trace(go.Bar(x=[make_beautiful_date(index_name)],
                                 y=[table.loc[index_name, col]],
                                 name=cut_string(col[-1], 30),
                                 text = f"<b>{make_beautiful_date(index_name)}</b><br>{col[-1]}: {int(table.loc[index_name, col])}",
                                 hoverinfo = 'text',
                                 marker=dict(color=my_colors_dict[col[-1]], line_color='grey'),
                                 legendgroup=col[-1],
                                 showlegend=True if ind==0 else False,
                                 orientation='v'
                                ),
                         ) 
            
    """ЗА ПОСЛЕДНЮЮ НЕДЕЛЮ"""
    try:
        table=pd.pivot_table(df[df['task'].isin([task]) & (df['date_conclusion'] > (date.today() + timedelta(days=-7)))], 
                                 values=['task'], 
                                 index=['date'], 
                                 columns=['Статус заключения'], 
                                 aggfunc='count', 
                                 fill_value=0.0, margins=True, dropna=True, 
                                 margins_name='Итого', observed=False)

        for ind, index_name in enumerate(table.index[:-1]):
                for col_ind, col in enumerate(table.columns[:-1]):
                    fig.add_trace(go.Bar(x=['За последние 7 дней'],
                                         y=[table.loc[index_name, col]],
                                         name=cut_string(col[-1], 30),
                                         text = f"<b>За период {make_beautiful_period(start_date = (date.today() + timedelta(days=-7)), end_date = date.today())}</b><br>{col[-1]}: {int(table.loc[index_name, col])}",
                                         hoverinfo = 'text',
                                         marker=dict(color=my_colors_dict[col[-1]], line_color='grey'),
                                         legendgroup=col[-1],
                                         showlegend=False,
                                         orientation='v'
                                        ),
                                 ) 
    except Exception as e:
        print(e, f'\n{task}: нет заключений за последнюю неделю')
        
    
    x_cat = [make_beautiful_date(date_) for date_ in dates]
    x_cat.extend(['За последние 7 дней'])
    fig.update_xaxes(visible=True, 
                     type='category', 
                 categoryarray=x_cat,) 
    fig.update_layout(title = dict(text = f'<b>{task}</b>    <br>Количество заключений в разрезе ответственных ИОГВ    <br>за период {make_beautiful_date(dates_[0])} - {make_beautiful_date(dates_[-1])}', 
                                   font = dict(size = 16), 
                                   xanchor = 'left', yanchor = 'top'),
                      barmode='stack',
                      legend=dict(yanchor='top', xanchor='right',
                                  y=1, x=1,
                                  font=dict(size=10),
                                  orientation = 'h'),
                      template='simple_white',
                     )
    fig.show()
    fig.write_html(f'Сводный_по месяцам_{task}.html')


# In[ ]:




