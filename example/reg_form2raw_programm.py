# pip install pypandoc

import pandas as pd
from pypandoc import convert_file, convert_text
from datetime import datetime


# загружаем данные

#sheet_id = "1p9JHisOGML5Wdmof8lC6dmylEnVreGfq" # идентификатор таблицы
#sheet_name = "answers" # название листа с данными

sheet_id = "1EyZJKY5OSo_VmAls8w7xUZWysOgj5teXQK636jZ8vl0"
sheet_name = "answers"

url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"

df = pd.read_csv(url)


# конвертировать будем с помощью pypandoc: markdown -> html -> docx

def md2html(md):
  return convert_text(md, 'html', format='md')

def html2docx(file):
  return convert_file(f'{file}.html', 'docx', format='html', outputfile=f'{file}.docx')


# функции для обработки записей

def short_name(df, row):
  """
  Возвращает "Фамилия И.О."
  """
  f_name_short = df.loc[row, "Фамилия Имя Отчество"]
  f_name_short = f_name_short.split(' ')

  return f_name_short[0] + ' ' + f_name_short[1][0] + '.' + f_name_short[2][0] + '.'

def records(df, row):
  """
  Обработка всех полей
  """
  rec = '## ' + df.loc[row, "Фамилия Имя Отчество"]
  rec += '\n**Дата регистрации**: ' + df.loc[row, "Отметка времени"]
  rec += '\n\n**Номер строки**: ' + str(row+1) + '\n'
  rec += '\n**Место работы**: ' + df.loc[row, "Место работы "]
  rec += '\n\n**Город**: ' + df.loc[row, "Город"]
  rec += '\n\n**Должность**: ' + df.loc[row, "Должность"]
  rec += '\n\n**Научная степень**: ' + str(df.loc[row, "Научная степень "])
  rec += '\n\n**Адрес электронной почты**: ' + df.loc[row, "Адрес электронной почты"]
  rec += '\n\n**Тема доклада**: ' + df.loc[row, "Тема доклада "]
  rec += '\n\n**Аннотация**:\n\n' + df.loc[row, "Аннотация доклада"]
  rec += '\n\n**Предполагаемая форма участия**: ' + df.loc[row, "Предполагаемая форма участия "]+ '\n\n'

  if isinstance(df.loc[row, "Комментарий "], str):
    rec += '\n\n**Комментарий**: ' + df.loc[row, "Комментарий "] + '\n\n'

  return rec


def reports1(df, row):
  rprt = '\n*' + df.loc[row, "Фамилия Имя Отчество"] + '*\n'
  rprt += '\n\n' + df.loc[row, "Место работы "]
  rprt += '\n\n**' + df.loc[row, "Тема доклада "] + '**\n\n'
  return rprt

def reports2(df, row):
  rprt = '\n\n' + short_name(df, row) + ' (*' + df.loc[row, "Место работы "] + '*) '
  rprt += '**' + df.loc[row, "Тема доклада "] + '**\n\n'
  return rprt

def person(df, row):
  return '\n\n **' + df.loc[row, "Фамилия Имя Отчество"] + '** – ' +  str(df.loc[row, "Научная степень "]) + ', ' + df.loc[row, "Место работы "] + '.\n\n'


# создадим два варианта списка докладов. 

reports_list1 = '# Список докладов (вариант 1)\n\n'
reports_list2 = '# Список докладов (вариант 2)\n\n'

# кроме списка докладов в итоговом файле будут список докладчиков и полные данные о регистрации

pers_list = '# Список докладчиков\n\n'
records_list = '# Полные данные регистрации\n\n'



# проходим все строки таблицы
for i in range(df.shape[0]):

  reports_list1 += reports1(df, i)
  reports_list2 += reports2(df, i)

  pers_list += person(df, i)

  records_list += records(df, i)

# объединяем списки, т.е. разделы итогового файла 
all_data = reports_list1 + reports_list2 + '\n\n---\n\n' + pers_list + '\n\n---\n\n' + records_list

# имя для файла с датой создания
now = datetime.now().strftime('%Y_%m_%d')
f_out = 'raw_prog_' + now

# сохраняем файл
with open(f'{f_out}.html', 'w', encoding='utf-8') as fout:
  fout.write(md2html(all_data))
  fout.close()

html2docx(f_out)