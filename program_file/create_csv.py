import csv
import sqlite3
import re

conn = sqlite3.connect(r'db.sqlite3')
cur = conn.cursor()

change_symbol = re.compile('[\)|\(|,| |\'|]')

# Считывание данных из CSV файла
with open('12_2022_import_UA.csv', encoding='utf-8') as r_file:
    file_reader = csv.reader(r_file, delimiter=";")
    total_row_count = 0
    
    name_of_table = input('Input name of table: ')
    string_value_head = ''
    string_value =''

    for row in file_reader:
        if total_row_count == 0:
            head_table = [i.replace('№', 'No') for i in row]
            count = 0
            # Вывод строки, содержащей заголовки для столбцов
            for value in head_table:
                count += 1
                value = change_symbol.sub('_', value)
                if count < len(head_table):    
                    string_value_head += value + ' VARCHAR, '
                    string_value += value + ', '
                elif count == len(head_table):
                    string_value_head += value + ' VARCHAR'
                    string_value += value 
                print(string_value_head)
            cur.execute(f'CREATE TABLE IF NOT EXISTS "{name_of_table}" ({string_value_head});')
            print(f'Table: "{name_of_table}" created successfully!')
            conn.commit()
            total_row_count += 1
 
        else: 
            sign_question = '?'
            for i in range(count - 1):
                sign_question += ', ?'              
        
            sql = f'INSERT INTO "{name_of_table}"({string_value}) values({sign_question});'
            data = []

            for row in file_reader:
                row = tuple(row)
                data.append(row)
            
            conn.executemany(sql, data)
            conn.commit()
            print('File save in db')

        total_row_count += 1
    
    print(f'Всего в файле {total_row_count} строк.')
    
