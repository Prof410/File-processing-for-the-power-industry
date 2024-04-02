# File-processing-for-the-power-industry
 Processing files for the electrical power industry
 
 Обработка файлов для электроэнергетики.
 
# Задача: есть скачанные фаилы в формате Excel (138 штук за каждый месяц).
# Task: there are downloaded files in Excel format (138 pieces for each month).

# Фаилы собираются на ежемесячной основе с https://www.atsenergo.ru/results/market/svnc 
# Files are collected on a monthly basis from https://www.atsenergo.ru/results/market/svnc

# Их нужно обработать и софрмировать 1 общий фаил по утвержденной заказчиком форме.
# They need to be processed and form 1 common file according to the form approved by the customer.


# Первый этап обработки фаилов.
# The first stage of file processing.

import os
import pandas as pd

# Запрашиваем у пользователя путь к папке
# Asking the user for the path to the folder
folder_path = input("Введите путь к папке, где хранятся файлы: ")

# Получаем список файлов в указанной папке
# Getting a list of files in the specified folder
file_list = os.listdir(folder_path)

# Проверяем, что список файлов не пуст
# Check that the file list is not empty
if len(file_list) > 0:
    for file_name in file_list:
        # Полный путь к текущему файлу
        # Full path to the current file
        file_path = os.path.join(folder_path, file_name)

        # Проверяем, что файл является Excel-файлом
        # Check that the file is an Excel file
        if file_name.endswith(".xls") or file_name.endswith(".xlsx"):
            # Открываем Excel-файл и создаем DataFrame
            # # Open the Excel file and create a Data Frame
            excel_df = pd.read_excel(file_path)

            # Удаляем столбцы 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5'
            # Deleting columns 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5'
            excel_df = excel_df.drop(columns=['Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5'], axis=1)

            # Заменяем значение ячейки 'параметры за расчетный период' на 'за расчетный период'
            # Replacing the value of the cell 'parameters for the billing period' with 'for the billing period'
            excel_df.replace('параметры за расчетный период', 'за расчетный период', inplace=True)

            # Удаляем строки, где значение в столбце 'Unnamed: 0' равно NaN
            # Deleting rows where the value in the column 'Unnamed: 0' is NaN
            excel_df = excel_df.dropna(subset=['Unnamed: 0'])

            # Выбираем нужные строки по заданным значениям в столбце 'Unnamed: 0'
            # Select the desired rows by the specified values in the column 'Unnamed: 0'
            excel_df = excel_df.loc[excel_df['Unnamed: 0'].isin(['за расчетный период',
                                                                 'для совокупности  ГТП',
                                                                 'для ГТП',
                                                                 'участника оптового рынка',
                                                                 'Дифференцированная по зонам суток расчетного периода средневзвешенная нерегулируемая цена на электрическую энергию (мощность) на оптовом рынке по трем зонам суток:',
                                                                 'Дифференцированная по зонам суток расчетного периода средневзвешенная нерегулируемая цена на электрическую энергию (мощность) на оптовом рынке по двум зонам суток:',
                                                                 'Средневзвешенная нерегулируемая цена на электрическую энергию на оптовом рынке, определяемая для соответствующей зоны суток:',
                                                                 'Ночная зона',
                                                                 'Дневная зона',
                                                                 'Полупиковая зона',
                                                                 'Пиковая зона'])]

            # Удаляем определенные строки по заданным индексам
            # # Deleting certain rows by the specified index
            # excel_df = excel_df.drop(index=[26, 27, 28, 30, 31])
            excel_df = excel_df.drop(excel_df.tail(5).index)
            
            # Создаем новое имя файла для сохранения (заменяем расширение на '.xlsx')
            # Create a new file name to save (replace the extension with '.xlsx')
            new_file_name = file_name.rsplit(".", 1)[0] + ".xlsx"

            # Полный путь к новому файлу
            # Full path to the new file
            new_file_path = os.path.join(folder_path, new_file_name)

            # Сохраняем DataFrame в Excel-файл
            # Saving the Data Frame to an Excel file
            excel_df.to_excel(new_file_path, index=False)

            # Удаляем старый Excel-файл
            # Deleting the old Excel file
            os.remove(file_path)

            print("Файл обработан и сохранен:", new_file_name)
        else:
            print("Файл", file_name, "не является Excel-файлом")
else:
    print("Папка пуста или указан некорректный путь")

# Визуальная проверка правильности вывода месяца.
# Visual verification of the correctness of the month output.
excel_df.head(100)


# Второй этап обработки фаилов + сохранение.
# The second stage of file processing + saving.

# Запросить у пользователя адрес папки
# Request a folder address from the user
folder_path = input("Введите путь к папке: ")

# Переходим в указанную папку
# Go to the specified folder
os.chdir(folder_path)

# Получаем список всех файлов в папке
# Getting a list of all the files in the folder
file_list = os.listdir()

# Создаем пустой DataFrame для объединения данных
# Creating an empty Data Frame to combine the data
combined_df = pd.DataFrame()

# Обрабатываем каждый файл Excel в папке
# Processing each Excel file in the folder
for file_name in file_list:
    if file_name.endswith(".xlsx"):
        # Открываем файл и создаем DataFrame
        # Open the file and create a Data Frame
        file_path = os.path.join(folder_path, file_name)
        df = pd.read_excel(file_path)

        # Копируем данные из столбца "Unnamed: 1" и вставляем их в итоговый DataFrame
        # Copy the data from the "Unnamed" column: 1" and insert them into the final DataFrame
        
   combined_df = pd.concat([combined_df, df["Unnamed: 1"]], axis=1)

# Сохраняем итоговый DataFrame в формате xlsx
# Saving the final Data Frame in xlsx format
output_file = os.path.join(folder_path, "combined_data.xlsx")
combined_df.to_excel(output_file, index=False)

print("Обработка завершена. Результат сохранен в файл 'combined_data.xlsx'.")
