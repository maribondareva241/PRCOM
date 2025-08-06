import openpyxl
import csv
import os

# Запрос у пользователя названия файла
file_name = input("Введите название файла Excel (с расширением, например, 'файл.xlsx'): ").strip()

# Путь к папке "Загрузки" пользователя "mariya"
downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')

# Полный путь к файлу
file_path = os.path.join(downloads_folder, file_name)

# Проверка существования файла
if not os.path.isfile(file_path):
    print(f"Файл не найден по пути: {file_path}")
    exit(1)

# Запрос значения для PRCOM-XXXX у пользователя
prcom_value = input("Введите значение для PRCOM-XXXX (например, 1234): ").strip()

# Запрос значения для tag у пользователя
tag_value = input("Введите значение tag (например, tag_bondareva): ").strip()

# Открываем файл Excel
wb = openpyxl.load_workbook(file_path)
ws = wb.active  # или wb['Имя листа'], если нужно

# 1. Находим столбец с имейлами по наличию '@' в ячейках
email_col_letter = None
for col_idx in range(1, ws.max_column + 1):
    for row_idx in range(2, ws.max_row + 1):  # пропускаем заголовок
        cell = ws.cell(row=row_idx, column=col_idx)
        if cell.value and isinstance(cell.value, str) and '@' in cell.value:
            email_col_letter = openpyxl.utils.get_column_letter(col_idx)
            break
    if email_col_letter:
        break

if not email_col_letter:
    raise ValueError("Столбец с имейлами не найден.")

# 2. Собираем все имейлы из этого столбца
emails = []
for row in ws.iter_rows(min_row=2, max_col=openpyxl.utils.column_index_from_string(email_col_letter), max_row=ws.max_row):
    cell = row[openpyxl.utils.column_index_from_string(email_col_letter) - 1]
    if cell.value and isinstance(cell.value, str) and '@' in cell.value:
        emails.append(cell.value)

# 3. Удаляем все строки после заголовка и вставляем только уникальные имейлы без пустых
ws.delete_rows(2, ws.max_row)  # удаляем все строки после заголовка

unique_emails = list(dict.fromkeys(emails))
for i, email in enumerate(unique_emails, start=2):
    ws.cell(row=i, column=1, value=email)

# 5. Копируем содержимое первого столбца во второй
for row_idx in range(2, len(emails)+2):
    email_value = ws.cell(row=row_idx, column=1).value
    ws.cell(row=row_idx, column=2).value = email_value

# 7. В C1 пишем 'tag_bondareva'
ws['C1'] = tag_value

# 8. Для остальных ячеек столбца C: если есть значение в B — пишем PRCOM-XXXX
for row_idx in range(2, len(unique_emails)+2):
    b_value = ws.cell(row=row_idx, column=2).value
    if b_value:
        ws.cell(row=row_idx, column=3).value = f'PRCOM-{prcom_value}'

# Сохраняем файл под новым именем или перезаписываем исходный
modified_excel_path = os.path.join(downloads_folder, 'modified_' + file_name)
wb.save(modified_excel_path)

print(f"Обработанный файл сохранен как: {modified_excel_path}")

# Теперь сохраняем данные из всех трех столбцов в CSV файл с именем из C2 (или по умолчанию)
tag_value_cell = ws['C2']
if tag_value_cell.value:
    filename_base = str(tag_value_cell.value).strip()
else:
    filename_base = 'output'

csv_filename = f"{filename_base}.csv"
csv_filepath = os.path.join(downloads_folder, csv_filename)

with open(csv_filepath, mode='w', newline='', encoding='utf-8') as csv_file:
    writer = csv.writer(csv_file)
    # Записываем заголовки трех колонок
    writer.writerow(['email', 'email', 'tag_bondareva'])
    # Записываем все строки данных
    for row_idx in range(2, len(unique_emails)+2):
        email_1 = ws.cell(row=row_idx, column=1).value
        email_2 = ws.cell(row=row_idx, column=2).value
        tag_bondareva = ws.cell(row=row_idx, column=3).value
        writer.writerow([email_1 or '', email_2 or '', tag_bondareva or ''])

print(f"Данные сохранены в файл: {csv_filepath}")