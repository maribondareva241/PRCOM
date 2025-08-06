import openpyxl
import csv
import os

def get_column_with_emails_by(sheet):
    # 1. Находим столбец с имейлами по наличию '@' в ячейках
    for col_idx in range(1, sheet.max_column + 1):
        for row_idx in range(2, sheet.max_row + 1):  # пропускаем заголовок
            cell = sheet.cell(row=row_idx, column=col_idx)
            if cell.value and isinstance(cell.value, str) and '@' in cell.value:
                return openpyxl.utils.get_column_letter(col_idx)
    raise ValueError("Столбец с имейлами не найден.")


def main():
    # Запрос у пользователя названия файла
    file_extension = ".xlsx"
    resource_file_name = input("Введите название файла Excel: ").strip() + file_extension

    # Путь к папке "Загрузки" пользователя "mariya"
    current_folder = os.getcwd()
    resources_folder = os.path.join(current_folder, "resources")
    outcome_folder = os.path.join(current_folder, "outcome")

    # Полный путь к файлу
    resource_file_path = os.path.join(resources_folder, resource_file_name)

    # Проверка существования файла
    if not os.path.isfile(resource_file_path):
        print(f"Файл не найден по пути: {resource_file_path}")
        exit(1)

    # Запрос значения для PRCOM-XXXX у пользователя
    prcom_value = input("Введите значение для PRCOM-XXXX (например, 1234): ").strip()

    # Запрос значения для tag у пользователя
    tag_column_name = input("Введите значение tag (например, tag_bondareva): ").strip()

    # Открываем файл Excel
    workbook = openpyxl.load_workbook(resource_file_path)
    work_sheet = workbook.active  # или wb['Имя листа'], если нужно

    email_column = get_column_with_emails_by(work_sheet)

    # 2. Собираем все имейлы из этого столбца
    emails = []
    for row in work_sheet.iter_rows(min_row=2, max_col=openpyxl.utils.column_index_from_string(email_column), max_row=work_sheet.max_row):
        cell = row[openpyxl.utils.column_index_from_string(email_column) - 1]
        if cell.value and isinstance(cell.value, str) and '@' in cell.value:
            emails.append(cell.value)

    # 3. Удаляем все строки после заголовка и вставляем только уникальные имейлы без пустых
    work_sheet.delete_rows(2, work_sheet.max_row)  # удаляем все строки после заголовка

    unique_emails = list(dict.fromkeys(emails))
    for i, email in enumerate(unique_emails, start=2):
        work_sheet.cell(row=i, column=1, value=email)

    # 5. Копируем содержимое первого столбца во второй
    for row_idx in range(2, len(emails)+2):
        email_value = work_sheet.cell(row=row_idx, column=1).value
        work_sheet.cell(row=row_idx, column=2).value = email_value

    # 7. В C1 пишем 'tag_bondareva'
    work_sheet['C1'] = tag_column_name

    # 8. Для остальных ячеек столбца C: если есть значение в B — пишем PRCOM-XXXX
    for row_idx in range(2, len(unique_emails)+2):
        b_value = work_sheet.cell(row=row_idx, column=2).value
        if b_value:
            work_sheet.cell(row=row_idx, column=3).value = f'PRCOM-{prcom_value}'

    # Сохраняем файл под новым именем или перезаписываем исходный
    modified_excel_path = os.path.join(outcome_folder, 'modified_' + resource_file_name)
    workbook.save(modified_excel_path)

    print(f"Обработанный файл сохранен как: {modified_excel_path}")

    # Теперь сохраняем данные из всех трех столбцов в CSV файл с именем из C2 (или по умолчанию)
    tag_value_cell = work_sheet['C2']
    if tag_value_cell.value:
        filename_base = str(tag_value_cell.value).strip()
    else:
        filename_base = 'output'

    csv_filename = f"{filename_base}.csv"
    csv_filepath = os.path.join(outcome_folder, csv_filename)

    with open(csv_filepath, mode='w', newline='', encoding='utf-8') as csv_file:
        writer = csv.writer(csv_file)
        # Записываем заголовки трех колонок
        writer.writerow(['email', 'email', tag_column_name])
        # Записываем все строки данных
        for row_idx in range(2, len(unique_emails)+2):
            email_1 = work_sheet.cell(row=row_idx, column=1).value
            email_2 = work_sheet.cell(row=row_idx, column=2).value
            tag_value = work_sheet.cell(row=row_idx, column=3).value
            writer.writerow([email_1 or '', email_2 or '', tag_value or ''])

    print(f"Данные сохранены в файл: {csv_filepath}")


if __name__ == "__main__":
    main()