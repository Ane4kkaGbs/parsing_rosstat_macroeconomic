import pandas as pd
from docx import Document
import re

# Код берет ворд файл. из и достает таблицы, которые сделаны из левой и правой части по горизонтали типо разбиение
# и еще есть разбиение по вертикали (типо много строк)
# сейчас код успешно парсит в 1 полноценную таблицу только перот
# цель сейчас это наладить этот парсинг в рамках 1 файла ( тоесть решить проблемы с объединением
# некоторых таблиц, а также чтобы не брало не нужные данные в таблицы


# Функция для извлечения данных из таблицы
def extract_table_data(table):
    table_data = []
    for row in table.rows:
        row_data = [cell.text.strip() for cell in row.cells]
        table_data.append(row_data)
    return table_data

# Функция для очистки заголовков (удаление пустых заголовков)
def clean_headers(headers):
    cleaned_headers = []
    header_count = {}
    for header in headers:
        if not header:
            header = "Unnamed"
        if header in header_count:
            header_count[header] += 1
            header = f"{header}_{header_count[header]}"
        else:
            header_count[header] = 0
        cleaned_headers.append(header)
    return cleaned_headers

# Функция для проверки, является ли таблица сноской
def is_footnote_table(header):
    # Проверяем, содержит ли заголовок строки маркеры сносок, такие как цифры с точкой/скобкой
    if re.match(r"^\d+\)", header) or re.match(r"^\d+\.\)", header):
        return True
    return False

# Открытие документа Word
doc_path = "/content/R_01.docx"  # Замените на путь к вашему файлу
doc = Document(doc_path)

# Переменные для хранения объединённых таблиц
merged_tables = {}
num_tables = len(doc.tables)
i = 0

while i < num_tables - 1:
    try:
        # Извлекаем данные из текущей пары таблиц
        left_table_data = extract_table_data(doc.tables[i])
        right_table_data = extract_table_data(doc.tables[i + 1])

        # Проверяем и пропускаем таблицы-сноски
        if left_table_data and right_table_data:
            left_header_candidate = left_table_data[0][0] if left_table_data[0] else ""
            right_header_candidate = right_table_data[0][0] if right_table_data[0] else ""
            if is_footnote_table(left_header_candidate) or is_footnote_table(right_header_candidate):
                print(f"Пропуск таблиц {i + 1} и {i + 2} из-за сносок.")
                i += 2
                continue

        # Создание DataFrame из левой части таблицы
        if len(left_table_data) >= 2:
            left_header = clean_headers(left_table_data[0])
            df_left = pd.DataFrame(left_table_data[1:], columns=left_header)
            print(f"Обработка таблицы {i + 1}")

        # Создание DataFrame из правой части таблицы
        if len(right_table_data) >= 2:
            right_header = clean_headers(right_table_data[0])
            df_right = pd.DataFrame(right_table_data[1:], columns=right_header)
            print(f"Обработка таблицы {i + 2}")

        # Удаляем столбцы с полностью пустыми значениями
        df_left = df_left.dropna(axis=1, how='all')
        df_right = df_right.dropna(axis=1, how='all')

        # Проверяем наличие столбца с названиями регионов в обеих таблицах
        left_region_column = df_left.columns[0]
        right_region_column = df_right.columns[-1]

        # Объединяем таблицы на основе столбца с названиями регионов
        df_combined = pd.merge(df_left, df_right, left_on=left_region_column, right_on=right_region_column, how='inner')

        # Создаём ключ для хранения таблиц с одинаковыми заголовками
        header_key = tuple(df_combined.columns)

        # Если ключ уже есть в merged_tables, добавляем строки к существующему DataFrame
        if header_key in merged_tables:
            merged_tables[header_key] = pd.concat([merged_tables[header_key], df_combined], ignore_index=True)
        else:
            # Если ключа ещё нет, создаём новый DataFrame
            merged_tables[header_key] = df_combined

        # Переход к следующей паре таблиц
        i += 2
    except Exception as e:
        print(f"Ошибка при обработке таблиц {i + 1} и {i + 2}: {e}")
        # Переход к следующей паре таблиц даже при ошибке
        i += 2

# Сохранение всех объединённых таблиц
for idx, (header_key, df) in enumerate(merged_tables.items(), start=1):
    csv_file_path = f"w/merged_table_{idx}.csv"
    df.to_csv(csv_file_path, index=False)
    print(f"Объединённая таблица сохранена в файл: {csv_file_path}")
