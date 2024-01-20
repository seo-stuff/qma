import pandas as pd
import time
from datetime import datetime
import openpyxl
import os
import subprocess

def process_brands(query, brand_variations):
    return 'Да' if any(brand.lower() in query.lower() for brand in brand_variations) else 'Нет'

def load_data(file_path):
    if not os.path.exists(file_path):
        print(f'Файл {file_path} не найден. Скачайте отчет и положите рядом со скриптом.')
        return None
    return pd.read_excel(file_path)

def add_word_count_column(df):
    df['Кол-во слов'] = df['Поисковые запросы'].apply(lambda x: len(x.split()))
    return df

def create_output_file_name(input_file):
    domain_name = input_file.split('.')[0]
    current_date_time = datetime.now().strftime('%Y-%m-%d %H-%M-%S')
    return f"{domain_name} - {current_date_time} - YaWM.xlsx"

def create_word_count_df(df):
    word_count = {}
    for query in df['Поисковые запросы']:
        words = query.split()
        for word in words:
            if word in word_count:
                word_count[word] += 1
            else:
                word_count[word] = 1
    word_count_df = pd.DataFrame(word_count.items(), columns=['Слово', 'Количество'])
    return word_count_df

def main():
    # Ввод параметров
    brand_variations_input = input("Введите брендовые запросы, например: Yandex, Яндекс, ya (или Enter для пропуска): ").strip()
    brand_variations = [brand.strip() for brand in brand_variations_input.split(',')] if brand_variations_input else []
    
    min_total_frequency_input = input("\nВведите минимальную суммарную частотность запросов за 14 дн.(для исключения НЧ запросов): ").strip()
    min_total_frequency = int(min_total_frequency_input) if min_total_frequency_input else 0

    # Загрузка данных
    input_file = 'wm.xlsx'
    df = load_data(input_file)
    if df is None:
        return

    # Преобразование данных
    df.rename(columns={'Indicator': 'Поисковые запросы'}, inplace=True)
    df = add_word_count_column(df)
    
    # Обработка данных
    demand_columns = [col for col in df.columns if col.endswith('_demand')]
    shows_columns = [col for col in df.columns if col.endswith('_shows')]
    position_columns = [col for col in df.columns if col.endswith('_position')]
    clicks_columns = [col for col in df.columns if col.endswith('_clicks')]

    df['Сум. частотность за 14 дн'] = df[demand_columns].sum(axis=1)
    df['Ср. позиция'] = df[position_columns].apply(lambda x: x.replace(0, pd.NA).mean(), axis=1).round(1).fillna(0)
    df['Бренд'] = df['Поисковые запросы'].apply(lambda x: process_brands(x, brand_variations))
    df['Охват'] = round((df[shows_columns].sum(axis=1) / df[demand_columns].sum(axis=1)) * 100, 1)
    df['Ср. дн. частотность'] = df[demand_columns].mean(axis=1).round(0)
    df['Ср. число кликов'] = df[clicks_columns].mean(axis=1).round(0)
    df['Сум. кликов за 14 дн.'] = df[clicks_columns].sum(axis=1)

    # Фильтрация и сортировка данных
    result_df = df.loc[df['Сум. частотность за 14 дн'] >= min_total_frequency].sort_values(by='Сум. частотность за 14 дн', ascending=False)
    result_df = result_df[['Поисковые запросы', 'Кол-во слов', 'Ср. позиция', 'Ср. дн. частотность', 'Сум. частотность за 14 дн', 'Ср. число кликов', 'Сум. кликов за 14 дн.', 'Охват', 'Бренд']]

    # Создание DataFrame для статистики слов
    word_count_df = create_word_count_df(result_df)

    # Сортировка по убыванию количества слов
    word_count_df = word_count_df.sort_values(by='Количество', ascending=False)

    # Создание имени выходного файла
    output_file_name = create_output_file_name(input_file)

    # Сохранение файла Excel
    with pd.ExcelWriter(output_file_name, engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False, sheet_name='Семантическое ядро')
        writer.sheets['Семантическое ядро'].column_dimensions['A'].width = 50
        for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
            writer.sheets['Семантическое ядро'].column_dimensions[col].width = 25

        word_count_df.to_excel(writer, index=False, sheet_name='Статистика слов')
        writer.sheets['Статистика слов'].column_dimensions['A'].width = 25
        writer.sheets['Статистика слов'].column_dimensions['B'].width = 25

        explanations = [
            ["Поисковый запрос", "Поисковый запрос"],
            ["Кол-во слов", "Количество слов в запросе"],
            ["Ср. позиция", "Средняя позиция запроса за две недели (без учета 0)"],
            ["Ср. дн. частотность", "Среднедневная частотность запроса (округлено до целых)"],
            ["Сум. частотность за 14 дн", "Суммарная частотность запросов за две недели"],
            ["Охват", "Отношение показов сайта к общему спросу в процентах"],
            ["Бренд", "Запросы с вхождением бренда"],
            ["Ср. число кликов", "Среднее число кликов в день"],
            ["Сум. кликов за 14 дн.", "Суммарное число кликов за 14 дней"]
        ]
        df_explanations = pd.DataFrame(explanations, columns=['Параметр', 'Пояснение'])
        df_explanations.to_excel(writer, index=False, sheet_name='Пояснения')
        writer.sheets['Пояснения'].column_dimensions['A'].width = 20
        writer.sheets['Пояснения'].column_dimensions['B'].width = 100

    # Вывод статистики и ожидание
    num_queries = len(result_df)
    processing_time = time.process_time()
    print(f"Результат сохранен в файл {output_file_name}")
    print(f"\nОбработано поисковых запросов: {num_queries}")
    print(f"Время выполнения скрипта: {processing_time} секунд")
    input("Нажмите Enter для завершения...")

    # Открытие сгенерированного файла
    subprocess.Popen(['start', 'excel', output_file_name], shell=True)

if __name__ == "__main__":
    main()
