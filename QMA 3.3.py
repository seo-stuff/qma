import pandas as pd
import time
from datetime import datetime
import openpyxl
import os
import pyfiglet

def load_data(file_path):
    if not os.path.exists(file_path):
        print(f'Файл {file_path} не найден. Скачайте отчет и положите рядом со скриптом.')
        return None
    return pd.read_excel(file_path)

def add_word_count_column(df):
    df['Кол-во слов'] = df['Поисковые запросы'].apply(lambda x: len(x.split()))
    return df

def create_output_file_name(input_file, mode):
    domain_name = input_file.split('.')[0]
    current_date_time = datetime.now().strftime('%Y-%m-%d %H-%M-%S')
    if mode == 1:
        return f"semantics-{current_date_time}.xlsx"
    elif mode == 2:
        return f"pages-{current_date_time}.xlsx"
    else:
        raise ValueError("Неверный режим анализа. Допустимые значения: 1 или 2.")

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
    text = "YaWm Analyser"
    ascii_art = pyfiglet.figlet_format(text, font="slant")
    print(ascii_art)
    
    # Указание адреса сайта
    site_url = input("Пожалуйста, введите адрес сайта в формате https://site.ru: ")

    # Выбор режима анализа
    mode = int(input("Выберите режим анализа:\n[1] Анализ поисковых запросов\n[2] Анализ страниц\n"))

    # Загрузка данных
    input_file = 'wm.xlsx'
    df = load_data(input_file)
    if df is None:
        return

    # Преобразование данных
    if mode == 1:
        df.rename(columns={'Indicator': 'Поисковые запросы'}, inplace=True)
        df = add_word_count_column(df)
    elif mode == 2:
        df.rename(columns={'Indicator': 'Адрес страницы'}, inplace=True)
        df['Адрес страницы'] = site_url + df['Адрес страницы']

    # Обработка данных
    if mode == 1:
        demand_columns = [col for col in df.columns if col.endswith('_demand')]
        shows_columns = [col for col in df.columns if col.endswith('_shows')]
        position_columns = [col for col in df.columns if col.endswith('_position')]
        clicks_columns = [col for col in df.columns if col.endswith('_clicks')]

        df['Сум. частотность за 14 дн'] = df[demand_columns].sum(axis=1)
        df['Ср. позиция'] = df[position_columns].apply(lambda x: x.replace(0, pd.NA).mean(), axis=1).round(1).fillna(0)
        df['Охват'] = round((df[shows_columns].sum(axis=1) / df[demand_columns].sum(axis=1)) * 100, 1)
        df['Ср. дн. частотность'] = df[demand_columns].mean(axis=1).round(0)
        df['Ср. число кликов'] = df[clicks_columns].mean(axis=1).round(0)
        df['Сум. кликов за 14 дн.'] = df[clicks_columns].sum(axis=1)

        # Фильтрация и сортировка данных
        min_total_frequency = 0  # Минимальная суммарная частотность для фильтрации НЧ запросов
        result_df = df.loc[df['Сум. частотность за 14 дн'] >= min_total_frequency].sort_values(by='Сум. частотность за 14 дн', ascending=False)
        result_df = result_df[['Поисковые запросы', 'Кол-во слов', 'Ср. позиция', 'Ср. дн. частотность', 'Сум. частотность за 14 дн', 'Ср. число кликов', 'Сум. кликов за 14 дн.', 'Охват']]

    elif mode == 2:
        shows_columns = [col for col in df.columns if col.endswith('_shows')]
        position_columns = [col for col in df.columns if col.endswith('_position')]
        clicks_columns = [col for col in df.columns if col.endswith('_clicks')]

        df['Сум. показов за 14 дн'] = df[shows_columns].sum(axis=1)
        df['Ср. позиция'] = df[position_columns].apply(lambda x: x.replace(0, pd.NA).mean(), axis=1).round(1).fillna(0)
        df['Ср. дн. показов'] = df[shows_columns].mean(axis=1).round(0)
        df['Ср. число кликов'] = df[clicks_columns].mean(axis=1).round(0)
        df['Сум. кликов за 14 дн.'] = df[clicks_columns].sum(axis=1)

        result_df = df[['Адрес страницы', 'Ср. позиция', 'Ср. дн. показов', 'Сум. показов за 14 дн', 'Ср. число кликов', 'Сум. кликов за 14 дн.']]

    # Создание DataFrame для статистики слов (только для режима анализа поисковых запросов)
    if mode == 1:
        word_count_df = create_word_count_df(result_df)
        word_count_df = word_count_df.sort_values(by='Количество', ascending=False)

    # Создание имени выходного файла
    output_file_name = create_output_file_name(input_file, mode)

    # Сохранение файла Excel
    with pd.ExcelWriter(output_file_name, engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False, sheet_name='Семантическое ядро')
        writer.sheets['Семантическое ядро'].column_dimensions['A'].width = 50
        for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H']:
            writer.sheets['Семантическое ядро'].column_dimensions[col].width = 25

        if mode == 1:
            word_count_df.to_excel(writer, index=False, sheet_name='Статистика слов')
            writer.sheets['Статистика слов'].column_dimensions['A'].width = 25
            writer.sheets['Статистика слов'].column_dimensions['B'].width = 25

    # Вывод статистики и ожидание
    num_queries = len(result_df)
    processing_time = time.process_time()
    print(f"Результат сохранен в файл {output_file_name}")
    print(f"\nОбработано поисковых запросов: {num_queries}")
    print(f"Время выполнения скрипта: {processing_time} секунд")
    input("Нажмите Enter для завершения...")

if __name__ == "__main__":
    main()