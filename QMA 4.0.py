import pandas as pd
import time
from datetime import datetime
import openpyxl
import os

def get_excel_files():
    return [f for f in os.listdir('.') if f.endswith('.xlsx')]

def select_file():
    excel_files = get_excel_files()
    
    if not excel_files:
        print('Excel файлы не найдены в текущей директории.')
        input("\nНажмите Enter для завершения...")
        return None
        
    print("\nДоступные Excel файлы:")
    for i, file in enumerate(excel_files, 1):
        print(f"[{i}] {file}")
        
    while True:
        try:
            choice = int(input("\nВыберите номер файла: "))
            if 1 <= choice <= len(excel_files):
                return excel_files[choice - 1]
            print("Неверный номер. Попробуйте еще раз.")
        except ValueError:
            print("Введите число.")

def determine_report_type(df):
    columns = df.columns.tolist()
    has_demand = any(col.endswith('_demand') for col in columns)
    
    try:
        url_index = columns.index('Url')
        query_index = columns.index('Query')
        is_pages_order = url_index < query_index
    except ValueError:
        is_pages_order = False
    
    if has_demand:
        return 1
    elif is_pages_order:
        return 2
    else:
        return None

def load_data(file_path):
    if not os.path.exists(file_path):
        print(f'Файл {file_path} не найден.')
        input("\nНажмите Enter для завершения...")
        return None
    return pd.read_excel(file_path)

def add_word_count_column(df):
    df['Кол-во слов'] = df['Query'].apply(lambda x: len(str(x).split()))
    return df

def create_output_file_name(input_file, domain, mode):
    current_date = datetime.now().strftime('%Y-%m-%d')
    if mode == 1:
        return f"{domain}-semantics-{current_date}.xlsx"
    elif mode == 2:
        return f"{domain}-pages-{current_date}.xlsx"
    else:
        raise ValueError("Неверный режим анализа. Допустимые значения: 1 или 2.")

def create_word_count_df(df):
    word_count = {}
    for query in df['Query']:
        words = str(query).split()
        for word in words:
            if word in word_count:
                word_count[word] += 1
            else:
                word_count[word] = 1
    word_count_df = pd.DataFrame(word_count.items(), columns=['Слово', 'Количество'])
    return word_count_df

def process_pages_data(df, site_url):
    shows_columns = [col for col in df.columns if col.endswith('_shows')]
    position_columns = [col for col in df.columns if col.endswith('_position')]
    clicks_columns = [col for col in df.columns if col.endswith('_clicks')]
    ctr_columns = [col for col in df.columns if col.endswith('_ctr')]

    df['Сум. показов за 14 дн'] = df[shows_columns].sum(axis=1)
    df['Ср. позиция'] = df[position_columns].apply(lambda x: x.replace(0, pd.NA).mean(), axis=1).round(1).fillna(0)
    df['Ср. дн. показов'] = df[shows_columns].mean(axis=1).round(0)
    df['Ср. число кликов'] = df[clicks_columns].mean(axis=1).round(0)
    df['Сум. кликов за 14 дн.'] = df[clicks_columns].sum(axis=1)
    df['Ср. CTR'] = df[ctr_columns].mean(axis=1).round(1)
    
    df['Полный URL'] = site_url + df['Url']

    result_df = df[[
        'Полный URL', 
        'Url',  
        'Query',  
        'Ср. позиция',
        'Ср. дн. показов',
        'Сум. показов за 14 дн',
        'Ср. число кликов',
        'Сум. кликов за 14 дн.',
        'Ср. CTR'
    ]]

    return result_df

def main():
    input_file = select_file()
    if input_file is None:
        return
        
    df = load_data(input_file)
    if df is None:
        return
        
    mode = determine_report_type(df)
    if mode is None:
        print("Не удалось определить тип отчета. Проверьте структуру файла.")
        input("\nНажмите Enter для завершения...")
        return
        
    site_url = input("\n\nПожалуйста, введите адрес сайта в формате https://site.ru: ")
    domain = site_url.split('//')[-1].split('/')[0]

    if mode == 1:
        print("\nОбнаружен отчет по поисковым запросам. Обработка...")
        df = add_word_count_column(df)
        
        demand_columns = [col for col in df.columns if col.endswith('_demand')]
        shows_columns = [col for col in df.columns if col.endswith('_shows')]
        position_columns = [col for col in df.columns if col.endswith('_position')]
        clicks_columns = [col for col in df.columns if col.endswith('_clicks')]

        df['Сум. частотность за 14 дн'] = df[demand_columns].sum(axis=1)
        df['Ср. позиция'] = df[position_columns].apply(lambda x: x.replace(0, pd.NA).mean(), axis=1).round(1).fillna(0)
        df['Медианная позиция'] = df[position_columns].apply(lambda x: x.replace(0, pd.NA).median(), axis=1).round(1).fillna(0)
        df['Охват'] = round((df[shows_columns].sum(axis=1) / df[demand_columns].sum(axis=1)) * 100, 1)
        df['Ср. дн. частотность'] = df[demand_columns].mean(axis=1).round(0)
        df['Ср. число кликов'] = df[clicks_columns].mean(axis=1).round(0)
        df['Сум. кликов за 14 дн.'] = df[clicks_columns].sum(axis=1)
        
        df['Полный URL'] = site_url + df['Url']

        min_total_frequency = 0
        result_df = df.loc[df['Сум. частотность за 14 дн'] >= min_total_frequency].sort_values(by='Сум. частотность за 14 дн', ascending=False)
        result_df = result_df[[
            'Query',
            'Url',
            'Полный URL',
            'Кол-во слов',
            'Ср. позиция',
            'Медианная позиция',
            'Ср. дн. частотность',
            'Сум. частотность за 14 дн',
            'Ср. число кликов',
            'Сум. кликов за 14 дн.',
            'Охват'
        ]]
        
        word_count_df = create_word_count_df(result_df)
        word_count_df = word_count_df.sort_values(by='Количество', ascending=False)

    elif mode == 2:
        print("\nОбнаружен отчет по страницам. Обработка...")
        result_df = process_pages_data(df, site_url)

    output_file_name = create_output_file_name(input_file, domain, mode)

    with pd.ExcelWriter(output_file_name, engine='openpyxl') as writer:
        sheet_name = 'Семантическое ядро' if mode == 1 else 'Страницы'
        result_df.to_excel(writer, index=False, sheet_name=sheet_name)
        
        for column in writer.sheets[sheet_name].columns:
            max_length = max(len(str(cell.value)) for cell in column)
            writer.sheets[sheet_name].column_dimensions[column[0].column_letter].width = min(max_length + 2, 50)

        if mode == 1:
            word_count_df.to_excel(writer, index=False, sheet_name='Статистика слов')
            writer.sheets['Статистика слов'].column_dimensions['A'].width = 25
            writer.sheets['Статистика слов'].column_dimensions['B'].width = 25

    num_queries = len(result_df)
    processing_time = time.process_time()
    print(f"\n***")
    print(f"Результат сохранен в файл {output_file_name}")
    if mode == 1:
        print(f"Обработано поисковых запросов: {num_queries}")
    elif mode == 2:
        print(f"Обработано адресов страниц: {num_queries}")
    print(f"***")
    input("\nНажмите Enter для завершения...")

if __name__ == "__main__":
    main()