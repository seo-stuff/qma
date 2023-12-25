import pandas as pd
import time
from datetime import datetime
import openpyxl

def process_brands(query, brand_variations):
    for brand in brand_variations:
        if brand.lower() in query.lower():
            return 'Да'
    return 'Нет'

# Вариации брендовых запросов
brand_variations_input = input("Введите брендовые запросы, например: Yandex, Яндекс, ya (или Enter для пропуска): ").strip()
brand_variations = [brand.strip() for brand in brand_variations_input.split(',')] if brand_variations_input else []

# Минимальная суммарная частотность
min_total_frequency_input = input("\nВведите минимальную суммарную частотность запросов за 14 дн.(для исключения НЧ запросов): ").strip()
min_total_frequency = int(min_total_frequency_input) if min_total_frequency_input else 0

# Чтение данных из файла wm.xlsx
import os
input_file = 'wm.xlsx'
if not os.path.exists(input_file):
    print('Файл', input_file, 'не найден. Скачайте отчет по запросам из сервиса Мониторинг запросов на сайте Webmaster.Yandex.ru, переименуйте файл и положите рядом со скриптом.')
    input('Нажмите Enter после того, как файл будет добавлен...')
    if not os.path.exists(input_file):
        print('Файл по-прежнему не найден. Завершение скрипта.')
        exit()
df = pd.read_excel(input_file)

# Переименовываем колонку
df.rename(columns={'Indicator': 'Поисковые запросы'}, inplace=True)

# Суммируем данные в колонках _demand и _shows
demand_columns = [col for col in df.columns if col.endswith('_demand')]
df['Сум. частотность за 14 дн'] = df[demand_columns].sum(axis=1)
shows_columns = [col for col in df.columns if col.endswith('_shows')]

# Считаем среднюю позицию, исключая значения 0
position_columns = [col for col in df.columns if col.endswith('_position')]
df['Ср. позиция'] = df[position_columns].apply(lambda x: x.replace(0, pd.NA).mean(), axis=1).round(1).fillna(0)

# Добавляем столбцы "Ядро", "Бренд", "Мелькает" и "Охват"
clicks_columns = [col for col in df.columns if col.endswith('_clicks')]
df['Ядро'] = df[clicks_columns].apply(lambda x: 'Да' if x.min() > 0 else 'Нет', axis=1)
df['Бренд'] = df['Поисковые запросы'].apply(lambda x: process_brands(x, brand_variations))
df['Мелькает'] = df[position_columns].apply(lambda x: 'Да' if x.min() > 0 and x.max() > 0 and x.min() != x.max() else 'Нет', axis=1)
df['Охват'] = round((df[shows_columns].sum(axis=1) / df[demand_columns].sum(axis=1)) * 100, 1)

# Добавляем столбец "Ср. дн. частотность"
df['Ср. дн. частотность'] = df[demand_columns].mean(axis=1).round(0)

# Фильтрация и сортировка данных
result_df = df.loc[df['Сум. частотность за 14 дн'] >= min_total_frequency].sort_values(by='Сум. частотность за 14 дн', ascending=False)
result_df = result_df[['Поисковые запросы', 'Ср. позиция', 'Ср. дн. частотность', 'Сум. частотность за 14 дн', 'Охват', 'Бренд', 'Мелькает', 'Ядро']]

# Получаем название домена и текущую дату
domain_name = input_file.split('.')[0]
current_date_time = datetime.now().strftime('%Y-%m-%d %H-%M-%S')

# Имя файла для сохранения
output_file_name = f"{domain_name} - {current_date_time} - YaWM.xlsx"

# Сохранение файла Excel
with pd.ExcelWriter(output_file_name, engine='openpyxl') as writer:
    result_df.to_excel(writer, index=False, sheet_name='Data')

    # Настройка ширины колонок
    writer.sheets['Data'].column_dimensions['A'].width = 50
    for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H']:
        writer.sheets['Data'].column_dimensions[col].width = 25

    # Добавление листа с пояснениями
    explanations = [
        "'Поисковый запрос': Поисковый запрос",
        "'Ср. позиция': Средняя позиция запроса за две недели (без учета 0)",
        "'Ср. дн. частотность': Среднедневная частотность запроса (округлено до целых)",
        "'Сум. частотность за 14 дн': Суммарная частотность запросов за две недели",
        "'Охват': Отношение показов к спросу в процентах (округлено до десятых)",
        "'Бренд': Запросы с вхождением бренда",
        "'Мелькает': Запросы, по которым позиция мелькала (появлялась и исчезала)",
        "'Ядро': Запросы, по которым были клики каждый день"
    ]
    df_explanations = pd.DataFrame(explanations, columns=['Пояснение'])
    df_explanations.to_excel(writer, index=False, sheet_name='Пояснения')

print(f"Результат сохранен в файл {output_file_name}")

# Вывод статистики и ожидание
num_queries = len(result_df)
processing_time = time.process_time()
print(f"\nОбработано поисковых запросов: {num_queries}")
print(f"Время выполнения скрипта: {processing_time} секунд")
input("Нажмите Enter для завершения...")
