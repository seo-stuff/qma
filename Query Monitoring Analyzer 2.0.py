import pandas as pd
import time
from datetime import datetime
import openpyxl

def process_brands(query, brand_variations):
    for brand in brand_variations:
        if brand.lower() in query.lower():
            return 'Да'
    return 'Нет'

def classify_commercialization(query):
    information_keywords = ["где", "зачем", "как", "какой", "какая", "какие", "какое", "каков", "когда", "который", "которая", "которое", "кто", "куда", "откуда", "почему", "сколько", "чей", "что"]
    commercial_keywords = ["цена", "цены", "цене", "купить", "стоимость", "москва", "москве", "москвы"]
    
    words = query.split()
    for word in words:
        if word.lower() in information_keywords:
            return "Информационный"
        elif word.lower() in commercial_keywords:
            return "Коммерческий"
    
    return "Неизвестно"

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

# Добавляем столбец "Бренд"
df['Бренд'] = df['Поисковые запросы'].apply(lambda x: process_brands(x, brand_variations))

# Добавляем столбец "Охват"
df['Охват'] = round((df[shows_columns].sum(axis=1) / df[demand_columns].sum(axis=1)) * 100, 1)

# Добавляем столбец "Ср. дн. частотность"
df['Ср. дн. частотность'] = df[demand_columns].mean(axis=1).round(0)

# Считаем среднее число кликов в день
clicks_columns = [col for col in df.columns if col.endswith('_clicks')]
df['Ср. число кликов'] = df[clicks_columns].mean(axis=1).round(0)

# Считаем суммарное число кликов за 14 дней
df['Сум. кликов за 14 дн.'] = df[clicks_columns].sum(axis=1)

# Добавляем столбец "Коммерциализация"
df['Коммерциализация'] = df['Поисковые запросы'].apply(lambda x: classify_commercialization(x))

# Фильтрация и сортировка данных
result_df = df.loc[df['Сум. частотность за 14 дн'] >= min_total_frequency].sort_values(by='Сум. частотность за 14 дн', ascending=False)
result_df = result_df[['Поисковые запросы', 'Ср. позиция', 'Ср. дн. частотность', 'Сум. частотность за 14 дн', 'Ср. число кликов', 'Сум. кликов за 14 дн.', 'Охват', 'Бренд', 'Коммерциализация']]

# Получаем название домена и текущую дату
domain_name = input_file.split('.')[0]
current_date_time = datetime.now().strftime('%Y-%m-%d %H-%M-%S')

# Имя файла для сохранения
output_file_name = f"{domain_name} - {current_date_time} - YaWM.xlsx"

# Создание статистики слов из поисковых запросов
word_count = {}
for query in result_df['Поисковые запросы']:
    words = query.split()
    for word in words:
        if word in word_count:
            word_count[word] += 1
        else:
            word_count[word] = 1

# Создание DataFrame для статистики слов
word_count_df = pd.DataFrame(word_count.items(), columns=['Слово', 'Количество'])
word_count_df = word_count_df.sort_values(by='Количество', ascending=False)

# Сохранение файла Excel с листами в нужном порядке и переименованием
with pd.ExcelWriter(output_file_name, engine='openpyxl') as writer:
    result_df.to_excel(writer, index=False, sheet_name='Семантическое ядро')

    # Настройка ширины колонок
    writer.sheets['Семантическое ядро'].column_dimensions['A'].width = 50
    for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
        writer.sheets['Семантическое ядро'].column_dimensions[col].width = 25

    # Создание нового листа для статистики слов
    word_count_df.to_excel(writer, index=False, sheet_name='Статистика слов')

    # Настройка ширины колонок в листе "Статистика слов"
    writer.sheets['Статистика слов'].column_dimensions['A'].width = 25
    writer.sheets['Статистика слов'].column_dimensions['B'].width = 25

    # Добавление листа с пояснениями
    explanations = [
        "'Поисковый запрос': Поисковый запрос",
        "'Ср. позиция': Средняя позиция запроса за две недели (без учета 0)",
        "'Ср. дн. частотность': Среднедневная частотность запроса (округлено до целых)",
        "'Сум. частотность за 14 дн': Суммарная частотность запросов за две недели",
        "'Охват': Отношение показов к спросу в процентах (округлено до десятых)",
        "'Бренд': Запросы с вхождением бренда",
        "'Ср. число кликов': Среднее число кликов в день",
        "'Сум. кликов за 14 дн.': Суммарное число кликов за 14 дней",
        "'Коммерциализация': Тип запроса (Информационный, Коммерческий, Неизвестно)"
    ]
    df_explanations = pd.DataFrame(explanations, columns=['Пояснение'])
    df_explanations.to_excel(writer, index=False, sheet_name='Пояснения')

# Вывод статистики и ожидание
num_queries = len(result_df)
processing_time = time.process_time()
print(f"Результат сохранен в файл {output_file_name}")
print(f"\nОбработано поисковых запросов: {num_queries}")
print(f"Время выполнения скрипта: {processing_time} секунд")
input("Нажмите Enter для завершения...")
