# Query Monitoring Analyzer
Скрипт для анализа поисковых запросов из Яндекс.Вебмастера (Мониторинг запросов)  
  
Видео с комментариями по работе скрипта - https://youtu.be/8B459Kf03hE  
  
Потребуется установить Python с официального сайта - https://www.python.org/downloads/windows/  
Через командную строку в Windows (CMD) установить библиотеки - pip install pandas openpyxl
  
Для запуска необходимо скачать XLSX файл с данными по поисковым запросам и положить рядом со скриптом с названием wm.xlsx 

## История изменений

### Загружена версия 2.0 (15.01.2024)
* Удален подсчет мелькания и ядра (оказались бесполезны)
* Добавлена статистика по словам из запросов (для быстрой аналитики, например, выявления запросов со словом авито, либо ненужным ГЕО, ...)

### Загружена версия 1.0 (24.12.2023)
* Первая версия скрипт
