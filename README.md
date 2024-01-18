# Query Monitoring Analyzer

## Описание
Скрипт для анализа поисковых запросов из Яндекс.Вебмастера (Мониторинг запросов)  
* Ведётся подсчет суммы и средне-дневных занчений для кликов и спроса
* Считается средняя позиция за 14 дней
* Подсчитывается важный параметр "Охваты" (низкие охваты чаще всего означают нерелевантность запросов)
* Можно задать брендовые запросы и быстро отфильтровать фразы по вхождению заданных брендов
* Ведётся подсчет статистики слов для выявления аномалий в целях дальнейшей фильтрации

![Окно программы](demo.png)

![Результаты](demo2.png)

## Начало работы
Видео с комментариями по работе скрипта - https://youtu.be/8B459Kf03hE  
  
Потребуется установить Python с официального сайта - https://www.python.org/downloads/windows/  
Через командную строку в Windows (CMD) установить библиотеки:
> pip install pandas openpyxl
  
Для запуска необходимо скачать XLSX файл с данными по поисковым запросам и положить рядом со скриптом с названием wm.xlsx 

## История изменений

### Загружена версия 3.0 (18.01.2024)
* Автооткрытие файла после генерации
* Удален функционал коммерциализации по вхождению (слишком много частных случаев, в случае необходимости можно пользовать версией 2.0)

### Загружена версия 2.0 (15.01.2024)
* Удален подсчет мелькания и ядра (оказались бесполезны)
* Добавлена статистика по словам из запросов (для быстрой аналитики, например, выявления запросов с вхождением бренда, либо ненужным ГЕО, ...)
* Добавлен подсчет среднего числа кликов и суммы кликов за 14 дней
* Добавлено автоматическое определение коммерциализации по вхождению слов. Список можно редактивровать самостоятельно в настройках

### Загружена версия 1.0 (24.12.2023)
* Первая версия скрипт
