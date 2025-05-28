# Excel Шаблонизатор

Мощный инструмент для генерации Excel отчетов на основе шаблонов с использованием Jinja2.

## 🚀 Возможности

- ✅ Подстановка переменных через Jinja2 синтаксис `{{variable}}`
- ✅ Автоматическая вставка таблиц из DataFrame или списков словарей
- ✅ Поддержка вложенных объектов `{{object.property}}`
- ✅ Сохранение форматирования Excel
- ✅ Работа с JSON данными
- ✅ Конвертация в PDF (при наличии LibreOffice)
- ✅ Вставка изображений (планируется)
- ✅ Группировка строк (планируется)

## 📦 Установка зависимостей

```bash
pip install pandas openpyxl jinja2 requests
```

## 🎯 Быстрый старт

### 1. Создание шаблона

Создайте Excel файл с переменными Jinja2:

```excel
A1: {{report_title}}
A2: Дата: {{report_date}}
A4: Всего сотрудников: {{summary.total_employees}}
A5: Средняя зарплата: {{summary.avg_salary}} руб.

A10: {{departments_start}}  # Маркер для вставки таблицы
A15: {{employees_start}}    # Маркер для другой таблицы
```

### 2. Подготовка данных

```python
data = {
    "report_title": "Отчет по сотрудникам",
    "report_date": "28.05.2024",
    "summary": {
        "total_employees": 6,
        "avg_salary": 65000
    },
    "departments": [
        {"name": "IT отдел", "count": 3, "avg_salary": 80000},
        {"name": "Продажи", "count": 2, "avg_salary": 50000}
    ],
    "employees": [
        {"name": "Иван", "position": "Разработчик", "salary": 90000},
        {"name": "Мария", "position": "Менеджер", "salary": 60000}
    ]
}
```

### 3. Генерация отчета

```python
from excel_gen import generate_report

# Из словаря
output_file = generate_report("my_template", data, "xlsx")

# Из JSON строки
import json
json_data = json.dumps(data)
output_file = generate_report("my_template", json_data, "xlsx")

print(f"Отчет создан: {output_file}")
```

## 📋 Подробное руководство

### Синтаксис переменных

| Тип | Пример | Описание |
|-----|--------|----------|
| Простая переменная | `{{title}}` | Подстановка значения |
| Вложенный объект | `{{user.name}}` | Доступ к свойству объекта |
| Ссылка | `{{website_url}}` | Автоматически становится гиперссылкой |

### Вставка таблиц

Для вставки таблиц используйте маркеры `{{table_name_start}}`:

1. **В шаблоне**: поместите `{{departments_start}}` в ячейку
2. **В данных**: передайте список словарей или DataFrame:

```python
data = {
    "departments": [
        {"name": "IT", "employees": 5},
        {"name": "Sales", "employees": 3}
    ]
}
```

Таблица будет вставлена начиная с ячейки с маркером.

### Типы данных

| Python тип | Excel результат |
|------------|-----------------|
| `str` | Текст |
| `int`, `float` | Число |
| `list[dict]` | Таблица |
| `pd.DataFrame` | Таблица |
| URL строка | Гиперссылка |

## 🛠️ API Reference

### `generate_report(template_name, json_data, output_format='xlsx')`

**Параметры:**
- `template_name` (str): Имя файла шаблона без расширения
- `json_data` (str|dict): Данные в формате JSON строки или словаря
- `output_format` (str): Формат вывода ('xlsx' или 'pdf')

**Возвращает:** Путь к созданному файлу

### `create_simple_report(title, data_dict, output_path=None)`

Создание простого отчета без шаблона.

**Параметры:**
- `title` (str): Заголовок отчета
- `data_dict` (dict): Данные для отчета
- `output_path` (str, optional): Путь для сохранения

## 📊 Примеры использования

### Пример 1: Простой отчет

```python
from excel_gen import create_simple_report

data = {
    "total_sales": 150000,
    "products": [
        {"name": "Товар А", "sales": 50000},
        {"name": "Товар Б", "sales": 100000}
    ]
}

report_path = create_simple_report("Отчет по продажам", data)
print(f"Простой отчет: {report_path}")
```

### Пример 2: Сложный шаблон

```python
from excel_gen import generate_report
import json

# Сложные данные
complex_data = {
    "company": "ООО Рога и Копыта",
    "period": "Q1 2024",
    "metrics": {
        "revenue": 1500000,
        "profit": 300000,
        "growth": 15.5
    },
    "departments": [
        {
            "name": "Разработка",
            "budget": 500000,
            "employees": [
                {"name": "Алексей", "role": "Senior Dev", "salary": 120000},
                {"name": "Мария", "role": "Junior Dev", "salary": 80000}
            ]
        }
    ]
}

# Генерация
output = generate_report("quarterly_template", complex_data)
```

### Пример 3: Работа с pandas

```python
import pandas as pd
from excel_gen import generate_report

# Создание DataFrame
df_sales = pd.DataFrame({
    'Месяц': ['Январь', 'Февраль', 'Март'],
    'Продажи': [100000, 120000, 110000],
    'Прибыль': [20000, 25000, 22000]
})

data = {
    "report_title": "Квартальный отчет",
    "sales_data": df_sales,  # DataFrame автоматически обработается
    "summary": {
        "total_sales": df_sales['Продажи'].sum(),
        "avg_profit": df_sales['Прибыль'].mean()
    }
}

output = generate_report("sales_template", data)
```

## 🔧 Расширенные возможности

### Специальные инструкции (в разработке)

```excel
{%excel image "company_logo"%}     # Вставка изображения
{%excel group "2" "5" "1"%}        # Группировка строк 2-5, уровень 1
```

### Условная логика Jinja2

```excel
{% if summary.profit > 0 %}
Прибыль: {{summary.profit}} руб.
{% else %}
Убыток: {{summary.loss}} руб.
{% endif %}
```

### Циклы Jinja2

```excel
{% for dept in departments %}
{{dept.name}}: {{dept.employee_count}} сотрудников
{% endfor %}
```

## 🐛 Устранение неполадок

### Частые ошибки

1. **FileNotFoundError**: Убедитесь, что файл шаблона существует
2. **JSON ошибки**: Проверьте корректность JSON данных
3. **Пустые ячейки**: Убедитесь, что маркеры точно соответствуют именам в данных

### Отладка

```python
# Включение подробного вывода ошибок
import traceback

try:
    output = generate_report("template", data)
except Exception as e:
    print(f"Ошибка: {e}")
    traceback.print_exc()
```

## 📝 Лицензия

MIT License

## 🤝 Вклад в проект

1. Форкните репозиторий
2. Создайте ветку для новой функции
3. Внесите изменения
4. Создайте Pull Request

---

**Автор:** AI Assistant  
**Версия:** 1.0.0  
**Дата:** 28.05.2024 