# Расширенный генератор Excel отчетов

## Обзор

Расширенный генератор отчетов предоставляет мощные возможности для создания сложных Excel отчетов с:

- 🔽 **Сворачиваемыми секциями** - интерактивные разделы с кнопками ▼/▶
- 🔍 **Автофильтрами** - встроенная фильтрация для всех таблиц
- 🎨 **Условным форматированием** - автоматическая цветовая индикация
- 📊 **Многоуровневой группировкой** - иерархическая структура данных
- 📈 **Интерактивными графиками** - встроенные диаграммы
- 🖼️ **Изображениями и рисованием** - поддержка изображений в секциях и таблицах
- 🎯 **Шаблонизацией Jinja2** - гибкое управление содержимым

## Быстрый старт

### Установка зависимостей

```bash
pip install pandas openpyxl jinja2 Pillow requests
```

### Базовый пример

```python
from advanced_report_generator import AdvancedExcelRenderer

# Данные для отчета
data = {
    "title": "Мой отчет",
    "subtitle": "Период: 01.01.2024 - 31.12.2024",
    "summary": {
        "total_revenue": 1000000,
        "total_orders": 500,
        "avg_order_value": 2000
    },
    "sections": [
        {
            "title": "Продажи по регионам",
            "type": "table",
            "collapsed": False,
            "data": [
                {"region": "Москва", "sales": 500000},
                {"region": "СПб", "sales": 300000}
            ]
        }
    ]
}

# Создание отчета
renderer = AdvancedExcelRenderer()
workbook = renderer.create_collapsible_report(data)
renderer.save_report("my_report.xlsx")
```

## Структура данных

### Основная структура

```python
report_data = {
    "title": "Заголовок отчета",           # Основной заголовок
    "subtitle": "Подзаголовок",            # Дополнительная информация
    "summary": {                           # Основные метрики (всегда видны)
        "metric1": value1,
        "metric2": value2
    },
    "sections": [                          # Сворачиваемые секции
        # Секции различных типов
    ]
}
```

### Типы секций

#### 1. Таблица с автофильтрами

```python
{
    "title": "Название секции",
    "type": "table",
    "collapsed": False,                    # True = свернута по умолчанию
    "data": [
        {"column1": "value1", "column2": 100},
        {"column1": "value2", "column2": 200}
    ]
}
```

**Возможности:**
- Автоматические фильтры в заголовках
- Условное форматирование числовых колонок
- Стилизация таблиц Excel
- Автоподбор ширины колонок

#### 2. Группированные данные

```python
{
    "title": "Анализ по категориям",
    "type": "grouped_data",
    "collapsed": True,
    "groups": [
        {
            "title": "Группа 1",
            "collapsed": False,
            "data": [
                {"item": "Item 1", "value": 100},
                {"item": "Item 2", "value": 200}
            ]
        },
        {
            "title": "Группа 2", 
            "collapsed": True,
            "data": [...]
        }
    ]
}
```

**Возможности:**
- Многоуровневая группировка (до 8 уровней)
- Независимое сворачивание групп
- Визуальные отступы для иерархии
- Индикаторы сворачивания ▼/▶

#### 3. Графики и диаграммы

```python
{
    "title": "Динамика продаж",
    "type": "chart",
    "chart_type": "line",                  # "bar", "line", "pie"
    "collapsed": False,
    "data": [
        {"period": "Q1", "sales": 1000000},
        {"period": "Q2", "sales": 1200000}
    ]
}
```

**Поддерживаемые типы графиков:**
- `bar` - столбчатая диаграмма
- `line` - линейный график
- `pie` - круговая диаграмма

#### 4. Секции с изображениями

```python
{
    "title": "Логотип компании",
    "type": "image",
    "collapsed": False,
    "image_config": {
        "type": "url",                     # "url", "base64", "file"
        "source": "https://example.com/logo.png",
        "width": 300,
        "height": 150,
        "anchor": "B5",                    # Ячейка для размещения
        "description": "Описание изображения"
    }
}
```

#### 5. Программное рисование

```python
{
    "title": "Диаграмма процесса",
    "type": "drawing",
    "collapsed": False,
    "drawing_config": {
        "type": "diagram",                 # "diagram", "infographic", "custom"
        "style": "flow",                   # "boxes", "circles", "flow"
        "title": "Процесс разработки",
        "width": 600,
        "height": 300,
        "data": [
            {"label": "Планирование", "color": "#4472C4"},
            {"label": "Разработка", "color": "#70AD47"},
            {"label": "Тестирование", "color": "#FFC000"}
        ]
    }
}
```

#### 6. Таблицы с изображениями

```python
{
    "title": "Компании с логотипами",
    "type": "table",
    "collapsed": False,
    "image_columns": ["logo", "qr_code"],  # Колонки с изображениями
    "data": [
        {
            "company": "Компания А",
            "logo": {
                "type": "url",
                "source": "https://example.com/logo-a.png",
                "width": 80,
                "height": 60
            },
            "revenue": 1500000,
            "qr_code": "https://example.com/qr-a.png"  # Простой URL
        }
    ]
}
```

**Поддерживаемые типы изображений:**
- `url` - загрузка из интернета
- `base64` - встроенные изображения
- `file` - локальные файлы
- Простые URL строки

**Типы рисования:**
- `diagram` - диаграммы (boxes, circles, flow)
- `infographic` - инфографика с метриками и прогресс-барами
- `custom` - пользовательские рисунки с командами

📖 **Подробная документация**: [IMAGE_FUNCTIONALITY.md](IMAGE_FUNCTIONALITY.md)

## Использование шаблонов Jinja2

### Создание шаблона

```python
from advanced_report_generator import create_complex_report_template

# Создает файл complex_report_template.json
template = create_complex_report_template()
```

### Структура шаблона

```json
{
    "title": "{{report_title}}",
    "subtitle": "Период: {{period_start}} - {{period_end}}",
    "summary": {
        "total_revenue": "{{summary.total_revenue}}",
        "total_orders": "{{summary.total_orders}}"
    },
    "sections": [
        {
            "title": "Продажи по регионам",
            "type": "table",
            "data": "{{regional_sales}}"
        }
    ]
}
```

### Рендеринг шаблона

```python
from advanced_report_generator import render_template_with_data

# Данные для подстановки
context_data = {
    "report_title": "Отчет по продажам",
    "period_start": "01.01.2024",
    "period_end": "31.12.2024",
    "summary": {
        "total_revenue": 1000000,
        "total_orders": 500
    },
    "regional_sales": [
        {"region": "Москва", "sales": 500000},
        {"region": "СПб", "sales": 300000}
    ]
}

# Рендеринг
rendered_template = render_template_with_data(template, context_data)

# Создание отчета
renderer = AdvancedExcelRenderer()
workbook = renderer.create_collapsible_report(rendered_template)
```

## Продвинутые возможности

### Условное форматирование

Автоматически применяется к числовым колонкам:
- 🔴 Красный - минимальные значения
- 🟡 Желтый - средние значения  
- 🟢 Зеленый - максимальные значения

### Стили и форматирование

Встроенные стили:
- `header_style` - заголовки таблиц
- `subheader_style` - заголовки секций
- `data_style` - обычные данные
- `number_style` - числовые данные с форматированием

### Группировка строк

```python
# Автоматическая группировка по уровням
renderer._setup_row_grouping(
    start_row=5, 
    end_row=10, 
    level=1,        # Уровень группировки (1-8)
    hidden=False    # Скрыть по умолчанию
)
```

### Закрепление областей

Автоматически закрепляется область с заголовками (строка 4).

## Примеры использования

### 1. Финансовый отчет

```python
from example_complex_report import example_2_financial_dashboard

# Создает финансовую панель управления с:
# - Доходами и расходами по месяцам
# - Анализом по подразделениям  
# - Динамикой ключевых показателей
# - Детальным анализом транзакций

output_file = example_2_financial_dashboard()
```

### 2. Аналитика продаж

```python
from example_complex_report import example_3_sales_analytics

# Создает отчет по продажам с:
# - Продажами по регионам и каналам
# - Сегментацией клиентов
# - Трендами по категориям
# - Детальными данными заказов

output_file = example_3_sales_analytics()
```

### 3. Операционный отчет

```python
from example_complex_report import example_4_operational_report

# Создает операционный отчет с:
# - KPI по подразделениям
# - Производственными показателями
# - Трендами эффективности
# - Журналом инцидентов

output_file = example_4_operational_report()
```

## Интеграция с существующим проектом

### Расширение основного генератора

```python
from main import ExcelRenderer
from advanced_report_generator import AdvancedExcelRenderer

# Можно использовать оба генератора
basic_renderer = ExcelRenderer()
advanced_renderer = AdvancedExcelRenderer()

# Для простых отчетов
basic_renderer.render_template("simple_template.xlsx", "output.xlsx", data)

# Для сложных отчетов
advanced_renderer.create_collapsible_report(complex_data)
```

### Миграция существующих шаблонов

1. Конвертируйте данные в новый формат:

```python
# Старый формат
old_data = {
    "sales_data": [{"region": "Москва", "sales": 1000}]
}

# Новый формат
new_data = {
    "title": "Отчет по продажам",
    "sections": [{
        "title": "Продажи по регионам",
        "type": "table", 
        "data": old_data["sales_data"]
    }]
}
```

2. Добавьте группировку и сворачивание:

```python
# Преобразование в группированные данные
grouped_data = {
    "title": "Анализ по категориям",
    "type": "grouped_data",
    "groups": [
        {
            "title": category,
            "data": category_data
        } for category, category_data in grouped_by_category.items()
    ]
}
```

## Производительность и ограничения

### Рекомендации по производительности

- **Размер данных**: до 10,000 строк на таблицу
- **Количество секций**: до 20 секций на отчет
- **Группировка**: до 8 уровней вложенности
- **Графики**: до 5 графиков на отчет

### Оптимизация

```python
# Ограничение данных для больших таблиц
large_data = detailed_orders[:1000]  # Первые 1000 записей

# Сворачивание больших секций по умолчанию
{
    "title": "Большая таблица",
    "type": "table",
    "collapsed": True,  # Свернута по умолчанию
    "data": large_data
}
```

## Устранение неполадок

### Частые проблемы

1. **Ошибка "Table name already exists"**
   ```python
   # Решение: используйте уникальные имена таблиц
   table = Table(displayName=f"Table_{start_row}_{random.randint(1000,9999)}")
   ```

2. **Медленная генерация больших отчетов**
   ```python
   # Решение: ограничьте размер данных
   data = large_dataset[:500]  # Максимум 500 строк
   ```

3. **Проблемы с кодировкой**
   ```python
   # Решение: используйте UTF-8
   with open('template.json', 'w', encoding='utf-8') as f:
       json.dump(data, f, ensure_ascii=False)
   ```

### Отладка

```python
# Включение подробного логирования
import logging
logging.basicConfig(level=logging.DEBUG)

# Проверка структуры данных
print(json.dumps(report_data, indent=2, ensure_ascii=False))
```

## API Reference

### AdvancedExcelRenderer

#### Методы

- `create_collapsible_report(data, template_config=None)` - создание отчета
- `save_report(filename)` - сохранение в файл
- `create_styles()` - создание стилей
- `_add_table_with_filters(section_data, start_row)` - добавление таблицы
- `_add_grouped_data(section_data, start_row)` - добавление групп
- `_add_chart_section(section_data, start_row)` - добавление графика

#### Параметры

- `data` - словарь с данными отчета
- `template_config` - конфигурация шаблона (опционально)
- `filename` - путь для сохранения файла

### Вспомогательные функции

- `create_complex_report_template()` - создание шаблона
- `generate_sample_data()` - генерация тестовых данных
- `render_template_with_data(template, context)` - рендеринг шаблона

## Лицензия

MIT License - см. файл LICENSE для деталей.

## Поддержка

Для вопросов и предложений создавайте issues в репозитории проекта. 