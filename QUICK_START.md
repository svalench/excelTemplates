# 🚀 Быстрый старт: Расширенный генератор отчетов

## Установка и запуск

### 1. Активация окружения
```bash
source .venv/bin/activate
pip install pandas openpyxl jinja2
```

### 2. Простейший пример
```python
from advanced_report_generator import AdvancedExcelRenderer

# Данные для отчета
data = {
    "title": "Мой первый сложный отчет",
    "subtitle": "Демонстрация возможностей",
    "summary": {
        "total_sales": 1000000,
        "orders_count": 500,
        "avg_order": 2000
    },
    "sections": [
        {
            "title": "📊 Продажи по регионам",
            "type": "table",
            "collapsed": False,
            "data": [
                {"region": "Москва", "sales": 600000, "growth": 15.2},
                {"region": "СПб", "sales": 400000, "growth": 8.7}
            ]
        }
    ]
}

# Создание отчета
renderer = AdvancedExcelRenderer()
workbook = renderer.create_collapsible_report(data)
renderer.save_report("my_first_advanced_report.xlsx")
```

### 3. Запуск примеров
```bash
# Все примеры сразу
python3 example_complex_report.py

# Интеграция с базовым генератором
python3 integration_example.py

# Базовая демонстрация
python3 advanced_report_generator.py
```

## Основные типы секций

### Таблица с фильтрами
```python
{
    "title": "Данные с фильтрами",
    "type": "table",
    "collapsed": False,
    "data": [
        {"name": "Товар 1", "price": 1000, "qty": 10},
        {"name": "Товар 2", "price": 1500, "qty": 5}
    ]
}
```

### Группированные данные
```python
{
    "title": "Анализ по категориям",
    "type": "grouped_data",
    "collapsed": True,
    "groups": [
        {
            "title": "Категория А",
            "collapsed": False,
            "data": [{"item": "A1", "value": 100}]
        }
    ]
}
```

### График
```python
{
    "title": "Динамика продаж",
    "type": "chart",
    "chart_type": "line",  # "bar", "line", "pie"
    "collapsed": False,
    "data": [
        {"month": "Янв", "sales": 100000},
        {"month": "Фев", "sales": 120000}
    ]
}
```

## Возможности

- ✅ **Сворачиваемые секции** - кнопки ▼/▶
- ✅ **Автофильтры** - в каждой таблице
- ✅ **Условное форматирование** - цветовая шкала
- ✅ **Многоуровневая группировка** - до 8 уровней
- ✅ **Интерактивные графики** - bar/line/pie
- ✅ **Профессиональный дизайн** - именованные стили

## Интеграция с существующим проектом

```python
from main import create_simple_report
from advanced_report_generator import AdvancedExcelRenderer

# Используйте базовый для простых отчетов
simple_report = create_simple_report("Простой", data)

# Используйте расширенный для сложных
renderer = AdvancedExcelRenderer()
complex_report = renderer.create_collapsible_report(complex_data)
```

## Полная документация

- **README_advanced_reports.md** - Подробное руководство
- **SUMMARY_ADVANCED_FEATURES.md** - Итоговый отчет
- **example_complex_report.py** - Примеры использования

---

**Готово к использованию!** 🎉 