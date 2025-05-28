#!/usr/bin/env python3
"""
Примеры интеграции расширенного генератора с базовым функционалом
Демонстрирует совместное использование обеих систем
"""

from main import ExcelRenderer, create_simple_report  # Базовый генератор
from advanced_report_generator import AdvancedExcelRenderer  # Расширенный генератор
import pandas as pd
from datetime import datetime, timedelta
import json
from jinja2 import Template


def example_1_template_migration():
    """Пример 1: Миграция простого шаблона в сложный отчет"""
    print("\n" + "="*60)
    print("🔄 Пример 1: Миграция простого шаблона в сложный отчет")
    print("="*60)
    
    # Простые данные
    simple_data = {
        "title": "Простой отчет",
        "date": datetime.now().strftime("%d.%m.%Y"),
        "sales_data": [
            {"region": "Москва", "sales": 1500000, "growth": 15.2},
            {"region": "СПб", "sales": 1200000, "growth": 8.7},
            {"region": "Екатеринбург", "sales": 800000, "growth": 12.1}
        ]
    }
    
    # Создаем простой отчет с базовым генератором
    simple_filename = create_simple_report("Простой отчет продаж", simple_data)
    
    # Теперь мигрируем в сложный отчет
    complex_data = {
        "title": "Мигрированный сложный отчет",
        "subtitle": f"Создан из простого шаблона | {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        "summary": {
            "total_sales": sum(item["sales"] for item in simple_data["sales_data"]),
            "avg_growth": sum(item["growth"] for item in simple_data["sales_data"]) / len(simple_data["sales_data"]),
            "regions_count": len(simple_data["sales_data"])
        },
        "sections": [
            {
                "title": "📊 Детальные данные по регионам",
                "type": "table",
                "collapsed": False,
                "data": simple_data["sales_data"]
            }
        ]
    }
    
    # Создаем сложный отчет
    advanced_renderer = AdvancedExcelRenderer()
    complex_wb = advanced_renderer.create_collapsible_report(complex_data)
    complex_filename = f"integration_migrated_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    advanced_renderer.save_report(complex_filename)
    
    print(f"✅ Простой отчет: {simple_filename}")
    print(f"✅ Сложный отчет: {complex_filename}")
    print("📋 Демонстрирует:")
    print("  • Миграцию данных из простого формата")
    print("  • Добавление сводных метрик")
    print("  • Преобразование в интерактивный формат")
    
    return simple_filename, complex_filename


def example_2_hybrid_report():
    """Пример 2: Гибридный отчет с использованием обеих систем"""
    print("\n" + "="*60)
    print("🔀 Пример 2: Гибридный отчет с использованием обеих систем")
    print("="*60)
    
    # Данные для гибридного отчета
    hybrid_data = {
        "company": "ООО 'Инновации'",
        "period": "Q1 2025",
        "prepared_by": "Аналитический отдел",
        "summary_metrics": {
            "revenue": 15000000,
            "profit": 3500000,
            "employees": 250,
            "projects": 45
        },
        "departments": [
            {
                "name": "Разработка",
                "budget": 5000000,
                "actual": 4800000,
                "efficiency": 96.0,
                "projects": [
                    {"name": "Проект A", "status": "Завершен", "budget": 1500000, "actual": 1450000},
                    {"name": "Проект B", "status": "В работе", "budget": 2000000, "actual": 1800000},
                    {"name": "Проект C", "status": "Планируется", "budget": 1500000, "actual": 0}
                ]
            },
            {
                "name": "Маркетинг",
                "budget": 3000000,
                "actual": 2900000,
                "efficiency": 96.7,
                "campaigns": [
                    {"name": "Кампания 1", "reach": 50000, "conversions": 1250, "cost": 800000},
                    {"name": "Кампания 2", "reach": 75000, "conversions": 1875, "cost": 1200000},
                    {"name": "Кампания 3", "reach": 30000, "conversions": 600, "cost": 500000}
                ]
            }
        ]
    }
    
    # Создаем сложный отчет с расширенным генератором
    complex_report_data = {
        "title": f"Отчет {hybrid_data['company']} за {hybrid_data['period']}",
        "subtitle": f"Подготовлен: {hybrid_data['prepared_by']} | {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        "summary": hybrid_data["summary_metrics"],
        "sections": []
    }
    
    # Добавляем секции по департаментам
    for dept in hybrid_data["departments"]:
        if dept["name"] == "Разработка":
            complex_report_data["sections"].append({
                "title": f"💻 Департамент: {dept['name']}",
                "type": "grouped_data",
                "collapsed": False,
                "groups": [
                    {
                        "title": "Проекты",
                        "collapsed": False,
                        "data": dept["projects"]
                    }
                ]
            })
        elif dept["name"] == "Маркетинг":
            complex_report_data["sections"].append({
                "title": f"📢 Департамент: {dept['name']}",
                "type": "table",
                "collapsed": False,
                "data": dept["campaigns"]
            })
    
    # Создаем основной отчет
    advanced_renderer = AdvancedExcelRenderer()
    workbook = advanced_renderer.create_collapsible_report(complex_report_data)
    
    # Добавляем дополнительный лист с базовым генератором
    # Создаем шаблон для детального анализа
    detail_template = """
ДЕТАЛЬНЫЙ АНАЛИЗ {{ company }}
Период: {{ period }}

Общие показатели:
- Выручка: {{ summary_metrics.revenue }}
- Прибыль: {{ summary_metrics.profit }}
- Сотрудники: {{ summary_metrics.employees }}
- Проекты: {{ summary_metrics.projects }}

Анализ по департаментам:
{% for dept in departments %}

{{ dept.name }}:
- Бюджет: {{ dept.budget }}
- Фактически: {{ dept.actual }}
- Эффективность: {{ dept.efficiency }}%
{% endfor %}
"""
    
    # Добавляем новый лист
    detail_ws = workbook.create_sheet("Детальный анализ")
    
    # Рендерим шаблон
    template = Template(detail_template)
    rendered_content = template.render(**hybrid_data)
    
    # Добавляем содержимое в лист
    lines = rendered_content.split('\n')
    for row, line in enumerate(lines, 1):
        detail_ws.cell(row=row, column=1, value=line.strip())
    
    # Сохраняем гибридный отчет
    filename = f"integration_complex_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    workbook.save(filename)
    
    print(f"✅ Гибридный отчет: {filename}")
    print("📋 Особенности:")
    print("  • Основной лист: расширенный генератор")
    print("  • Дополнительный лист: базовый генератор")
    print("  • Интерактивные элементы + текстовый анализ")
    print("  • Многоуровневая структура данных")
    
    return filename


def example_3_performance_comparison():
    """Пример 3: Сравнение производительности генераторов"""
    print("\n" + "="*60)
    print("⚡ Пример 3: Сравнение производительности генераторов")
    print("="*60)
    
    import time
    
    # Генерируем большой объем данных
    large_dataset = []
    for i in range(1000):
        large_dataset.append({
            "id": f"ID-{i:04d}",
            "name": f"Элемент {i}",
            "value": i * 1.5,
            "category": f"Категория {i % 10}",
            "date": (datetime.now() - timedelta(days=i % 365)).strftime("%d.%m.%Y")
        })
    
    # Тест базового генератора
    print("🔄 Тестирование базового генератора...")
    start_time = time.time()
    
    base_data = {"items": large_dataset[:100]}  # Ограничиваем для базового
    base_filename = create_simple_report("Тест производительности - Базовый", base_data)
    base_time = time.time() - start_time
    
    # Тест расширенного генератора
    print("🔄 Тестирование расширенного генератора...")
    start_time = time.time()
    
    advanced_renderer = AdvancedExcelRenderer()
    advanced_data = {
        "title": "Тест производительности",
        "subtitle": f"Обработка {len(large_dataset)} записей",
        "summary": {
            "total_records": len(large_dataset),
            "avg_value": sum(item["value"] for item in large_dataset) / len(large_dataset),
            "categories": len(set(item["category"] for item in large_dataset))
        },
        "sections": [
            {
                "title": "📊 Полный набор данных",
                "type": "table",
                "collapsed": True,
                "data": large_dataset
            }
        ]
    }
    
    advanced_wb = advanced_renderer.create_collapsible_report(advanced_data)
    advanced_time = time.time() - start_time
    
    # Сохраняем результаты
    advanced_filename = f"performance_advanced_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    advanced_renderer.save_report(advanced_filename)
    
    # Результаты
    performance_data = {
        "title": "Результаты тестирования производительности",
        "subtitle": f"Тест выполнен: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        "summary": {
            "base_time": round(base_time, 3),
            "advanced_time": round(advanced_time, 3),
            "base_records": len(base_data["items"]),
            "advanced_records": len(large_dataset)
        },
        "sections": [
            {
                "title": "📈 Сравнительная таблица",
                "type": "table",
                "collapsed": False,
                "data": [
                    {
                        "generator": "Базовый",
                        "time_seconds": round(base_time, 3),
                        "records": len(base_data["items"]),
                        "features": "Простые шаблоны"
                    },
                    {
                        "generator": "Расширенный",
                        "time_seconds": round(advanced_time, 3),
                        "records": len(large_dataset),
                        "features": "Интерактивные элементы"
                    }
                ]
            }
        ]
    }
    
    # Создаем отчет о производительности
    perf_renderer = AdvancedExcelRenderer()
    perf_wb = perf_renderer.create_collapsible_report(performance_data)
    perf_filename = f"performance_test_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    perf_renderer.save_report(perf_filename)
    
    print(f"✅ Базовый генератор: {base_filename} ({base_time:.3f}s)")
    print(f"✅ Расширенный генератор: {advanced_filename} ({advanced_time:.3f}s)")
    print(f"✅ Отчет о производительности: {perf_filename}")
    print("📋 Выводы:")
    print(f"  • Базовый: {len(base_data['items'])} записей за {base_time:.3f}s")
    print(f"  • Расширенный: {len(large_dataset)} записей за {advanced_time:.3f}s")
    print(f"  • Соотношение: {advanced_time/base_time:.1f}x при {len(large_dataset)/len(base_data['items'])}x данных")
    
    return base_filename, advanced_filename, perf_filename


def main():
    """Главная функция демонстрации интеграции"""
    print("🚀 Демонстрация интеграции генераторов Excel отчетов")
    print("=" * 60)
    
    results = []
    
    # Пример 1: Миграция шаблона
    try:
        simple_file, complex_file = example_1_template_migration()
        results.extend([simple_file, complex_file])
    except Exception as e:
        print(f"❌ Ошибка в примере 1: {e}")
    
    # Пример 2: Гибридный отчет
    try:
        hybrid_file = example_2_hybrid_report()
        results.append(hybrid_file)
    except Exception as e:
        print(f"❌ Ошибка в примере 2: {e}")
    
    # Пример 3: Тест производительности
    try:
        base_file, adv_file, perf_file = example_3_performance_comparison()
        results.extend([base_file, adv_file, perf_file])
    except Exception as e:
        print(f"❌ Ошибка в примере 3: {e}")
    
    # Итоги
    print("\n" + "="*60)
    print("📋 ИТОГИ ИНТЕГРАЦИИ")
    print("="*60)
    print(f"✅ Создано файлов: {len(results)}")
    for file in results:
        print(f"  • {file}")
    
    print("\n🎯 Ключевые возможности интеграции:")
    print("  • Совместное использование обеих систем")
    print("  • Миграция простых шаблонов в сложные")
    print("  • Гибридные отчеты с разными листами")
    print("  • Сравнение производительности")
    print("  • Полная совместимость API")


if __name__ == "__main__":
    main() 