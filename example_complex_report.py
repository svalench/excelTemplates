#!/usr/bin/env python3
"""
Примеры использования расширенного генератора отчетов
Демонстрирует различные сценарии создания сложных отчетов
"""

from advanced_report_generator import (
    AdvancedExcelRenderer, 
    create_complex_report_template,
    generate_sample_data,
    render_template_with_data
)
import pandas as pd
from datetime import datetime, timedelta
import random


def example_1_basic_complex_report():
    """Пример 1: Базовый сложный отчет"""
    print("\n" + "="*60)
    print("📊 Пример 1: Базовый сложный отчет со сворачиваемыми секциями")
    print("="*60)
    
    # Создаем шаблон
    template = create_complex_report_template()
    
    # Генерируем данные
    data = generate_sample_data()
    
    # Рендерим шаблон
    rendered_template = render_template_with_data(template, data)
    
    # Создаем отчет
    renderer = AdvancedExcelRenderer()
    workbook = renderer.create_collapsible_report(rendered_template)
    
    # Сохраняем
    output_file = f"example_1_basic_complex_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    renderer.save_report(output_file)
    
    print(f"✅ Создан: {output_file}")
    print("📋 Особенности:")
    print("  • Основные метрики всегда видны")
    print("  • Секция 'Продажи по регионам' развернута")
    print("  • Секция 'Анализ продуктов' свернута")
    print("  • Автофильтры в таблицах")
    print("  • Условное форматирование числовых данных")
    
    return output_file


def example_2_financial_dashboard():
    """Пример 2: Финансовая панель управления"""
    print("\n" + "="*60)
    print("💰 Пример 2: Финансовая панель управления")
    print("="*60)
    
    # Данные для финансовой панели
    financial_data = {
        "title": "Финансовая панель управления",
        "subtitle": f"Отчетный период: Q4 2024 | Создан: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        "summary": {
            "total_revenue": 5200000,
            "total_expenses": 3800000,
            "net_profit": 1400000,
            "profit_margin": 26.9,
            "ebitda": 1650000,
            "cash_flow": 1200000
        },
        "sections": [
            {
                "title": "💼 Доходы и расходы по месяцам",
                "type": "table",
                "collapsed": False,
                "data": [
                    {"month": "Октябрь", "revenue": 1650000, "expenses": 1200000, "profit": 450000, "margin": 27.3},
                    {"month": "Ноябрь", "revenue": 1750000, "expenses": 1300000, "profit": 450000, "margin": 25.7},
                    {"month": "Декабрь", "revenue": 1800000, "expenses": 1300000, "profit": 500000, "margin": 27.8}
                ]
            },
            {
                "title": "🏢 Анализ по подразделениям",
                "type": "grouped_data",
                "collapsed": True,
                "groups": [
                    {
                        "title": "Операционная деятельность",
                        "collapsed": False,
                        "data": [
                            {"department": "Продажи", "budget": 800000, "actual": 750000, "variance": -6.3},
                            {"department": "Маркетинг", "budget": 600000, "actual": 580000, "variance": -3.3},
                            {"department": "Производство", "budget": 1200000, "actual": 1250000, "variance": 4.2}
                        ]
                    },
                    {
                        "title": "Административные расходы",
                        "collapsed": True,
                        "data": [
                            {"department": "HR", "budget": 400000, "actual": 380000, "variance": -5.0},
                            {"department": "IT", "budget": 500000, "actual": 520000, "variance": 4.0},
                            {"department": "Финансы", "budget": 300000, "actual": 290000, "variance": -3.3}
                        ]
                    }
                ]
            },
            {
                "title": "📈 Динамика ключевых показателей",
                "type": "chart",
                "chart_type": "line",
                "collapsed": False,
                "data": [
                    {"period": "Q1", "revenue": 4200000, "profit": 1100000, "margin": 26.2},
                    {"period": "Q2", "revenue": 4600000, "profit": 1250000, "margin": 27.2},
                    {"period": "Q3", "revenue": 4800000, "profit": 1300000, "margin": 27.1},
                    {"period": "Q4", "revenue": 5200000, "profit": 1400000, "margin": 26.9}
                ]
            },
            {
                "title": "🔍 Детальный анализ транзакций",
                "type": "table",
                "collapsed": True,
                "data": generate_transaction_data(100)
            }
        ]
    }
    
    # Создаем отчет
    renderer = AdvancedExcelRenderer()
    workbook = renderer.create_collapsible_report(financial_data)
    
    # Сохраняем
    output_file = f"example_2_financial_dashboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    renderer.save_report(output_file)
    
    print(f"✅ Создан: {output_file}")
    print("📋 Особенности:")
    print("  • Финансовые метрики в сводке")
    print("  • Многоуровневая группировка подразделений")
    print("  • График динамики показателей")
    print("  • Детальные транзакции (свернуты)")
    
    return output_file


def example_3_sales_analytics():
    """Пример 3: Аналитика продаж с фильтрами"""
    print("\n" + "="*60)
    print("🛒 Пример 3: Аналитика продаж с расширенными фильтрами")
    print("="*60)
    
    # Генерируем данные продаж
    sales_data = generate_sales_analytics_data()
    
    sales_report = {
        "title": "Аналитика продаж и клиентов",
        "subtitle": f"Анализ за последние 6 месяцев | Обновлено: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        "summary": {
            "total_sales": sales_data['total_sales'],
            "unique_customers": sales_data['unique_customers'],
            "avg_order_value": sales_data['avg_order_value'],
            "repeat_customers": sales_data['repeat_customers'],
            "conversion_rate": sales_data['conversion_rate'],
            "customer_lifetime_value": sales_data['customer_lifetime_value']
        },
        "sections": [
            {
                "title": "🌍 Продажи по регионам и каналам",
                "type": "table",
                "collapsed": False,
                "data": sales_data['regional_channel_sales']
            },
            {
                "title": "👥 Сегментация клиентов",
                "type": "grouped_data",
                "collapsed": False,
                "groups": [
                    {
                        "title": "По объему покупок",
                        "collapsed": False,
                        "data": sales_data['customer_segments_volume']
                    },
                    {
                        "title": "По частоте покупок",
                        "collapsed": True,
                        "data": sales_data['customer_segments_frequency']
                    },
                    {
                        "title": "По географии",
                        "collapsed": True,
                        "data": sales_data['customer_segments_geo']
                    }
                ]
            },
            {
                "title": "📊 Тренды продаж по категориям",
                "type": "chart",
                "chart_type": "bar",
                "collapsed": False,
                "data": sales_data['category_trends']
            },
            {
                "title": "🔎 Детальные данные по заказам",
                "type": "table",
                "collapsed": True,
                "data": sales_data['detailed_orders'][:200]  # Ограничиваем для производительности
            }
        ]
    }
    
    # Создаем отчет
    renderer = AdvancedExcelRenderer()
    workbook = renderer.create_collapsible_report(sales_report)
    
    # Сохраняем
    output_file = f"example_3_sales_analytics_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    renderer.save_report(output_file)
    
    print(f"✅ Создан: {output_file}")
    print("📋 Особенности:")
    print("  • Многомерная сегментация клиентов")
    print("  • Анализ по каналам продаж")
    print("  • Тренды по категориям товаров")
    print("  • Детальные данные с автофильтрами")
    
    return output_file


def example_4_operational_report():
    """Пример 4: Операционный отчет с KPI"""
    print("\n" + "="*60)
    print("⚙️ Пример 4: Операционный отчет с KPI и метриками")
    print("="*60)
    
    # Генерируем операционные данные
    operational_data = generate_operational_data()
    
    operational_report = {
        "title": "Операционный отчет и KPI",
        "subtitle": f"Мониторинг операционной эффективности | {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        "summary": {
            "efficiency_score": operational_data['efficiency_score'],
            "quality_score": operational_data['quality_score'],
            "customer_satisfaction": operational_data['customer_satisfaction'],
            "employee_productivity": operational_data['employee_productivity'],
            "cost_per_unit": operational_data['cost_per_unit'],
            "defect_rate": operational_data['defect_rate']
        },
        "sections": [
            {
                "title": "📈 KPI по подразделениям",
                "type": "table",
                "collapsed": False,
                "data": operational_data['department_kpi']
            },
            {
                "title": "🏭 Производственные показатели",
                "type": "grouped_data",
                "collapsed": False,
                "groups": [
                    {
                        "title": "Производительность линий",
                        "collapsed": False,
                        "data": operational_data['production_lines']
                    },
                    {
                        "title": "Качество продукции",
                        "collapsed": True,
                        "data": operational_data['quality_metrics']
                    },
                    {
                        "title": "Использование ресурсов",
                        "collapsed": True,
                        "data": operational_data['resource_utilization']
                    }
                ]
            },
            {
                "title": "📊 Тренды эффективности",
                "type": "chart",
                "chart_type": "line",
                "collapsed": False,
                "data": operational_data['efficiency_trends']
            },
            {
                "title": "⚠️ Инциденты и проблемы",
                "type": "table",
                "collapsed": True,
                "data": operational_data['incidents']
            }
        ]
    }
    
    # Создаем отчет
    renderer = AdvancedExcelRenderer()
    workbook = renderer.create_collapsible_report(operational_report)
    
    # Сохраняем
    output_file = f"example_4_operational_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    renderer.save_report(output_file)
    
    print(f"✅ Создан: {output_file}")
    print("📋 Особенности:")
    print("  • KPI-дашборд по подразделениям")
    print("  • Производственные метрики")
    print("  • Тренды эффективности")
    print("  • Журнал инцидентов")
    
    return output_file


def generate_transaction_data(count=100):
    """Генерация данных транзакций"""
    transactions = []
    for i in range(count):
        transactions.append({
            "transaction_id": f"TXN-{10000+i}",
            "date": (datetime.now() - timedelta(days=random.randint(1, 90))).strftime("%d.%m.%Y"),
            "amount": random.randint(1000, 50000),
            "type": random.choice(["Доход", "Расход"]),
            "category": random.choice(["Продажи", "Закупки", "Зарплата", "Аренда", "Маркетинг"]),
            "description": f"Операция {i+1}"
        })
    return transactions


def generate_sales_analytics_data():
    """Генерация данных для аналитики продаж"""
    
    # Основные метрики
    total_sales = 8500000
    unique_customers = 2150
    avg_order_value = 3953
    repeat_customers = 1290
    conversion_rate = 12.5
    customer_lifetime_value = 15600
    
    # Продажи по регионам и каналам
    regional_channel_sales = [
        {"region": "Москва", "online": 1200000, "retail": 800000, "b2b": 600000, "total": 2600000},
        {"region": "СПб", "online": 900000, "retail": 600000, "b2b": 400000, "total": 1900000},
        {"region": "Екатеринбург", "online": 600000, "retail": 400000, "b2b": 300000, "total": 1300000},
        {"region": "Новосибирск", "online": 500000, "retail": 350000, "b2b": 250000, "total": 1100000},
        {"region": "Другие", "online": 800000, "retail": 500000, "b2b": 300000, "total": 1600000}
    ]
    
    # Сегментация клиентов
    customer_segments_volume = [
        {"segment": "VIP (>100k)", "customers": 85, "avg_order": 25000, "total_revenue": 2125000},
        {"segment": "Премиум (50-100k)", "customers": 215, "avg_order": 15000, "total_revenue": 3225000},
        {"segment": "Стандарт (10-50k)", "customers": 860, "avg_order": 3500, "total_revenue": 3010000},
        {"segment": "Базовый (<10k)", "customers": 990, "avg_order": 1400, "total_revenue": 1386000}
    ]
    
    customer_segments_frequency = [
        {"segment": "Постоянные (>10 заказов)", "customers": 320, "avg_frequency": 15, "retention": 95},
        {"segment": "Регулярные (5-10 заказов)", "customers": 645, "avg_frequency": 7, "retention": 78},
        {"segment": "Периодические (2-4 заказа)", "customers": 825, "avg_frequency": 3, "retention": 45},
        {"segment": "Разовые (1 заказ)", "customers": 360, "avg_frequency": 1, "retention": 12}
    ]
    
    customer_segments_geo = [
        {"region": "Центральный ФО", "customers": 1200, "avg_order": 4200, "penetration": 8.5},
        {"region": "Северо-Западный ФО", "customers": 450, "avg_order": 3800, "penetration": 6.2},
        {"region": "Уральский ФО", "customers": 300, "avg_order": 3500, "penetration": 4.8},
        {"region": "Сибирский ФО", "customers": 200, "avg_order": 3200, "penetration": 3.1}
    ]
    
    # Тренды по категориям
    category_trends = [
        {"category": "Электроника", "q1": 1200000, "q2": 1350000, "q3": 1400000, "q4": 1500000},
        {"category": "Одежда", "q1": 800000, "q2": 900000, "q3": 950000, "q4": 1100000},
        {"category": "Дом и сад", "q1": 600000, "q2": 750000, "q3": 800000, "q4": 900000},
        {"category": "Спорт", "q1": 400000, "q2": 500000, "q3": 550000, "q4": 600000},
        {"category": "Книги", "q1": 200000, "q2": 250000, "q3": 280000, "q4": 320000}
    ]
    
    # Детальные заказы
    detailed_orders = []
    for i in range(500):
        detailed_orders.append({
            "order_id": f"ORD-{20000+i}",
            "customer_id": f"CUST-{1000+random.randint(1, 2150)}",
            "date": (datetime.now() - timedelta(days=random.randint(1, 180))).strftime("%d.%m.%Y"),
            "amount": random.randint(500, 25000),
            "category": random.choice(["Электроника", "Одежда", "Дом и сад", "Спорт", "Книги"]),
            "channel": random.choice(["Online", "Retail", "B2B"]),
            "region": random.choice(["Москва", "СПб", "Екатеринбург", "Новосибирск", "Другие"]),
            "status": random.choice(["Выполнен", "В обработке", "Отправлен", "Отменен"])
        })
    
    return {
        "total_sales": total_sales,
        "unique_customers": unique_customers,
        "avg_order_value": avg_order_value,
        "repeat_customers": repeat_customers,
        "conversion_rate": conversion_rate,
        "customer_lifetime_value": customer_lifetime_value,
        "regional_channel_sales": regional_channel_sales,
        "customer_segments_volume": customer_segments_volume,
        "customer_segments_frequency": customer_segments_frequency,
        "customer_segments_geo": customer_segments_geo,
        "category_trends": category_trends,
        "detailed_orders": detailed_orders
    }


def generate_operational_data():
    """Генерация операционных данных"""
    
    # Основные KPI
    efficiency_score = 87.5
    quality_score = 94.2
    customer_satisfaction = 4.3
    employee_productivity = 112.8
    cost_per_unit = 245.50
    defect_rate = 0.8
    
    # KPI по подразделениям
    department_kpi = [
        {"department": "Производство", "efficiency": 89, "quality": 96, "cost": 220, "target": 225},
        {"department": "Логистика", "efficiency": 92, "quality": 88, "cost": 45, "target": 50},
        {"department": "Продажи", "efficiency": 85, "quality": 91, "cost": 180, "target": 175},
        {"department": "Поддержка", "efficiency": 78, "quality": 95, "cost": 120, "target": 115},
        {"department": "R&D", "efficiency": 82, "quality": 98, "cost": 350, "target": 340}
    ]
    
    # Производственные линии
    production_lines = [
        {"line": "Линия A", "capacity": 1000, "actual": 920, "efficiency": 92, "downtime": 8},
        {"line": "Линия B", "capacity": 800, "actual": 760, "efficiency": 95, "downtime": 5},
        {"line": "Линия C", "capacity": 1200, "actual": 1080, "efficiency": 90, "downtime": 10},
        {"line": "Линия D", "capacity": 600, "actual": 540, "efficiency": 90, "downtime": 10}
    ]
    
    # Метрики качества
    quality_metrics = [
        {"metric": "Дефекты на миллион", "value": 850, "target": 1000, "trend": "↓"},
        {"metric": "Первый проход качества", "value": 96.5, "target": 95, "trend": "↑"},
        {"metric": "Возвраты клиентов", "value": 0.3, "target": 0.5, "trend": "↓"},
        {"metric": "Время устранения дефектов", "value": 2.1, "target": 3, "trend": "↓"}
    ]
    
    # Использование ресурсов
    resource_utilization = [
        {"resource": "Оборудование", "utilization": 87, "capacity": 100, "efficiency": 92},
        {"resource": "Персонал", "utilization": 95, "capacity": 100, "efficiency": 88},
        {"resource": "Материалы", "utilization": 78, "capacity": 85, "efficiency": 94},
        {"resource": "Энергия", "utilization": 82, "capacity": 90, "efficiency": 89}
    ]
    
    # Тренды эффективности
    efficiency_trends = []
    for i in range(12):
        month = (datetime.now() - timedelta(days=30*(11-i))).strftime("%B")
        efficiency_trends.append({
            "month": month,
            "efficiency": random.uniform(80, 95),
            "quality": random.uniform(90, 98),
            "cost": random.uniform(200, 300)
        })
    
    # Инциденты
    incidents = []
    for i in range(30):
        incidents.append({
            "incident_id": f"INC-{3000+i}",
            "date": (datetime.now() - timedelta(days=random.randint(1, 60))).strftime("%d.%m.%Y"),
            "type": random.choice(["Качество", "Безопасность", "Оборудование", "Процесс"]),
            "severity": random.choice(["Низкая", "Средняя", "Высокая", "Критическая"]),
            "status": random.choice(["Открыт", "В работе", "Решен", "Закрыт"]),
            "department": random.choice(["Производство", "Логистика", "Качество", "Безопасность"])
        })
    
    return {
        "efficiency_score": efficiency_score,
        "quality_score": quality_score,
        "customer_satisfaction": customer_satisfaction,
        "employee_productivity": employee_productivity,
        "cost_per_unit": cost_per_unit,
        "defect_rate": defect_rate,
        "department_kpi": department_kpi,
        "production_lines": production_lines,
        "quality_metrics": quality_metrics,
        "resource_utilization": resource_utilization,
        "efficiency_trends": efficiency_trends,
        "incidents": incidents
    }


def main():
    """Запуск всех примеров"""
    print("🚀 Демонстрация расширенного генератора отчетов")
    print("="*60)
    
    examples = [
        example_1_basic_complex_report,
        example_2_financial_dashboard,
        example_3_sales_analytics,
        example_4_operational_report
    ]
    
    created_files = []
    
    for example_func in examples:
        try:
            output_file = example_func()
            created_files.append(output_file)
        except Exception as e:
            print(f"❌ Ошибка в {example_func.__name__}: {e}")
    
    print("\n" + "="*60)
    print("📋 ИТОГИ ДЕМОНСТРАЦИИ")
    print("="*60)
    print(f"✅ Создано файлов: {len(created_files)}")
    for file in created_files:
        print(f"  • {file}")
    
    print("\n🎯 Ключевые возможности:")
    print("  • Сворачиваемые секции с визуальными индикаторами")
    print("  • Автофильтры для всех таблиц")
    print("  • Условное форматирование числовых данных")
    print("  • Многоуровневая группировка строк")
    print("  • Интерактивные графики")
    print("  • Автоматическое форматирование")
    print("  • Закрепление областей")
    print("  • Поддержка шаблонов Jinja2")


if __name__ == "__main__":
    main() 