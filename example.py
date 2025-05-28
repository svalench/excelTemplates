from advanced_report_generator import AdvancedExcelRenderer

# Данные для отчета с поддержкой сворачиваемых строк
data = {
    "title": "Отчет от 12221",
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
                {"region": "Москва", "sales": 500000, "orders": 250, "growth": 15.2},
                {"region": "СПб", "sales": 300000, "orders": 150, "growth": 8.7},
                {"region": "Екатеринбург", "sales": 150000, "orders": 75, "growth": 12.1},
                {"region": "Новосибирск", "sales": 50000, "orders": 25, "growth": 5.3},
            ]
        },
        {
            "title": "Детальный анализ по категориям",
            "type": "grouped_data",
            "collapsed": True,  # Эта секция будет свернута
            "groups": [
                {
                    "title": "Электроника",
                    "collapsed": False,
                    "data": [
                        {"product": "Смартфоны", "sales": 200000, "units": 100, "margin": 25.5},
                        {"product": "Ноутбуки", "sales": 150000, "units": 50, "margin": 18.2},
                        {"product": "Планшеты", "sales": 100000, "units": 80, "margin": 22.1}
                    ]
                },
                {
                    "title": "Одежда",
                    "collapsed": True,  # Эта группа будет свернута
                    "data": [
                        {"product": "Куртки", "sales": 80000, "units": 160, "margin": 45.2},
                        {"product": "Джинсы", "sales": 60000, "units": 120, "margin": 38.5},
                        {"product": "Футболки", "sales": 40000, "units": 200, "margin": 52.1}
                    ]
                },
                {
                    "title": "Книги",
                    "collapsed": False,
                    "data": [
                        {"product": "Художественная литература", "sales": 30000, "units": 300, "margin": 35.2},
                        {"product": "Техническая литература", "sales": 25000, "units": 125, "margin": 28.5},
                        {"product": "Детские книги", "sales": 15000, "units": 150, "margin": 42.1}
                    ]
                }
            ]
        },
        {
            "title": "Динамика продаж по месяцам",
            "type": "chart",
            "chart_type": "line",
            "collapsed": False,
            "data": [
                {"month": "Январь", "sales": 80000, "orders": 40},
                {"month": "Февраль", "sales": 85000, "orders": 42},
                {"month": "Март", "sales": 90000, "orders": 45},
                {"month": "Апрель", "sales": 95000, "orders": 48},
                {"month": "Май", "sales": 100000, "orders": 50},
                {"month": "Июнь", "sales": 105000, "orders": 52}
            ]
        },
        {
            "title": "Подробные транзакции",
            "type": "table",
            "collapsed": True,  # Эта секция будет свернута
            "data": [
                {"transaction_id": "TXN-001", "date": "15.01.2024", "amount": 5000, "customer": "Клиент А", "status": "Выполнен"},
                {"transaction_id": "TXN-002", "date": "16.01.2024", "amount": 3500, "customer": "Клиент Б", "status": "Выполнен"},
                {"transaction_id": "TXN-003", "date": "17.01.2024", "amount": 7200, "customer": "Клиент В", "status": "В обработке"},
                {"transaction_id": "TXN-004", "date": "18.01.2024", "amount": 2800, "customer": "Клиент Г", "status": "Выполнен"},
                {"transaction_id": "TXN-005", "date": "19.01.2024", "amount": 4100, "customer": "Клиент Д", "status": "Отменен"},
                {"transaction_id": "TXN-006", "date": "20.01.2024", "amount": 6300, "customer": "Клиент Е", "status": "Выполнен"},
                {"transaction_id": "TXN-007", "date": "21.01.2024", "amount": 1900, "customer": "Клиент Ж", "status": "В обработке"},
                {"transaction_id": "TXN-008", "date": "22.01.2024", "amount": 8500, "customer": "Клиент З", "status": "Выполнен"}
            ]
        }
    ]
}


if __name__ == "__main__":
    # Создание отчета
    print("🚀 Создание отчета со сворачиваемыми строками...")
    renderer = AdvancedExcelRenderer()
    workbook = renderer.create_collapsible_report(data)
    renderer.save_report("my_report.xlsx")
    print("✅ Отчет создан: my_report.xlsx")
    print("\n📋 Возможности сворачивания:")
    print("  • Секция 'Детальный анализ по категориям' - свернута (группированные данные)")
    print("  • Группа 'Одежда' внутри анализа - свернута")
    print("  • Секция 'Подробные транзакции' - свернута")
    print("  • Используйте кнопки ▼/▶ для сворачивания/разворачивания")
    print("  • Используйте кнопки группировки (1, 2, 3) слева от строк")