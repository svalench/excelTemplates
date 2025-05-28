#!/usr/bin/env python3
"""
Расширенный тест функциональности работы с изображениями
Демонстрация всех возможностей: секции с изображениями, таблицы с изображениями, рисование
"""

from advanced_report_generator import AdvancedExcelRenderer
from datetime import datetime
import base64

def create_sample_base64_image():
    """Создание простого base64 изображения (1x1 пиксель)"""
    # Простое PNG изображение 1x1 пиксель красного цвета
    return "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg=="

def create_enhanced_test_data():
    """Создание расширенных тестовых данных с изображениями"""
    
    sample_base64 = create_sample_base64_image()
    
    data = {
        "title": "Расширенный тест работы с изображениями",
        "subtitle": f"Полная демонстрация функциональности | {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        "summary": {
            "total_images": 8,
            "image_types": 4,
            "drawing_types": 3,
            "success_rate": 95.0
        },
        "sections": [
            {
                "title": "Изображение из URL",
                "type": "image",
                "collapsed": False,
                "image_config": {
                    "type": "url",
                    "source": "https://via.placeholder.com/400x200/4472C4/FFFFFF?text=URL+IMAGE",
                    "width": 400,
                    "height": 200,
                    "anchor": "B5",
                    "description": "Изображение загружено из URL (placeholder)"
                }
            },
            {
                "title": "Изображение из Base64",
                "type": "image",
                "collapsed": False,
                "image_config": {
                    "type": "base64",
                    "source": sample_base64,
                    "width": 100,
                    "height": 100,
                    "anchor": "B15",
                    "description": "Изображение из base64 строки (1x1 пиксель)"
                }
            },
            {
                "title": "Таблица с изображениями в ячейках",
                "type": "table",
                "collapsed": False,
                "image_columns": ["logo", "qr_code"],  # Колонки с изображениями
                "data": [
                    {
                        "company": "Компания А",
                        "logo": {
                            "type": "url",
                            "source": "https://img.freepik.com/free-vector/indonesian-halal-logo-new-branding-2022_17005-1495.jpg",
                            "width": 80,
                            "height": 60
                        },
                        "revenue": 1500000,
                        "qr_code": {
                            "type": "url", 
                            "source": "https://via.placeholder.com/60x60/000000/FFFFFF?text=QR",
                            "width": 60,
                            "height": 60
                        }
                    },
                    {
                        "company": "Компания Б",
                        "logo": {
                            "type": "base64",
                            "source": sample_base64,
                            "width": 80,
                            "height": 60
                        },
                        "revenue": 2300000,
                        "qr_code": {
                            "type": "url",
                            "source": "https://via.placeholder.com/60x60/FF0000/FFFFFF?text=QR2",
                            "width": 60,
                            "height": 60
                        }
                    },
                    {
                        "company": "Компания В",
                        "logo": "https://via.placeholder.com/80x60/FFC000/000000?text=LOGO+C",  # Простой URL
                        "revenue": 1800000,
                        "qr_code": "https://via.placeholder.com/60x60/4472C4/FFFFFF?text=QR3"  # Простой URL
                    }
                ]
            },
            {
                "title": "Диаграмма с прямоугольниками",
                "type": "drawing",
                "collapsed": False,
                "drawing_config": {
                    "type": "diagram",
                    "style": "boxes",
                    "title": "Статистика по компаниям",
                    "width": 600,
                    "height": 300,
                    "data": [
                        {"label": "Компания А", "value": "1.5M", "color": "#70AD47"},
                        {"label": "Компания Б", "value": "2.3M", "color": "#4472C4"},
                        {"label": "Компания В", "value": "1.8M", "color": "#FFC000"}
                    ]
                }
            },
            {
                "title": "Инфографика с метриками",
                "type": "drawing",
                "collapsed": False,
                "drawing_config": {
                    "type": "infographic",
                    "title": "Ключевые показатели бизнеса",
                    "width": 800,
                    "height": 400,
                    "data": [
                        {
                            "type": "metric",
                            "label": "Общая выручка",
                            "value": "5.6M ₽",
                            "color": "#4472C4"
                        },
                        {
                            "type": "metric",
                            "label": "Количество клиентов",
                            "value": "342",
                            "color": "#70AD47"
                        },
                        {
                            "type": "progress",
                            "label": "Выполнение плана",
                            "progress": 87,
                            "color": "#70AD47"
                        },
                        {
                            "type": "progress",
                            "label": "Рост продаж",
                            "progress": 65,
                            "color": "#4472C4"
                        },
                        {
                            "type": "icon",
                            "symbol": "📈",
                            "text": "Положительная динамика",
                            "color": "#70AD47"
                        },
                        {
                            "type": "icon",
                            "symbol": "🎯",
                            "text": "Цели достигнуты",
                            "color": "#4472C4"
                        }
                    ]
                }
            },
            {
                "title": "Пользовательская схема архитектуры",
                "type": "drawing",
                "collapsed": False,
                "drawing_config": {
                    "type": "custom",
                    "width": 700,
                    "height": 350,
                    "commands": [
                        # База данных
                        {
                            "type": "rectangle",
                            "coords": [50, 150, 150, 220],
                            "color": "#4472C4"
                        },
                        {
                            "type": "text",
                            "position": [75, 175],
                            "text": "База данных",
                            "color": "white",
                            "size": 12
                        },
                        # API сервер
                        {
                            "type": "rectangle",
                            "coords": [250, 150, 350, 220],
                            "color": "#70AD47"
                        },
                        {
                            "type": "text",
                            "position": [285, 175],
                            "text": "API",
                            "color": "white",
                            "size": 12
                        },
                        # Веб-приложение
                        {
                            "type": "rectangle",
                            "coords": [450, 100, 550, 170],
                            "color": "#FFC000"
                        },
                        {
                            "type": "text",
                            "position": [470, 125],
                            "text": "Web App",
                            "color": "black",
                            "size": 12
                        },
                        # Мобильное приложение
                        {
                            "type": "rectangle",
                            "coords": [450, 200, 550, 270],
                            "color": "#C55A5A"
                        },
                        {
                            "type": "text",
                            "position": [470, 225],
                            "text": "Mobile App",
                            "color": "white",
                            "size": 12
                        },
                        # Соединения
                        {
                            "type": "line",
                            "coords": [150, 185, 250, 185],
                            "color": "black",
                            "width": 3
                        },
                        {
                            "type": "line",
                            "coords": [350, 185, 450, 135],
                            "color": "black",
                            "width": 3
                        },
                        {
                            "type": "line",
                            "coords": [350, 185, 450, 235],
                            "color": "black",
                            "width": 3
                        },
                        # Заголовок
                        {
                            "type": "text",
                            "position": [250, 50],
                            "text": "Архитектура системы",
                            "color": "black",
                            "size": 18
                        },
                        # Облако
                        {
                            "type": "circle",
                            "coords": [580, 150, 650, 220],
                            "color": "#E7E6E6"
                        },
                        {
                            "type": "text",
                            "position": [595, 175],
                            "text": "Cloud",
                            "color": "black",
                            "size": 10
                        }
                    ]
                }
            },
            {
                "title": "Таблица с обычными данными",
                "type": "table",
                "collapsed": True,
                "data": [
                    {"metric": "Загрузка изображений", "status": "✅ Успешно", "count": 6, "time": "2.3s"},
                    {"metric": "Создание диаграмм", "status": "✅ Успешно", "count": 3, "time": "1.8s"},
                    {"metric": "Рендеринг таблиц", "status": "✅ Успешно", "count": 2, "time": "0.5s"},
                    {"metric": "Общее время", "status": "✅ Завершено", "count": 11, "time": "4.6s"}
                ]
            }
        ]
    }
    
    return data

def main():
    """Основная функция тестирования расширенной функциональности"""
    print("🖼️  Тестирование расширенной функциональности работы с изображениями")
    print("=" * 80)
    
    try:
        # Создаем данные
        data = create_enhanced_test_data()
        
        # Создаем рендерер
        renderer = AdvancedExcelRenderer()
        
        # Генерируем отчет
        print("📊 Создание расширенного отчета с изображениями...")
        workbook = renderer.create_collapsible_report(data)
        
        # Сохраняем файл
        filename = "enhanced_images_report.xlsx"
        renderer.save_report(filename)
        
        print(f"✅ Расширенный отчет успешно создан: {filename}")
        
        print("\n🎨 Созданные элементы:")
        print("  • Изображение из URL")
        print("  • Изображение из Base64")
        print("  • Таблица с изображениями в ячейках")
        print("  • Диаграмма с прямоугольниками")
        print("  • Инфографика с метриками")
        print("  • Пользовательская схема архитектуры")
        print("  • Обычная таблица с данными")
        
        print("\n📋 Поддерживаемые типы изображений:")
        print("  • URL - загрузка из интернета")
        print("  • Base64 - встроенные изображения")
        print("  • File - локальные файлы")
        
        print("\n🔧 Поддерживаемые типы рисования:")
        print("  • diagram - диаграммы (boxes, circles, flow)")
        print("  • infographic - инфографика с метриками")
        print("  • custom - пользовательские рисунки")
        
    except Exception as e:
        print(f"❌ Ошибка при создании расширенного отчета: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main() 