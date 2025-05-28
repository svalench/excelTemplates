#!/usr/bin/env python3
"""
Тест функциональности работы с изображениями без зависимости от интернета
Создает локальные изображения и демонстрирует все возможности
"""

from advanced_report_generator import AdvancedExcelRenderer
from datetime import datetime
import base64
from PIL import Image as PILImage, ImageDraw, ImageFont
import os

def create_local_test_images():
    """Создание локальных тестовых изображений"""
    
    # Создаем папку для изображений
    if not os.path.exists('test_images'):
        os.makedirs('test_images')
    
    # 1. Простое цветное изображение (логотип)
    logo_img = PILImage.new('RGB', (200, 100), '#4472C4')
    draw = ImageDraw.Draw(logo_img)
    
    try:
        font = ImageFont.truetype("arial.ttf", 20)
    except:
        font = ImageFont.load_default()
    
    # Добавляем текст
    draw.text((50, 35), "LOGO", fill='white', font=font)
    logo_img.save('test_images/logo.png')
    
    # 2. QR-код имитация
    qr_img = PILImage.new('RGB', (100, 100), 'white')
    draw = ImageDraw.Draw(qr_img)
    
    # Рисуем простой паттерн как QR-код
    for i in range(0, 100, 10):
        for j in range(0, 100, 10):
            if (i + j) % 20 == 0:
                draw.rectangle([i, j, i+8, j+8], fill='black')
    
    qr_img.save('test_images/qr_code.png')
    
    # 3. График-изображение
    chart_img = PILImage.new('RGB', (300, 200), 'white')
    draw = ImageDraw.Draw(chart_img)
    
    # Рисуем простой столбчатый график
    bars = [50, 80, 120, 90, 110]
    bar_width = 40
    max_height = 150
    
    for i, height in enumerate(bars):
        x = 30 + i * 50
        y = 180 - height
        draw.rectangle([x, y, x + bar_width, 180], fill='#70AD47', outline='black')
        
        # Подписи
        try:
            small_font = ImageFont.truetype("arial.ttf", 12)
        except:
            small_font = ImageFont.load_default()
        
        draw.text((x + 15, 185), f'Q{i+1}', fill='black', font=small_font)
        draw.text((x + 10, y - 15), str(height), fill='black', font=small_font)
    
    # Заголовок
    draw.text((100, 10), "Продажи по кварталам", fill='black', font=font)
    chart_img.save('test_images/sales_chart.png')
    
    print("✅ Локальные тестовые изображения созданы в папке test_images/")

def create_base64_image():
    """Создание base64 изображения"""
    # Создаем простое изображение 50x50 пикселей
    img = PILImage.new('RGB', (50, 50), '#FFC000')
    draw = ImageDraw.Draw(img)
    
    # Рисуем круг
    draw.ellipse([10, 10, 40, 40], fill='#4472C4', outline='black', width=2)
    
    # Сохраняем в base64
    from io import BytesIO
    buffer = BytesIO()
    img.save(buffer, format='PNG')
    img_str = base64.b64encode(buffer.getvalue()).decode()
    
    return img_str

def create_local_test_data():
    """Создание тестовых данных с локальными изображениями"""
    
    base64_img = create_base64_image()
    
    data = {
        "title": "Тест с локальными изображениями",
        "subtitle": f"Демонстрация без интернета | {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        "summary": {
            "local_images": 3,
            "base64_images": 1,
            "drawings": 4,
            "total_elements": 8
        },
        "sections": [
            {
                "title": "Логотип из локального файла",
                "type": "image",
                "collapsed": False,
                "image_config": {
                    "type": "file",
                    "source": "test_images/logo.png",
                    "width": 200,
                    "height": 100,
                    "anchor": "B5",
                    "description": "Логотип компании (локальный файл)"
                }
            },
            {
                "title": "Изображение из Base64",
                "type": "image",
                "collapsed": False,
                "image_config": {
                    "type": "base64",
                    "source": base64_img,
                    "width": 100,
                    "height": 100,
                    "anchor": "B12",
                    "description": "Круг из base64 данных"
                }
            },
            {
                "title": "Таблица компаний с локальными изображениями",
                "type": "table",
                "collapsed": False,
                "image_columns": ["logo", "qr_code"],
                "data": [
                    {
                        "company": "ООО Альфа",
                        "logo": {
                            "type": "file",
                            "source": "test_images/logo.png",
                            "width": 60,
                            "height": 30
                        },
                        "revenue": 2500000,
                        "qr_code": {
                            "type": "file",
                            "source": "test_images/qr_code.png",
                            "width": 40,
                            "height": 40
                        }
                    },
                    {
                        "company": "ООО Бета",
                        "logo": {
                            "type": "base64",
                            "source": base64_img,
                            "width": 60,
                            "height": 30
                        },
                        "revenue": 1800000,
                        "qr_code": "test_images/qr_code.png"  # Простой путь к файлу
                    },
                    {
                        "company": "ООО Гамма",
                        "logo": "test_images/logo.png",  # Простой путь к файлу
                        "revenue": 3200000,
                        "qr_code": "test_images/qr_code.png"
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
                    "title": "Выручка по компаниям",
                    "width": 500,
                    "height": 250,
                    "data": [
                        {"label": "Альфа", "value": "2.5M", "color": "#4472C4"},
                        {"label": "Бета", "value": "1.8M", "color": "#70AD47"},
                        {"label": "Гамма", "value": "3.2M", "color": "#FFC000"}
                    ]
                }
            },
            {
                "title": "Инфографика KPI",
                "type": "drawing",
                "collapsed": False,
                "drawing_config": {
                    "type": "infographic",
                    "title": "Ключевые показатели",
                    "width": 700,
                    "height": 350,
                    "data": [
                        {
                            "type": "metric",
                            "label": "Общая выручка",
                            "value": "7.5M ₽",
                            "color": "#4472C4"
                        },
                        {
                            "type": "metric",
                            "label": "Количество компаний",
                            "value": "3",
                            "color": "#70AD47"
                        },
                        {
                            "type": "progress",
                            "label": "Выполнение плана",
                            "progress": 92,
                            "color": "#70AD47"
                        },
                        {
                            "type": "progress",
                            "label": "Рост к прошлому году",
                            "progress": 78,
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
                            "text": "План перевыполнен",
                            "color": "#4472C4"
                        }
                    ]
                }
            },
            {
                "title": "Схема бизнес-процесса",
                "type": "drawing",
                "collapsed": False,
                "drawing_config": {
                    "type": "custom",
                    "width": 600,
                    "height": 300,
                    "commands": [
                        # Заголовок
                        {
                            "type": "text",
                            "position": [200, 20],
                            "text": "Процесс продаж",
                            "color": "black",
                            "size": 16
                        },
                        # Этап 1: Лид
                        {
                            "type": "circle",
                            "coords": [50, 100, 120, 170],
                            "color": "#4472C4"
                        },
                        {
                            "type": "text",
                            "position": [70, 125],
                            "text": "Лид",
                            "color": "white",
                            "size": 12
                        },
                        # Стрелка 1
                        {
                            "type": "line",
                            "coords": [120, 135, 180, 135],
                            "color": "black",
                            "width": 3
                        },
                        # Этап 2: Переговоры
                        {
                            "type": "rectangle",
                            "coords": [180, 110, 280, 160],
                            "color": "#70AD47"
                        },
                        {
                            "type": "text",
                            "position": [200, 125],
                            "text": "Переговоры",
                            "color": "white",
                            "size": 10
                        },
                        # Стрелка 2
                        {
                            "type": "line",
                            "coords": [280, 135, 340, 135],
                            "color": "black",
                            "width": 3
                        },
                        # Этап 3: Сделка
                        {
                            "type": "circle",
                            "coords": [340, 100, 410, 170],
                            "color": "#FFC000"
                        },
                        {
                            "type": "text",
                            "position": [355, 125],
                            "text": "Сделка",
                            "color": "black",
                            "size": 12
                        },
                        # Стрелка 3
                        {
                            "type": "line",
                            "coords": [410, 135, 470, 135],
                            "color": "black",
                            "width": 3
                        },
                        # Этап 4: Оплата
                        {
                            "type": "rectangle",
                            "coords": [470, 110, 550, 160],
                            "color": "#C55A5A"
                        },
                        {
                            "type": "text",
                            "position": [490, 125],
                            "text": "Оплата",
                            "color": "white",
                            "size": 12
                        },
                        # Подписи времени
                        {
                            "type": "text",
                            "position": [70, 180],
                            "text": "День 1",
                            "color": "gray",
                            "size": 10
                        },
                        {
                            "type": "text",
                            "position": [210, 180],
                            "text": "День 2-5",
                            "color": "gray",
                            "size": 10
                        },
                        {
                            "type": "text",
                            "position": [360, 180],
                            "text": "День 6",
                            "color": "gray",
                            "size": 10
                        },
                        {
                            "type": "text",
                            "position": [490, 180],
                            "text": "День 7-10",
                            "color": "gray",
                            "size": 10
                        }
                    ]
                }
            },
            {
                "title": "Статистика по тестированию",
                "type": "table",
                "collapsed": True,
                "data": [
                    {"элемент": "Локальные изображения", "статус": "✅ Успешно", "количество": 3},
                    {"элемент": "Base64 изображения", "статус": "✅ Успешно", "количество": 1},
                    {"элемент": "Таблицы с изображениями", "статус": "✅ Успешно", "количество": 1},
                    {"элемент": "Диаграммы", "статус": "✅ Успешно", "количество": 1},
                    {"элемент": "Инфографика", "статус": "✅ Успешно", "количество": 1},
                    {"элемент": "Пользовательские схемы", "статус": "✅ Успешно", "количество": 1}
                ]
            }
        ]
    }
    
    return data

def main():
    """Основная функция тестирования с локальными изображениями"""
    print("🖼️  Тестирование функциональности с локальными изображениями")
    print("=" * 70)
    
    try:
        # Создаем локальные изображения
        print("🎨 Создание локальных тестовых изображений...")
        create_local_test_images()
        
        # Создаем данные
        print("📊 Подготовка данных для отчета...")
        data = create_local_test_data()
        
        # Создаем рендерер
        renderer = AdvancedExcelRenderer()
        
        # Генерируем отчет
        print("📈 Создание отчета с изображениями...")
        workbook = renderer.create_collapsible_report(data)
        
        # Сохраняем файл
        filename = "local_images_report.xlsx"
        renderer.save_report(filename)
        
        print(f"✅ Отчет с локальными изображениями создан: {filename}")
        
        print("\n🎨 Созданные элементы:")
        print("  • Изображение из локального файла")
        print("  • Изображение из Base64")
        print("  • Таблица с изображениями в ячейках")
        print("  • Диаграмма с прямоугольниками")
        print("  • Инфографика с метриками")
        print("  • Схема бизнес-процесса")
        print("  • Статистика тестирования")
        
        print("\n📁 Созданные файлы:")
        print("  • test_images/logo.png")
        print("  • test_images/qr_code.png")
        print("  • test_images/sales_chart.png")
        print(f"  • {filename}")
        
    except Exception as e:
        print(f"❌ Ошибка при создании отчета: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main() 