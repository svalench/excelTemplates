#!/usr/bin/env python3
"""
Тестовый скрипт для демонстрации работы Excel шаблонизатора
"""

import json
import os
from datetime import datetime
from create_test_template import create_test_template
from main import generate_report

def create_test_data():
    """Создание тестовых данных для отчета"""
    
    # Основные данные
    test_data = {
        "report_title": "Отчет по сотрудникам компании",
        "report_date": datetime.now().strftime('%d.%m.%Y'),
        "summary": {
            "total_employees": 6,
            "avg_salary": 65000,
            "report_link": "https://company.com/reports/2023"
        }
    }
    
    # Данные по отделам (для таблицы departments)
    departments_data = [
        {"name": "Отдел продаж", "employee_count": 3, "avg_salary": 55000},
        {"name": "IT отдел", "employee_count": 2, "avg_salary": 80000},
        {"name": "Бухгалтерия", "employee_count": 1, "avg_salary": 50000}
    ]
    
    # Детальные данные по сотрудникам
    employees_data = [
        {"name": "Иван Петров", "position": "Менеджер", "salary": 50000, "department": "Отдел продаж"},
        {"name": "Ольга Сидорова", "position": "Аналитик", "salary": 65000, "department": "Отдел продаж"},
        {"name": "Петр Иванов", "position": "Стажер", "salary": 30000, "department": "Отдел продаж"},
        {"name": "Алексей Козлов", "position": "Разработчик", "salary": 90000, "department": "IT отдел"},
        {"name": "Мария Волкова", "position": "Тестировщик", "salary": 70000, "department": "IT отдел"},
        {"name": "Анна Смирнова", "position": "Бухгалтер", "salary": 50000, "department": "Бухгалтерия"}
    ]
    
    # Добавляем табличные данные
    test_data["departments"] = departments_data
    test_data["employees"] = employees_data
    
    return test_data

def test_excel_generator():
    """Основная функция тестирования"""
    print("🚀 Тестирование Excel шаблонизатора")
    print("=" * 50)
    
    try:
        # 1. Создаем тестовый шаблон
        print("1. Создание тестового шаблона...")
        template_path = create_test_template()
        print(f"✅ Шаблон создан: {template_path}")
        
        # 2. Подготавливаем тестовые данные
        print("\n2. Подготовка тестовых данных...")
        test_data = create_test_data()
        print(f"✅ Данные подготовлены: {len(test_data['employees'])} сотрудников, {len(test_data['departments'])} отделов")
        
        # 3. Генерируем отчет
        print("\n3. Генерация Excel отчета...")
        output_file = generate_report("test_template", test_data, "xlsx")
        print(f"✅ Отчет сгенерирован: {output_file}")
        
        # 4. Проверяем результат
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"✅ Файл создан успешно, размер: {file_size} байт")
            
            # Переименовываем для удобства
            final_name = f"generated_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            os.rename(output_file, final_name)
            print(f"📊 Итоговый файл: {final_name}")
            
            return final_name
        else:
            print("❌ Файл не был создан")
            return None
            
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
        return None

def test_with_json_string():
    """Тест с JSON строкой вместо словаря"""
    print("\n" + "=" * 50)
    print("🧪 Тест с JSON строкой")
    print("=" * 50)
    
    try:
        # Создаем данные как JSON строку
        test_data = create_test_data()
        json_string = json.dumps(test_data, ensure_ascii=False, indent=2)
        
        print("1. Генерация отчета из JSON строки...")
        output_file = generate_report("test_template", json_string, "xlsx")
        
        if os.path.exists(output_file):
            final_name = f"json_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            os.rename(output_file, final_name)
            print(f"✅ JSON отчет создан: {final_name}")
            return final_name
        else:
            print("❌ JSON отчет не создан")
            return None
            
    except Exception as e:
        print(f"❌ Ошибка JSON теста: {e}")
        return None

def cleanup_test_files():
    """Очистка тестовых файлов"""
    test_files = ['test_template.xlsx']
    for file in test_files:
        if os.path.exists(file):
            os.remove(file)
            print(f"🗑️ Удален тестовый файл: {file}")

if __name__ == '__main__':
    try:
        # Основной тест
        result1 = test_excel_generator()
        
        # Тест с JSON
        result2 = test_with_json_string()
        
        print("\n" + "=" * 50)
        print("📋 РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ")
        print("=" * 50)
        
        if result1:
            print(f"✅ Основной тест: {result1}")
        else:
            print("❌ Основной тест: ПРОВАЛЕН")
            
        if result2:
            print(f"✅ JSON тест: {result2}")
        else:
            print("❌ JSON тест: ПРОВАЛЕН")
            
        print("\n🎉 Тестирование завершено!")
        
        # Показываем инструкции
        print("\n📖 ИНСТРУКЦИИ:")
        print("1. Откройте сгенерированные .xlsx файлы в Excel или LibreOffice")
        print("2. Проверьте, что все данные корректно подставились")
        print("3. Убедитесь, что таблицы отформатированы правильно")
        
    except KeyboardInterrupt:
        print("\n⏹️ Тестирование прервано пользователем")
    finally:
        # Очистка (по желанию)
        # cleanup_test_files()
        pass 