#!/usr/bin/env python3
"""
Расширенный генератор отчетов с поддержкой:
- Сворачиваемых колонок и строк
- Автофильтров
- Условного форматирования
- Многоуровневой группировки
- Интерактивных элементов
"""

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, NamedStyle
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from datetime import datetime, timedelta
import json
from jinja2 import Template
import tempfile
import os


class AdvancedExcelRenderer:
    """Расширенный рендерер Excel с продвинутыми возможностями"""
    
    def __init__(self):
        self.wb = None
        self.ws = None
        self.styles_created = False
        
    def create_styles(self):
        """Создание именованных стилей"""
        if self.styles_created:
            return
            
        # Стиль заголовка
        header_style = NamedStyle(name="header_style")
        header_style.font = Font(bold=True, size=14, color="FFFFFF")
        header_style.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_style.alignment = Alignment(horizontal="center", vertical="center")
        header_style.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Стиль подзаголовка
        subheader_style = NamedStyle(name="subheader_style")
        subheader_style.font = Font(bold=True, size=12, color="000000")
        subheader_style.fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
        subheader_style.alignment = Alignment(horizontal="left", vertical="center")
        
        # Стиль данных
        data_style = NamedStyle(name="data_style")
        data_style.font = Font(size=10)
        data_style.alignment = Alignment(horizontal="left", vertical="center")
        data_style.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Стиль числовых данных
        number_style = NamedStyle(name="number_style")
        number_style.font = Font(size=10)
        number_style.alignment = Alignment(horizontal="right", vertical="center")
        number_style.number_format = '#,##0.00'
        number_style.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Добавляем стили в книгу
        try:
            self.wb.add_named_style(header_style)
            self.wb.add_named_style(subheader_style)
            self.wb.add_named_style(data_style)
            self.wb.add_named_style(number_style)
        except ValueError:
            # Стили уже существуют
            pass
            
        self.styles_created = True
    
    def create_collapsible_report(self, data, template_config=None):
        """
        Создание отчета со сворачиваемыми секциями
        
        Args:
            data: Данные для отчета
            template_config: Конфигурация шаблона
        """
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = "Сложный отчет"
        
        self.create_styles()
        
        current_row = 1
        
        # Заголовок отчета
        current_row = self._add_report_header(data, current_row)
        
        # Основные метрики (всегда видимые)
        if 'summary' in data:
            current_row = self._add_summary_section(data['summary'], current_row)
        
        # Детальные секции (сворачиваемые)
        if 'sections' in data:
            for section in data['sections']:
                current_row = self._add_collapsible_section(section, current_row)
        
        # Добавляем автофильтры и форматирование
        self._apply_advanced_formatting()
        
        return self.wb
    
    def _add_report_header(self, data, start_row):
        """Добавление заголовка отчета"""
        title = data.get('title', 'Сложный отчет')
        subtitle = data.get('subtitle', f"Создан: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
        
        # Основной заголовок
        self.ws.cell(row=start_row, column=1, value=title)
        self.ws.cell(row=start_row, column=1).style = "header_style"
        self.ws.merge_cells(f'A{start_row}:F{start_row}')
        
        # Подзаголовок
        self.ws.cell(row=start_row + 1, column=1, value=subtitle)
        self.ws.cell(row=start_row + 1, column=1).style = "subheader_style"
        self.ws.merge_cells(f'A{start_row + 1}:F{start_row + 1}')
        
        return start_row + 3
    
    def _add_summary_section(self, summary_data, start_row):
        """Добавление секции сводки"""
        # Заголовок секции
        self.ws.cell(row=start_row, column=1, value="📊 ОСНОВНЫЕ ПОКАЗАТЕЛИ")
        self.ws.cell(row=start_row, column=1).style = "subheader_style"
        self.ws.merge_cells(f'A{start_row}:F{start_row}')
        
        current_row = start_row + 1
        
        # Метрики в виде карточек
        col = 1
        for key, value in summary_data.items():
            if col > 6:  # Переход на новую строку
                current_row += 1
                col = 1
            
            # Название метрики
            self.ws.cell(row=current_row, column=col, value=key.replace('_', ' ').title())
            self.ws.cell(row=current_row, column=col).font = Font(bold=True, size=9)
            
            # Значение метрики
            self.ws.cell(row=current_row + 1, column=col, value=value)
            self.ws.cell(row=current_row + 1, column=col).style = "number_style"
            
            col += 1
        
        return current_row + 3
    
    def _add_collapsible_section(self, section_data, start_row):
        """Добавление сворачиваемой секции"""
        section_title = section_data.get('title', 'Секция')
        section_type = section_data.get('type', 'table')
        is_collapsed = section_data.get('collapsed', False)
        
        # Заголовок секции с кнопкой сворачивания
        collapse_symbol = "▼" if not is_collapsed else "▶"
        header_text = f"{collapse_symbol} {section_title}"
        
        self.ws.cell(row=start_row, column=1, value=header_text)
        self.ws.cell(row=start_row, column=1).style = "subheader_style"
        self.ws.merge_cells(f'A{start_row}:F{start_row}')
        
        section_start_row = start_row + 1
        
        if section_type == 'table':
            current_row = self._add_table_with_filters(section_data, section_start_row)
        elif section_type == 'grouped_data':
            current_row = self._add_grouped_data(section_data, section_start_row)
        elif section_type == 'chart':
            current_row = self._add_chart_section(section_data, section_start_row)
        else:
            current_row = section_start_row
        
        # Настройка группировки для сворачивания
        if current_row > section_start_row:
            self._setup_row_grouping(section_start_row, current_row - 1, hidden=is_collapsed)
        
        return current_row + 1
    
    def _add_table_with_filters(self, section_data, start_row):
        """Добавление таблицы с автофильтрами"""
        data = section_data.get('data', [])
        if not data:
            return start_row
        
        # Проверяем, что data - это список словарей
        if not isinstance(data, list):
            print(f"Предупреждение: неверный формат данных в секции '{section_data.get('title', 'Unknown')}' - ожидается список")
            return start_row
            
        if not data or not isinstance(data[0], dict):
            print(f"Предупреждение: пустые данные или неверный формат в секции '{section_data.get('title', 'Unknown')}'")
            return start_row

        df = pd.DataFrame(data)
        
        # Заголовки таблицы
        for col_idx, column in enumerate(df.columns, 1):
            cell = self.ws.cell(row=start_row, column=col_idx, value=column)
            cell.style = "header_style"
        
        # Данные таблицы
        for row_idx, row_data in enumerate(dataframe_to_rows(df, index=False, header=False), start_row + 1):
            for col_idx, value in enumerate(row_data, 1):
                cell = self.ws.cell(row=row_idx, column=col_idx, value=value)
                if isinstance(value, (int, float)):
                    cell.style = "number_style"
                else:
                    cell.style = "data_style"
        
        # Создание таблицы Excel с автофильтрами
        table_range = f"A{start_row}:{chr(64 + len(df.columns))}{start_row + len(df)}"
        table = Table(displayName=f"Table{start_row}", ref=table_range)
        
        # Стиль таблицы
        style = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        table.tableStyleInfo = style
        
        self.ws.add_table(table)
        
        # Условное форматирование для числовых колонок
        self._apply_conditional_formatting(df, start_row)
        
        return start_row + len(df) + 1
    
    def _add_grouped_data(self, section_data, start_row):
        """Добавление группированных данных с многоуровневым сворачиванием"""
        groups = section_data.get('groups', [])
        
        if not isinstance(groups, list):
            print(f"Предупреждение: неверный формат групп в секции '{section_data.get('title', 'Unknown')}' - ожидается список")
            return start_row
        
        current_row = start_row
        
        for group in groups:
            if not isinstance(group, dict):
                print(f"Предупреждение: неверный формат группы в секции '{section_data.get('title', 'Unknown')}'")
                continue
                
            group_title = group.get('title', 'Группа')
            group_data = group.get('data', [])
            is_collapsed = group.get('collapsed', False)
            
            # Заголовок группы (уровень 1)
            collapse_symbol = "▼" if not is_collapsed else "▶"
            self.ws.cell(row=current_row, column=1, value=f"  {collapse_symbol} {group_title}")
            self.ws.cell(row=current_row, column=1).font = Font(bold=True, size=11)
            
            group_start_row = current_row + 1
            
            # Проверяем данные группы
            if not isinstance(group_data, list) or not group_data:
                current_row = group_start_row
                continue
                
            if not isinstance(group_data[0], dict):
                print(f"Предупреждение: неверный формат данных группы '{group_title}'")
                current_row = group_start_row
                continue
            
            # Данные группы
            df = pd.DataFrame(group_data)
            
            # Заголовки
            for col_idx, column in enumerate(df.columns, 2):  # Смещение для отступа
                cell = self.ws.cell(row=group_start_row, column=col_idx, value=column)
                cell.font = Font(bold=True, size=9)
                cell.fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
            
            # Данные
            for row_idx, row_data in enumerate(dataframe_to_rows(df, index=False, header=False), group_start_row + 1):
                for col_idx, value in enumerate(row_data, 2):
                    cell = self.ws.cell(row=row_idx, column=col_idx, value=value)
                    if isinstance(value, (int, float)):
                        cell.style = "number_style"
                    else:
                        cell.style = "data_style"
            
            current_row = group_start_row + len(df) + 1
            
            # Группировка строк группы (уровень 2)
            if current_row > group_start_row:
                self._setup_row_grouping(group_start_row, current_row - 1, level=2, hidden=is_collapsed)
            
            current_row += 1
        
        return current_row
    
    def _add_chart_section(self, section_data, start_row):
        """Добавление секции с графиком"""
        chart_type = section_data.get('chart_type', 'bar')
        data = section_data.get('data', [])
        
        # Проверяем данные
        if not isinstance(data, list) or not data:
            print(f"Предупреждение: пустые данные или неверный формат в графике '{section_data.get('title', 'Unknown')}'")
            return start_row
            
        if not isinstance(data[0], dict):
            print(f"Предупреждение: неверный формат данных графика в секции '{section_data.get('title', 'Unknown')}'")
            return start_row
        
        df = pd.DataFrame(data)
        
        # Добавляем данные для графика
        for row_idx, row_data in enumerate(dataframe_to_rows(df, index=False, header=True), start_row):
            for col_idx, value in enumerate(row_data, 1):
                self.ws.cell(row=row_idx, column=col_idx, value=value)
        
        # Создаем график
        if chart_type == 'bar':
            chart = BarChart()
        elif chart_type == 'line':
            chart = LineChart()
        elif chart_type == 'pie':
            chart = PieChart()
        else:
            chart = BarChart()
        
        # Настройка данных графика
        data_range = Reference(self.ws, min_col=2, min_row=start_row + 1, 
                              max_col=len(df.columns), max_row=start_row + len(df))
        categories = Reference(self.ws, min_col=1, min_row=start_row + 1, 
                              max_row=start_row + len(df))
        
        chart.add_data(data_range, titles_from_data=True)
        chart.set_categories(categories)
        
        # Размещение графика
        chart.width = 15
        chart.height = 10
        self.ws.add_chart(chart, f"H{start_row}")
        
        return start_row + len(df) + 15  # Учитываем высоту графика
    
    def _setup_row_grouping(self, start_row, end_row, level=1, hidden=False):
        """Настройка группировки строк"""
        for row in range(start_row, end_row + 1):
            self.ws.row_dimensions[row].outline_level = level
            if hidden:
                self.ws.row_dimensions[row].hidden = True
    
    def _apply_conditional_formatting(self, df, start_row):
        """Применение условного форматирования"""
        for col_idx, column in enumerate(df.columns, 1):
            if df[column].dtype in ['int64', 'float64']:
                # Цветовая шкала для числовых данных
                range_str = f"{chr(64 + col_idx)}{start_row + 1}:{chr(64 + col_idx)}{start_row + len(df)}"
                rule = ColorScaleRule(
                    start_type='min', start_color='F8696B',
                    mid_type='percentile', mid_value=50, mid_color='FFEB9C',
                    end_type='max', end_color='63BE7B'
                )
                self.ws.conditional_formatting.add(range_str, rule)
    
    def _apply_advanced_formatting(self):
        """Применение продвинутого форматирования"""
        # Автоподбор ширины колонок
        for column in self.ws.columns:
            max_length = 0
            column_letter = None
            
            for cell in column:
                # Пропускаем объединенные ячейки
                if hasattr(cell, 'column_letter'):
                    column_letter = cell.column_letter
                    try:
                        if cell.value and len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
            
            if column_letter:
                adjusted_width = min(max_length + 2, 50)
                self.ws.column_dimensions[column_letter].width = adjusted_width
        
        # Закрепление области
        self.ws.freeze_panes = 'A4'
    
    def save_report(self, filename):
        """Сохранение отчета"""
        self.wb.save(filename)
        return filename


def create_complex_report_template():
    """Создание шаблона для сложного отчета"""
    template_data = {
        "title": "{{report_title}}",
        "subtitle": "Период: {{period_start}} - {{period_end}} | Создан: {{creation_date}}",
        "summary": {
            "total_revenue": "{{summary.total_revenue}}",
            "total_orders": "{{summary.total_orders}}",
            "avg_order_value": "{{summary.avg_order_value}}",
            "growth_rate": "{{summary.growth_rate}}",
            "profit_margin": "{{summary.profit_margin}}",
            "customer_count": "{{summary.customer_count}}"
        },
        "sections": [
            {
                "title": "Продажи по регионам",
                "type": "table",
                "collapsed": False,
                "data": "{{regional_sales}}"
            },
            {
                "title": "Анализ продуктов",
                "type": "grouped_data",
                "collapsed": True,
                "groups": "{{product_groups}}"
            },
            {
                "title": "Динамика продаж",
                "type": "chart",
                "chart_type": "line",
                "collapsed": False,
                "data": "{{sales_dynamics}}"
            },
            {
                "title": "Детальная аналитика",
                "type": "table",
                "collapsed": True,
                "data": "{{detailed_analytics}}"
            }
        ]
    }
    
    # Сохраняем конфигурацию шаблона
    with open('complex_report_template.json', 'w', encoding='utf-8') as f:
        json.dump(template_data, f, ensure_ascii=False, indent=2)
    
    return template_data


def generate_sample_data():
    """Генерация примерных данных для демонстрации"""
    import random
    
    # Основные метрики
    summary = {
        "total_revenue": 2500000,
        "total_orders": 1250,
        "avg_order_value": 2000,
        "growth_rate": 15.5,
        "profit_margin": 28.5,
        "customer_count": 850
    }
    
    # Продажи по регионам
    regional_sales = [
        {"region": "Москва", "sales": 850000, "orders": 425, "growth": 18.2},
        {"region": "СПб", "sales": 620000, "orders": 310, "growth": 12.5},
        {"region": "Екатеринбург", "sales": 380000, "orders": 190, "growth": 8.7},
        {"region": "Новосибирск", "sales": 320000, "orders": 160, "growth": 22.1},
        {"region": "Казань", "sales": 280000, "orders": 140, "growth": 5.3},
        {"region": "Нижний Новгород", "sales": 50000, "orders": 25, "growth": -2.1}
    ]
    
    # Группированные данные по продуктам
    product_groups = [
        {
            "title": "Электроника",
            "collapsed": False,
            "data": [
                {"product": "Смартфоны", "sales": 450000, "units": 150, "margin": 25.5},
                {"product": "Ноутбуки", "sales": 380000, "units": 95, "margin": 18.2},
                {"product": "Планшеты", "sales": 220000, "units": 110, "margin": 22.1}
            ]
        },
        {
            "title": "Одежда",
            "collapsed": True,
            "data": [
                {"product": "Куртки", "sales": 180000, "units": 360, "margin": 45.2},
                {"product": "Джинсы", "sales": 150000, "units": 300, "margin": 38.5},
                {"product": "Футболки", "sales": 120000, "units": 600, "margin": 52.1}
            ]
        },
        {
            "title": "Книги",
            "collapsed": True,
            "data": [
                {"product": "Художественная литература", "sales": 85000, "units": 850, "margin": 35.2},
                {"product": "Техническая литература", "sales": 65000, "units": 325, "margin": 28.5},
                {"product": "Детские книги", "sales": 45000, "units": 450, "margin": 42.1}
            ]
        }
    ]
    
    # Динамика продаж
    sales_dynamics = []
    base_date = datetime(2024, 1, 1)
    for i in range(12):
        month_date = base_date + timedelta(days=30*i)
        sales_dynamics.append({
            "month": month_date.strftime("%B"),
            "sales": random.randint(180000, 250000),
            "orders": random.randint(90, 130),
            "customers": random.randint(60, 90)
        })
    
    # Детальная аналитика
    detailed_analytics = []
    for i in range(50):
        detailed_analytics.append({
            "order_id": f"ORD-{1000+i}",
            "customer": f"Клиент {i+1}",
            "product": random.choice(["Смартфон", "Ноутбук", "Куртка", "Книга"]),
            "amount": random.randint(500, 5000),
            "date": (datetime.now() - timedelta(days=random.randint(1, 90))).strftime("%d.%m.%Y"),
            "status": random.choice(["Выполнен", "В обработке", "Отменен"])
        })
    
    return {
        "report_title": "Комплексный отчет по продажам",
        "period_start": "01.01.2024",
        "period_end": "31.12.2024",
        "creation_date": datetime.now().strftime("%d.%m.%Y %H:%M"),
        "summary": summary,
        "regional_sales": regional_sales,
        "product_groups": product_groups,
        "sales_dynamics": sales_dynamics,
        "detailed_analytics": detailed_analytics
    }


def render_template_with_data(template_data, context_data):
    """Рендеринг шаблона с данными используя Jinja2"""
    
    def render_recursive(obj, context):
        """Рекурсивный рендеринг объекта"""
        if isinstance(obj, str):
            if "{{" in obj or "{%" in obj:
                try:
                    template = Template(obj)
                    return template.render(**context)
                except Exception as e:
                    print(f"Ошибка рендеринга шаблона: {e}")
                    return obj
            return obj
        elif isinstance(obj, dict):
            return {key: render_recursive(value, context) for key, value in obj.items()}
        elif isinstance(obj, list):
            return [render_recursive(item, context) for item in obj]
        else:
            return obj
    
    # Сначала обрабатываем простые подстановки
    rendered = render_recursive(template_data, context_data)
    
    # Затем обрабатываем прямые ссылки на данные
    def resolve_data_references(obj, context):
        """Разрешение прямых ссылок на данные"""
        if isinstance(obj, str):
            # Проверяем, является ли строка ссылкой на данные
            if obj.startswith("{{") and obj.endswith("}}"):
                key = obj[2:-2].strip()
                # Поддержка вложенных ключей (например, summary.total_revenue)
                try:
                    value = context
                    for part in key.split('.'):
                        value = value[part]
                    return value
                except (KeyError, TypeError):
                    print(f"Не найден ключ: {key}")
                    return obj
            return obj
        elif isinstance(obj, dict):
            return {key: resolve_data_references(value, context) for key, value in obj.items()}
        elif isinstance(obj, list):
            return [resolve_data_references(item, context) for item in obj]
        else:
            return obj
    
    return resolve_data_references(rendered, context_data)


if __name__ == "__main__":
    # Демонстрация использования
    print("🚀 Создание сложного отчета со сворачиваемыми секциями...")
    
    # Создаем шаблон
    template = create_complex_report_template()
    print("✅ Шаблон создан: complex_report_template.json")
    
    # Генерируем данные
    sample_data = generate_sample_data()
    
    # Рендерим шаблон с данными
    rendered_template = render_template_with_data(template, sample_data)
    
    # Создаем отчет
    renderer = AdvancedExcelRenderer()
    workbook = renderer.create_collapsible_report(rendered_template)
    
    # Сохраняем
    output_file = f"complex_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    renderer.save_report(output_file)
    
    print(f"✅ Сложный отчет создан: {output_file}")
    print("\n📋 Возможности отчета:")
    print("  • Сворачиваемые секции с кнопками ▼/▶")
    print("  • Автофильтры в таблицах")
    print("  • Условное форматирование")
    print("  • Многоуровневая группировка")
    print("  • Интерактивные графики")
    print("  • Закрепленные области")
    print("  • Автоподбор ширины колонок") 