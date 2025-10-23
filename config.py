# -*- coding: utf-8 -*-
"""
Конфигурационный файл для программы сравнения периодов Excel файлов
Содержит все настройки и параметры для работы программы
"""

import os
from pathlib import Path

# Базовые пути (кроссплатформенные)
BASE_DIR = Path(__file__).parent.absolute()
IN_XLSX_DIR = BASE_DIR / "IN_XLSX"
OUT_XLSX_DIR = BASE_DIR / "OUT_XLSX"
LOGS_DIR = BASE_DIR / "LOGS"

# Создание каталогов если они не существуют
IN_XLSX_DIR.mkdir(exist_ok=True)
OUT_XLSX_DIR.mkdir(exist_ok=True)
LOGS_DIR.mkdir(exist_ok=True)

# Настройки логирования
from datetime import datetime
log_timestamp = datetime.now().strftime("%Y%m%d_%H%M")
LOG_CONFIG = {
    'level': 'DEBUG',  # INFO или DEBUG
    'file': str(LOGS_DIR / f'comparison_{log_timestamp}.log'),
    'format': '%(asctime)s - %(levelname)s - %(message)s'
}

# Сообщения для логирования удалены - теперь используются прямые тексты в коде

# Параметры анализа
ANALYSIS_CONFIG = {
    # Количество файлов для анализа (2 или 3)
    'file_count': 3,
    
    # Режим расчета прироста
    'growth_calculation_mode': 'report_date',  # 'report_date' или 'deal_date'
    
    # Настройки файлов
    'files': [
        {
            'path': str(IN_XLSX_DIR / 'QS ФОТ (30-06-2025).xlsx'),
            'sheet_name': 'Sheet1',
            'columns': {
                'Таб. номер': 'tab_number',
                'КМ': 'fio',
                'ТБ': 'tb',
                'ГОСБ': 'gosb',
                'ИНН': 'client_id',
                'Клиент': 'client_name',
                'ФОТ': 'value'
            }
        },
        {
            'path': str(IN_XLSX_DIR / 'QS ФОТ (31-05-2025).xlsx'),
            'sheet_name': 'Sheet1',
            'columns': {
                'Таб. номер': 'tab_number',
                'КМ': 'fio',
                'ТБ': 'tb',
                'ГОСБ': 'gosb',
                'ИНН': 'client_id',
                'Клиент': 'client_name',
                'ФОТ': 'value'
            }
        },
        {
            'path': str(IN_XLSX_DIR / 'QS ФОТ (30-04-2025).xlsx'),
            'sheet_name': 'Sheet1',
            'columns': {
                'Таб. номер': 'tab_number',
                'КМ': 'fio',
                'ТБ': 'tb',
                'ГОСБ': 'gosb',
                'ИНН': 'client_id',
                'Клиент': 'client_name',
                'ФОТ': 'value'
            }
        }
    ],
    
    # Настройки выходного файла
    'output': {
        'file_name': str(OUT_XLSX_DIR / 'comparison_result'),
        'add_timestamp': True,  # Добавлять временную метку к имени файла
        'sheets': {
            'clients': 'Клиенты',
            'managers': 'Клиентские менеджеры',
            'managers_deal_date': 'КМ по дате сделки'
        },
        'columns': {
            'clients': [
                'ID client',
                'Client Name',
                'val (T-0)',
                'val (T-1)',
                'val (T-2)',
                'Gain',
                'TN (final)',
                'ФИО КМ (final)',
                'ГОСБ',
                'ТБ'
            ],
            'managers': [
                'TN (unic)',
                'ФИО',
                'ТБ',
                'ГОСБ',
                'val (T-0)',
                'val (T-1)',
                'val (T-2)',
                'Gain (total)'
            ],
            'managers_deal_date': [
                'TN (unic)',
                'ФИО',
                'ТБ',
                'ГОСБ',
                'val (T-0)',
                'val (T-1)',
                'val (T-2)',
                'Gain (total)'
            ]
        },
        # Форматирование колонок
        'formatting': {
            'clients': {
                'ID client': {'type': 'text_padded', 'format': '20', 'pad_char': '0'},
                'Client Name': {'type': 'text'},
                'val (T-0)': {'type': 'number', 'format': '#,##0.00'},
                'val (T-1)': {'type': 'number', 'format': '#,##0.00'},
                'val (T-2)': {'type': 'number', 'format': '#,##0.00'},
                'Gain': {'type': 'number', 'format': '#,##0.00'},
                'TN (final)': {'type': 'text_padded', 'format': '8', 'pad_char': '0'},
                'ФИО КМ (final)': {'type': 'text'},
                'ГОСБ': {'type': 'text'},
                'ТБ': {'type': 'text'}
            },
            'managers': {
                'TN (unic)': {'type': 'text_padded', 'format': '8', 'pad_char': '0'},
                'ФИО': {'type': 'text'},
                'ТБ': {'type': 'text'},
                'ГОСБ': {'type': 'text'},
                'val (T-0)': {'type': 'number', 'format': '#,##0.00'},
                'val (T-1)': {'type': 'number', 'format': '#,##0.00'},
                'val (T-2)': {'type': 'number', 'format': '#,##0.00'},
                'Gain (total)': {'type': 'number', 'format': '#,##0.00'}
            },
            'managers_deal_date': {
                'TN (unic)': {'type': 'text_padded', 'format': '8', 'pad_char': '0'},
                'ФИО': {'type': 'text'},
                'ТБ': {'type': 'text'},
                'ГОСБ': {'type': 'text'},
                'val (T-0)': {'type': 'number', 'format': '#,##0.00'},
                'val (T-1)': {'type': 'number', 'format': '#,##0.00'},
                'val (T-2)': {'type': 'number', 'format': '#,##0.00'},
                'Gain (total)': {'type': 'number', 'format': '#,##0.00'}
            }
        }
    }
}

# Настройки для создания тестовых данных
TEST_DATA_CONFIG = {
    'create_test_data': False,  # Создавать тестовые данные при запуске (True/False)
    'clients_count': 25000,
    'managers_count': 1000,
    'tb_count': 11,
    'gosb_per_tb': (5, 10),  # диапазон ГОСБ в каждом ТБ
    'value_range': (10000, 1000000),  # диапазон показателей
    'manager_change_rate': 0.15,  # 15% изменений в менеджерах между периодами
    'value_increase_rate': 1.2,  # увеличение показателей в следующих периодах
    
    # Названия тестовых файлов (отдельно от основных)
    'test_files': [
        'test_period1.xlsx',
        'test_period2.xlsx', 
        'test_period3.xlsx'
    ]
}

# Режимы работы программы
PROGRAM_MODES = {
    'mode': 4,  # 1-4 варианта работы
    
    # Варианты работы:
    # 1 - просто сгенерировать тест данные
    # 2 - посчитать на тест данных
    # 3 - посчитать на обычных данных
    # 4 - сгенерировать и посчитать на тест данных сразу
}
