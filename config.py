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
LOG_CONFIG = {
    'level': 'DEBUG',  # INFO или DEBUG
    'file': str(LOGS_DIR / 'comparison.log'),
    'format': '%(asctime)s - %(levelname)s - %(message)s'
}

# Сообщения для логирования
LOG_MESSAGES = {
    'program_start': 'Запуск программы сравнения периодов',
    'program_end': 'Программа завершена успешно',
    'file_loading': 'Загрузка файла: {}',
    'file_loaded': 'Файл загружен успешно: {}',
    'data_processing': 'Обработка данных из файла: {}',
    'calculation_start': 'Начало расчета приростов',
    'calculation_end': 'Расчет приростов завершен',
    'output_creation': 'Создание выходного файла: {}',
    'output_created': 'Выходной файл создан: {}',
    'error_occurred': 'Произошла ошибка: {}',
    'config_loaded': 'Конфигурация загружена',
    'test_data_creation': 'Создание тестовых данных',
    'test_data_created': 'Тестовые данные созданы'
}

# Параметры анализа
ANALYSIS_CONFIG = {
    # Количество файлов для анализа (2 или 3)
    'file_count': 3,
    
    # Режим расчета прироста
    'growth_calculation_mode': 'report_date',  # 'report_date' или 'deal_date'
    
    # Настройки файлов
    'files': [
        {
            'path': str(IN_XLSX_DIR / 'test_data_period1.xlsx'),
            'sheet_name': 'Данные',
            'columns': {
                'Таб. номер': 'tab_number',
                'Фамилия': 'fio',
                'коротко ТБ': 'tb',
                'короткое наименование ГОСБ': 'gosb',
                'ЕПК ИД': 'client_id',
                'Наименование клиента': 'client_name',
                'СДО руб': 'value'
            }
        },
        {
            'path': str(IN_XLSX_DIR / 'test_data_period2.xlsx'),
            'sheet_name': 'Данные',
            'columns': {
                'Таб. номер': 'tab_number',
                'Фамилия': 'fio',
                'коротко ТБ': 'tb',
                'короткое наименование ГОСБ': 'gosb',
                'ЕПК ИД': 'client_id',
                'Наименование клиента': 'client_name',
                'СДО руб': 'value'
            }
        },
        {
            'path': str(IN_XLSX_DIR / 'test_data_period3.xlsx'),
            'sheet_name': 'Данные',
            'columns': {
                'Таб. номер': 'tab_number',
                'Фамилия': 'fio',
                'коротко ТБ': 'tb',
                'короткое наименование ГОСБ': 'gosb',
                'ЕПК ИД': 'client_id',
                'Наименование клиента': 'client_name',
                'СДО руб': 'value'
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
                'Идентификатор клиента',
                'Наименование клиента',
                'Показатель 1',
                'Показатель 2',
                'Показатель 3',
                'Прирост',
                'Табельный итоговый',
                'ФИО итогового сотрудника',
                'ГОСБ',
                'ТБ'
            ],
            'managers': [
                'Табельный уникальный',
                'ФИО',
                'ТБ',
                'ГОСБ',
                'Показатель 1',
                'Показатель 2',
                'Показатель 3',
                'Суммарный прирост'
            ],
            'managers_deal_date': [
                'Табельный уникальный',
                'ФИО',
                'ТБ',
                'ГОСБ',
                'Показатель 1',
                'Показатель 2',
                'Показатель 3',
                'Суммарный прирост'
            ]
        },
        # Форматирование колонок
        'formatting': {
            'clients': {
                'Идентификатор клиента': {'type': 'text_padded', 'format': '20', 'pad_char': '0'},
                'Наименование клиента': {'type': 'text'},
                'Показатель 1': {'type': 'number', 'format': '#,##0.00'},
                'Показатель 2': {'type': 'number', 'format': '#,##0.00'},
                'Показатель 3': {'type': 'number', 'format': '#,##0.00'},
                'Прирост': {'type': 'number', 'format': '#,##0.00'},
                'Табельный итоговый': {'type': 'text_padded', 'format': '8', 'pad_char': '0'},
                'ФИО итогового сотрудника': {'type': 'text'},
                'ГОСБ': {'type': 'text'},
                'ТБ': {'type': 'text'}
            },
            'managers': {
                'Табельный уникальный': {'type': 'text_padded', 'format': '8', 'pad_char': '0'},
                'ФИО': {'type': 'text'},
                'ТБ': {'type': 'text'},
                'ГОСБ': {'type': 'text'},
                'Показатель 1': {'type': 'number', 'format': '#,##0.00'},
                'Показатель 2': {'type': 'number', 'format': '#,##0.00'},
                'Показатель 3': {'type': 'number', 'format': '#,##0.00'},
                'Суммарный прирост': {'type': 'number', 'format': '#,##0.00'}
            },
            'managers_deal_date': {
                'Табельный уникальный': {'type': 'text_padded', 'format': '8', 'pad_char': '0'},
                'ФИО': {'type': 'text'},
                'ТБ': {'type': 'text'},
                'ГОСБ': {'type': 'text'},
                'Показатель 1': {'type': 'number', 'format': '#,##0.00'},
                'Показатель 2': {'type': 'number', 'format': '#,##0.00'},
                'Показатель 3': {'type': 'number', 'format': '#,##0.00'},
                'Суммарный прирост': {'type': 'number', 'format': '#,##0.00'}
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
    'value_increase_rate': 1.2  # увеличение показателей в следующих периодах
}
