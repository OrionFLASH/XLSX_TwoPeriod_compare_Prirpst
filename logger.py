# -*- coding: utf-8 -*-
"""
Модуль для настройки системы логирования
Обеспечивает запись логов в файл с различными уровнями детализации
"""

import logging
import os
from datetime import datetime
from config import LOG_CONFIG


class Logger:
    """
    Класс для управления логированием в программе
    Поддерживает уровни INFO и DEBUG
    """
    
    def __init__(self):
        """
        Инициализация логгера
        Настраивает формат и уровень логирования
        """
        self.logger = logging.getLogger('comparison_logger')
        self.logger.setLevel(getattr(logging, LOG_CONFIG['level']))
        
        # Очистка существующих обработчиков
        self.logger.handlers.clear()
        
        # Создание обработчика для записи в файл
        file_handler = logging.FileHandler(LOG_CONFIG['file'], encoding='utf-8')
        file_handler.setLevel(getattr(logging, LOG_CONFIG['level']))
        
        # Создание обработчика для консоли
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)  # В консоль только INFO и выше
        
        # Создание форматтера для файла
        file_formatter = logging.Formatter(LOG_CONFIG['format'])
        file_handler.setFormatter(file_formatter)
        
        # Создание форматтера для консоли (более простой)
        console_formatter = logging.Formatter('%(asctime)s - %(message)s')
        console_handler.setFormatter(console_formatter)
        
        # Добавление обработчиков к логгеру
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)
        
        # Логирование запуска программы
        self.info("Запуск программы сравнения периодов")
        self.info("Конфигурация загружена")
    
    def info(self, message):
        """
        Запись информационного сообщения в лог
        
        Args:
            message (str): Сообщение для записи
        """
        self.logger.info(message)
    
    def debug(self, message):
        """
        Запись отладочного сообщения в лог
        
        Args:
            message (str): Сообщение для записи
        """
        self.logger.debug(message)
    
    def error(self, message):
        """
        Запись сообщения об ошибке в лог
        
        Args:
            message (str): Сообщение об ошибке
        """
        self.logger.error(message)
    
    def warning(self, message):
        """
        Запись предупреждения в лог
        
        Args:
            message (str): Сообщение предупреждения
        """
        self.logger.warning(message)
    
    def log_file_loading(self, file_path):
        """
        Логирование загрузки файла
        
        Args:
            file_path (str): Путь к загружаемому файлу
        """
        self.info(f"Загрузка файла: {file_path}")
    
    def log_file_loaded(self, file_path):
        """
        Логирование успешной загрузки файла
        
        Args:
            file_path (str): Путь к загруженному файлу
        """
        self.info(f"Файл загружен успешно: {file_path}")
    
    def log_data_processing(self, file_path):
        """
        Логирование обработки данных
        
        Args:
            file_path (str): Путь к обрабатываемому файлу
        """
        self.debug(f"Обработка данных из файла: {file_path}")
    
    def log_calculation_start(self):
        """
        Логирование начала расчета приростов
        """
        self.info("Начало расчета приростов")
    
    def log_calculation_end(self):
        """
        Логирование завершения расчета приростов
        """
        self.info("Расчет приростов завершен")
    
    def log_output_creation(self, file_path):
        """
        Логирование создания выходного файла
        
        Args:
            file_path (str): Путь к создаваемому файлу
        """
        self.info(f"Создание выходного файла: {file_path}")
    
    def log_output_created(self, file_path):
        """
        Логирование успешного создания выходного файла
        
        Args:
            file_path (str): Путь к созданному файлу
        """
        self.info(f"Выходной файл создан: {file_path}")
    
    def log_error(self, error_message):
        """
        Логирование ошибки
        
        Args:
            error_message (str): Сообщение об ошибке
        """
        self.error(f"Произошла ошибка: {error_message}")
    
    def log_program_end(self):
        """
        Логирование завершения программы
        """
        self.info("Программа завершена успешно")
    
    # Расширенные методы логирования для загрузки файлов
    def log_file_loading_start(self, filename):
        """Логирование начала загрузки файла"""
        self.info(f"Начало загрузки файла: {filename}")
        self.debug(f"Начинаем загрузку файла: {filename}")
    
    def log_file_load_error(self, filename, error):
        """Логирование ошибки загрузки файла"""
        self.error(f"Ошибка загрузки файла: {filename}")
        self.debug(f"Детали ошибки загрузки файла {filename}: {error}")
    
    def log_file_columns_renamed(self, filename, columns):
        """Логирование переименования колонок"""
        self.info(f"Колонки файла переименованы: {filename}")
        self.debug(f"Переименованы колонки в файле {filename}: {columns}")
    
    def log_file_data_cleaned(self, filename, rows_before, rows_after):
        """Логирование очистки данных"""
        self.info(f"Данные файла очищены от пустых значений: {filename}")
        self.debug(f"Очистка данных в файле {filename}: было {rows_before} строк, стало {rows_after} строк")
    
    def log_file_data_processed(self, filename, rows_count):
        """Логирование обработки данных"""
        self.info(f"Данные файла обработаны: {filename}")
        self.debug(f"Обработано {rows_count} строк в файле {filename}")
    
    # Методы логирования для анализа данных
    def log_analysis_start(self):
        """Логирование начала анализа"""
        self.info("Начало анализа данных")
        self.debug("Начинаем анализ данных")
    
    def log_analysis_complete(self):
        """Логирование завершения анализа"""
        self.info("Анализ завершен успешно")
        self.debug("Анализ данных завершен")
    
    def log_clients_base_created(self, count):
        """Логирование создания базы клиентов"""
        self.info(f"База клиентов создана: {count} записей")
        self.debug(f"Создана база клиентов с {count} уникальными клиентами")
    
    def log_growth_calculated(self, count):
        """Логирование расчета приростов"""
        self.info(f"Приросты рассчитаны: {count} записей")
        self.debug(f"Рассчитаны приросты для {count} записей")
    
    def log_managers_summary_created(self, count):
        """Логирование создания сводки по менеджерам"""
        self.info(f"Сводка по менеджерам создана: {count} записей")
        self.debug(f"Создана сводка по {count} менеджерам")
    
    def log_managers_deal_date_created(self, count):
        """Логирование создания сводки по менеджерам по дате сделки"""
        self.info(f"Сводка по менеджерам по дате сделки создана: {count} записей")
        self.debug(f"Создана сводка по {count} менеджерам по дате сделки")
    
    # Методы логирования для создания выходного файла
    def log_output_creation_start(self, filename):
        """Логирование начала создания выходного файла"""
        self.info(f"Начало создания выходного файла: {filename}")
        self.debug(f"Начинаем создание выходного файла: {filename}")
    
    def log_output_formatting_applied(self, filename):
        """Логирование применения форматирования"""
        self.info(f"Форматирование применено к файлу: {filename}")
        self.debug(f"Форматирование применено к файлу: {filename}")
    
    # Методы логирования для тестовых данных
    def log_test_files_deleted(self, files):
        """Логирование удаления старых тестовых файлов"""
        self.info(f"Старые тестовые файлы удалены: {len(files)}")
        self.debug(f"Удалены старые тестовые файлы: {files}")
    
    def log_test_files_created(self, files):
        """Логирование создания новых тестовых файлов"""
        self.info(f"Новые тестовые файлы созданы: {len(files)}")
        self.debug(f"Созданы новые тестовые файлы: {files}")
    
    def log_critical_error(self, message):
        """Логирование критической ошибки"""
        self.error(f"Критическая ошибка: {message}")
        self.debug(f"Детали критической ошибки: {message}")


# Создание глобального экземпляра логгера
logger = Logger()
