# -*- coding: utf-8 -*-
"""
Модуль для настройки системы логирования
Обеспечивает запись логов в файл с различными уровнями детализации
"""

import logging
import os
from datetime import datetime
from config import LOG_CONFIG, LOG_MESSAGES


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
        
        # Создание форматтера
        formatter = logging.Formatter(LOG_CONFIG['format'])
        file_handler.setFormatter(formatter)
        
        # Добавление обработчика к логгеру
        self.logger.addHandler(file_handler)
        
        # Логирование запуска программы
        self.info(LOG_MESSAGES['program_start'])
        self.info(LOG_MESSAGES['config_loaded'])
    
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
        self.info(LOG_MESSAGES['file_loading'].format(file_path))
    
    def log_file_loaded(self, file_path):
        """
        Логирование успешной загрузки файла
        
        Args:
            file_path (str): Путь к загруженному файлу
        """
        self.info(LOG_MESSAGES['file_loaded'].format(file_path))
    
    def log_data_processing(self, file_path):
        """
        Логирование обработки данных
        
        Args:
            file_path (str): Путь к обрабатываемому файлу
        """
        self.debug(LOG_MESSAGES['data_processing'].format(file_path))
    
    def log_calculation_start(self):
        """
        Логирование начала расчета приростов
        """
        self.info(LOG_MESSAGES['calculation_start'])
    
    def log_calculation_end(self):
        """
        Логирование завершения расчета приростов
        """
        self.info(LOG_MESSAGES['calculation_end'])
    
    def log_output_creation(self, file_path):
        """
        Логирование создания выходного файла
        
        Args:
            file_path (str): Путь к создаваемому файлу
        """
        self.info(LOG_MESSAGES['output_creation'].format(file_path))
    
    def log_output_created(self, file_path):
        """
        Логирование успешного создания выходного файла
        
        Args:
            file_path (str): Путь к созданному файлу
        """
        self.info(LOG_MESSAGES['output_created'].format(file_path))
    
    def log_error(self, error_message):
        """
        Логирование ошибки
        
        Args:
            error_message (str): Сообщение об ошибке
        """
        self.error(LOG_MESSAGES['error_occurred'].format(error_message))
    
    def log_program_end(self):
        """
        Логирование завершения программы
        """
        self.info(LOG_MESSAGES['program_end'])


# Создание глобального экземпляра логгера
logger = Logger()
