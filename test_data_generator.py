# -*- coding: utf-8 -*-
"""
Модуль для генерации тестовых данных
Создает Excel файлы с тестовыми данными для проверки работы программы
"""

import pandas as pd
import numpy as np
import random
from typing import List, Tuple
from config import TEST_DATA_CONFIG
from logger import logger


class TestDataGenerator:
    """
    Класс для генерации тестовых данных
    Создает Excel файлы с реалистичными данными для тестирования
    """
    
    def __init__(self):
        """
        Инициализация генератора тестовых данных
        Загружает конфигурацию и настраивает параметры
        """
        self.config = TEST_DATA_CONFIG
        self.clients_count = self.config['clients_count']
        self.managers_count = self.config['managers_count']
        self.tb_count = self.config['tb_count']
        self.gosb_per_tb_range = self.config['gosb_per_tb']
        self.value_range = self.config['value_range']
        self.manager_change_rate = self.config['manager_change_rate']
        self.value_increase_rate = self.config['value_increase_rate']
        
        # Генерация базовых данных
        self._generate_base_data()
        
        logger.debug("Генератор тестовых данных инициализирован")
    
    def _generate_base_data(self) -> None:
        """
        Генерация базовых данных для тестирования
        Создает списки менеджеров, ТБ, ГОСБ и клиентов
        """
        logger.debug("Генерация базовых данных")
        
        # Генерация списка менеджеров
        self.managers = []
        for i in range(self.managers_count):
            manager = {
                'tab_number': i + 1,
                'fio': f"Менеджер_{i+1:04d}",
                'tb': f"ТБ_{random.randint(1, self.tb_count):02d}",
                'gosb': f"ГОСБ_{random.randint(1, random.randint(*self.gosb_per_tb_range)):02d}"
            }
            self.managers.append(manager)
        
        # Генерация списка клиентов
        self.clients = []
        for i in range(self.clients_count):
            client = {
                'client_id': f"CLIENT_{i+1:06d}",
                'client_name': f"Клиент_{i+1:06d}"
            }
            self.clients.append(client)
        
        logger.debug(f"Сгенерировано {len(self.managers)} менеджеров и {len(self.clients)} клиентов")
    
    def _get_random_manager(self) -> dict:
        """
        Получение случайного менеджера
        
        Returns:
            dict: Словарь с данными менеджера
        """
        return random.choice(self.managers)
    
    def _get_random_client(self) -> dict:
        """
        Получение случайного клиента
        
        Returns:
            dict: Словарь с данными клиента
        """
        return random.choice(self.clients)
    
    def _generate_value(self, base_value: float = None) -> float:
        """
        Генерация случайного показателя
        
        Args:
            base_value (float, optional): Базовое значение для расчета
            
        Returns:
            float: Сгенерированный показатель
        """
        if base_value is None:
            return random.uniform(*self.value_range)
        else:
            # Увеличение показателя с некоторой вариацией
            increase = random.uniform(1.1, self.value_increase_rate)
            return base_value * increase
    
    def _should_change_manager(self) -> bool:
        """
        Определение необходимости смены менеджера
        
        Returns:
            bool: True если менеджер должен смениться
        """
        return random.random() < self.manager_change_rate
    
    def generate_period_data(self, period: int) -> pd.DataFrame:
        """
        Генерация данных для конкретного периода
        
        Args:
            period (int): Номер периода (1, 2 или 3)
            
        Returns:
            pd.DataFrame: Данные периода
        """
        logger.debug(f"Генерация данных для периода {period}")
        
        data = []
        
        for client in self.clients:
            # Получение менеджера для клиента
            if period == 1:
                # В первом периоде случайный менеджер
                manager = self._get_random_manager()
            else:
                # В последующих периодах возможна смена менеджера
                if self._should_change_manager():
                    manager = self._get_random_manager()
                else:
                    # Используем того же менеджера (упрощение для тестирования)
                    manager = self._get_random_manager()
            
            # Генерация показателя
            if period == 1:
                value = self._generate_value()
            else:
                # Увеличение показателя в последующих периодах
                base_value = self._generate_value()
                value = self._generate_value(base_value)
            
            # Создание записи
            record = {
                'Таб. номер': manager['tab_number'],
                'Фамилия': manager['fio'],
                'коротко ТБ': manager['tb'],
                'короткое наименование ГОСБ': manager['gosb'],
                'ЕПК ИД': client['client_id'],
                'Наименование клиента': client['client_name'],
                'СДО руб': round(value, 2)
            }
            
            data.append(record)
        
        # Создание DataFrame
        df = pd.DataFrame(data)
        
        logger.debug(f"Сгенерировано {len(df)} записей для периода {period}")
        return df
    
    def create_test_files(self) -> None:
        """
        Создание тестовых Excel файлов
        Генерирует файлы для всех периодов согласно конфигурации
        """
        logger.info(LOG_MESSAGES['test_data_creation'])
        
        try:
            # Создание файлов для каждого периода
            for period in range(1, 4):  # 3 периода
                # Генерация данных
                df = self.generate_period_data(period)
                
                # Создание имени файла
                filename = f"test_data_period{period}.xlsx"
                
                # Сохранение в Excel файл
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Данные', index=False)
                
                logger.debug(f"Создан тестовый файл: {filename}")
            
            logger.info(LOG_MESSAGES['test_data_created'])
            print("Тестовые файлы созданы успешно:")
            print("- test_data_period1.xlsx")
            print("- test_data_period2.xlsx") 
            print("- test_data_period3.xlsx")
            
        except Exception as e:
            error_msg = f"Ошибка создания тестовых файлов: {str(e)}"
            logger.log_error(error_msg)
            raise Exception(error_msg)


def create_test_data():
    """
    Функция для создания тестовых данных
    Запускает генерацию тестовых Excel файлов
    """
    try:
        # Создание генератора
        generator = TestDataGenerator()
        
        # Создание тестовых файлов
        generator.create_test_files()
        
        return True
        
    except Exception as e:
        logger.log_error(f"Ошибка в функции создания тестовых данных: {str(e)}")
        return False


if __name__ == "__main__":
    # Импорт логгера для использования в main
    from logger import logger, LOG_MESSAGES
    
    # Создание тестовых данных
    success = create_test_data()
    if success:
        print("Тестовые данные созданы успешно!")
    else:
        print("Ошибка при создании тестовых данных!")
