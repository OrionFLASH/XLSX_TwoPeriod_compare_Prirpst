# -*- coding: utf-8 -*-
"""
Модуль для генерации тестовых данных
Создает Excel файлы с тестовыми данными для проверки работы программы
"""

import pandas as pd
import numpy as np
import random
from typing import List, Tuple
from config import TEST_DATA_CONFIG, IN_XLSX_DIR, TB_GOSB_CODES
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
        self.managers_count_min = self.config['managers_count_min']
        self.managers_count_max = self.config['managers_count_max']
        self.tb_codes = self.config['tb_codes']
        self.managers_per_gosb_min = self.config['managers_per_gosb_min']
        self.managers_per_gosb_max = self.config['managers_per_gosb_max']
        self.value_range = self.config['value_range']
        self.managers_change_rate = self.config['managers_change_rate']
        self.value_increase_rate = self.config['value_increase_rate']
        
        # Получаем реальные коды ТБ и ГОСБ
        self.tb_codes = list(TB_GOSB_CODES['tb_codes'].keys())
        self.gosb_codes = list(TB_GOSB_CODES['gosb_codes'].keys())
        
        # Создаем веса для пропорционального распределения
        self._create_distribution_weights()
        
        # Генерация базовых данных
        self._generate_base_data()
        
        logger.debug("Генератор тестовых данных инициализирован")
    
    def _create_distribution_weights(self) -> None:
        """
        Создание весов для пропорционального распределения клиентов и КМ по ТБ и ГОСБ
        """
        # Создаем веса для ТБ (пропорционально количеству ГОСБ в каждом ТБ)
        self.tb_weights = {}
        for tb_code in self.tb_codes:
            # Подсчитываем количество ГОСБ для каждого ТБ
            gosb_count = sum(1 for (tb, gosb) in self.gosb_codes if tb == tb_code)
            self.tb_weights[tb_code] = max(1, gosb_count)  # Минимум 1 для каждого ТБ
        
        # Создаем веса для ГОСБ (равномерно внутри каждого ТБ)
        self.gosb_weights = {}
        for (tb_code, gosb_code) in self.gosb_codes:
            self.gosb_weights[(tb_code, gosb_code)] = 1
        
        logger.debug(f"Созданы веса для {len(self.tb_weights)} ТБ и {len(self.gosb_weights)} ГОСБ")
    
    def _generate_manager_fio(self) -> str:
        """
        Генерация ФИО менеджера из списков имен, фамилий и отчеств
        """
        import random
        
        first_name = random.choice(self.config['manager_names']['first_names'])
        last_name = random.choice(self.config['manager_names']['last_names'])
        middle_name = random.choice(self.config['manager_names']['middle_names'])
        
        return f"{last_name} {first_name} {middle_name}"
    
    def _generate_client_name(self, index: int) -> str:
        """
        Генерация читаемого названия клиента
        """
        from config import CLIENT_NAMES_CONFIG
        
        # Получаем настройки из конфигурации
        company_names = CLIENT_NAMES_CONFIG['company_names']
        prefixes = CLIENT_NAMES_CONFIG['prefixes']
        suffixes = CLIENT_NAMES_CONFIG['suffixes']
        
        # Выбираем базовое название
        base_name = company_names[index % len(company_names)]
        
        # Если индекс больше количества базовых названий, добавляем префикс или суффикс
        if index >= len(company_names):
            if index < len(company_names) + len(prefixes):
                prefix = prefixes[index - len(company_names)]
                return f"{prefix} {base_name}"
            elif index < len(company_names) + len(prefixes) + len(suffixes):
                suffix_index = index - len(company_names) - len(prefixes)
                suffix = suffixes[suffix_index % len(suffixes)]
                return f"{base_name} {suffix}"
            else:
                # Для очень больших индексов добавляем числовой суффикс
                number = (index - len(company_names) - len(prefixes) - len(suffixes)) // len(suffixes) + 1
                suffix = suffixes[index % len(suffixes)]
                return f"{base_name} {suffix} №{number}"
        
        return base_name
    
    def _generate_base_data(self) -> None:
        """
        Генерация базовых данных для тестирования
        Создает списки менеджеров, ТБ, ГОСБ и клиентов
        """
        logger.debug("Генерация базовых данных")
        
        # Генерация списка менеджеров с реальными кодами ТБ и ГОСБ
        self.managers = []
        
        # Создаем список всех доступных ГОСБ
        available_gosb = []
        for (tb_code, gosb_code) in self.gosb_codes:
            gosb_name = TB_GOSB_CODES['gosb_codes'].get((tb_code, gosb_code), "")
            available_gosb.append({
                'tb_code': tb_code,
                'gosb_code': gosb_code,
                'gosb_name': gosb_name,
                'tb_name': TB_GOSB_CODES['tb_codes'][tb_code]['short_name']
            })
        
        # Распределяем менеджеров по ГОСБ (5-18 менеджеров в каждом ГОСБ)
        import random
        manager_id = 1
        
        # Случайно выбираем ГОСБ для распределения менеджеров
        # Убеждаемся, что общее количество менеджеров попадает в диапазон 1500-1600
        total_managers_needed = random.randint(self.managers_count_min, self.managers_count_max)
        managers_created = 0
        
        # Перемешиваем ГОСБ для случайного распределения
        random.shuffle(available_gosb)
        
        # Сначала заполняем до минимума, затем добавляем до максимума
        for gosb_info in available_gosb:
            if managers_created >= total_managers_needed:
                break
                
            # Случайное количество менеджеров в ГОСБ (5-18)
            remaining_needed = total_managers_needed - managers_created
            if remaining_needed <= 0:
                break
                
            # Если осталось мало менеджеров, добавляем их в текущий ГОСБ
            if remaining_needed < self.managers_per_gosb_min:
                managers_in_gosb = remaining_needed
            else:
                max_in_gosb = min(self.managers_per_gosb_max, remaining_needed)
                managers_in_gosb = random.randint(self.managers_per_gosb_min, max_in_gosb)
            
            for i in range(managers_in_gosb):
                # Генерируем ФИО менеджера
                fio = self._generate_manager_fio()
                
                manager = {
                    'tab_number': str(manager_id).zfill(8),  # 8 знаков с лидирующими нулями
                    'fio': fio,
                    'tb': gosb_info['tb_name'],
                    'gosb': gosb_info['gosb_name']
                }
                self.managers.append(manager)
                manager_id += 1
                managers_created += 1
        
        # Генерация списка клиентов
        self.clients = []
        for i in range(self.clients_count):
            client = {
                'client_id': str(i + 1).zfill(20),  # 20 знаков с лидирующими нулями
                'client_name': self._generate_client_name(i)
            }
            self.clients.append(client)
        
        # Проверяем, что все менеджеры имеют табельные номера
        for manager in self.managers:
            if not manager['tab_number'] or manager['tab_number'] == '0' or manager['tab_number'] == '':
                logger.error(f"Найден менеджер без табельного номера: {manager}")
                raise ValueError(f"Менеджер без табельного номера: {manager}")
        
        logger.debug(f"Сгенерировано {len(self.managers)} менеджеров и {len(self.clients)} клиентов")
    
    def _select_weighted_tb(self) -> int:
        """
        Выбор ТБ с учетом весов для пропорционального распределения
        
        Returns:
            int: Код выбранного ТБ
        """
        tb_codes = list(self.tb_weights.keys())
        weights = list(self.tb_weights.values())
        
        # Нормализуем веса
        total_weight = sum(weights)
        normalized_weights = [w / total_weight for w in weights]
        
        return int(np.random.choice(tb_codes, p=normalized_weights))
    
    def _select_weighted_gosb(self, tb_code: int) -> int:
        """
        Выбор ГОСБ для заданного ТБ с учетом весов
        
        Args:
            tb_code: Код ТБ
            
        Returns:
            int: Код выбранного ГОСБ
        """
        # Получаем все ГОСБ для данного ТБ
        available_gosb = [(tb, gosb) for (tb, gosb) in self.gosb_codes if tb == tb_code]
        
        if available_gosb:
            # Выбираем случайный ГОСБ из доступных
            selected_tb, selected_gosb = random.choice(available_gosb)
            return int(selected_gosb)
        else:
            # Если нет ГОСБ для данного ТБ, берем первый доступный ГОСБ
            return int(self.gosb_codes[0][1])
    
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
        return random.random() < self.managers_change_rate
    
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
        
        for i, client in enumerate(self.clients):
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
            
            # Добавление некорректных данных для тестирования валидации
            tab_number = f"'{manager['tab_number']}"
            client_id = f"'{client['client_id']}"
            value_to_use = round(value, 2)
            
            # Добавляем некорректные значения для тестирования валидации
            manager_name = manager['fio']
            if i % 10 == 0:  # Каждый 10-й клиент получает некорректные данные
                if i % 30 == 0:  # grey_zone
                    tab_number = "grey_zone"
                    manager_name = "Серая зона"  # Специальное имя для серой зоны
                elif i % 30 == 10:  # пустое значение
                    tab_number = ""
                elif i % 30 == 20:  # дефис
                    tab_number = "-"
                
                if i % 20 == 0:  # Некорректные показатели
                    if i % 40 == 0:  # пустое значение
                        value_to_use = ""
                    elif i % 40 == 20:  # дефис
                        value_to_use = "-"
            
            # Создание записи
            record = {
                'Таб. номер': tab_number,
                'КМ': manager_name,
                'ТБ': manager['tb'],
                'ГОСБ': manager['gosb'],
                'ИНН': client_id,
                'Клиент': client['client_name'],
                'ФОТ': value_to_use
            }
            
            data.append(record)
        
        # Создание DataFrame
        df = pd.DataFrame(data)
        
        logger.debug(f"Сгенерировано {len(df)} записей для периода {period}")
        return df
    
    def create_test_files(self) -> bool:
        """
        Создание тестовых Excel файлов
        Генерирует файлы для всех периодов согласно конфигурации
        """
        logger.info("Создание тестовых данных")
        
        try:
            # Создание файлов для каждого периода
            for period in range(1, 4):  # 3 периода
                # Генерация данных
                df = self.generate_period_data(period)
                
                # Создание имени файла в каталоге IN_XLSX
                # Используем тестовые имена файлов из конфигурации
                from config import TEST_DATA_CONFIG
                test_files = TEST_DATA_CONFIG['test_files']
                
                if period == 1:
                    filename = IN_XLSX_DIR / test_files[0]  # test_period1.xlsx
                elif period == 2:
                    filename = IN_XLSX_DIR / test_files[1]  # test_period2.xlsx
                else:
                    filename = IN_XLSX_DIR / test_files[2]  # test_period3.xlsx
                
                # Сохранение в Excel файл
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Sheet1', index=False)
                
                logger.debug(f"Создан тестовый файл: {filename}")
            
            logger.info("Тестовые данные созданы")
            print("Тестовые файлы созданы успешно:")
            print("- test_data_period1.xlsx")
            print("- test_data_period2.xlsx") 
            print("- test_data_period3.xlsx")
            return True
            
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
    from logger import logger
    
    # Создание тестовых данных
    success = create_test_data()
    if success:
        print("Тестовые данные созданы успешно!")
    else:
        print("Ошибка при создании тестовых данных!")
