# -*- coding: utf-8 -*-
"""
Основная программа для сравнения показателей по двум или трем периодам
Анализирует Excel файлы и рассчитывает приросты по клиентским менеджерам и клиентам
"""

import pandas as pd
import numpy as np
from typing import List, Dict, Tuple, Optional
from config import ANALYSIS_CONFIG, TEST_DATA_CONFIG
from logger import logger


class PeriodComparison:
    """
    Класс для сравнения показателей по периодам
    Обрабатывает Excel файлы и рассчитывает приросты
    """
    
    def __init__(self):
        """
        Инициализация класса сравнения периодов
        Загружает конфигурацию и настраивает параметры
        """
        self.config = ANALYSIS_CONFIG
        self.file_count = self.config['file_count']
        self.files_config = self.config['files']
        self.output_config = self.config['output']
        
        # Словари для хранения данных из каждого файла
        self.data_frames = {}
        self.clients_data = {}
        self.managers_data = {}
        
        logger.debug(f"Инициализация с количеством файлов: {self.file_count}")
    
    def load_excel_file(self, file_path: str, sheet_name: str, columns: Dict[str, str]) -> pd.DataFrame:
        """
        Загрузка данных из Excel файла
        
        Args:
            file_path (str): Путь к Excel файлу
            sheet_name (str): Название листа
            columns (Dict[str, str]): Словарь соответствия колонок
            
        Returns:
            pd.DataFrame: Загруженные данные
            
        Raises:
            FileNotFoundError: Если файл не найден
            Exception: При ошибке загрузки данных
        """
        try:
            logger.log_file_loading(file_path)
            
            # Загрузка данных из Excel файла
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # Переименование колонок согласно конфигурации
            df = df.rename(columns=columns)
            
            # Отладочная информация
            logger.debug(f"Колонки после переименования: {df.columns.tolist()}")
            
            # Очистка данных от пустых значений (после переименования)
            # Проверяем наличие необходимых колонок
            required_columns = ['client_id', 'value']
            available_columns = [col for col in required_columns if col in df.columns]
            logger.debug(f"Доступные колонки для очистки: {available_columns}")
            if available_columns:
                df = df.dropna(subset=available_columns)
            
            # Преобразование типов данных (после переименования)
            if 'tab_number' in df.columns:
                df['tab_number'] = pd.to_numeric(df['tab_number'], errors='coerce').fillna(0).astype(int)
            if 'value' in df.columns:
                df['value'] = pd.to_numeric(df['value'], errors='coerce').fillna(0)
            
            logger.log_file_loaded(file_path)
            logger.debug(f"Загружено {len(df)} записей из файла {file_path}")
            
            return df
            
        except FileNotFoundError:
            error_msg = f"Файл не найден: {file_path}"
            logger.log_error(error_msg)
            raise FileNotFoundError(error_msg)
        except Exception as e:
            error_msg = f"Ошибка загрузки файла {file_path}: {str(e)}"
            logger.log_error(error_msg)
            raise Exception(error_msg)
    
    def load_all_files(self) -> None:
        """
        Загрузка всех файлов согласно конфигурации
        Сохраняет данные в словарь data_frames
        """
        logger.debug("Начало загрузки всех файлов")
        
        for i, file_config in enumerate(self.files_config[:self.file_count]):
            try:
                df = self.load_excel_file(
                    file_config['path'],
                    file_config['sheet_name'],
                    file_config['columns']
                )
                self.data_frames[f'period_{i+1}'] = df
                logger.debug(f"Файл периода {i+1} загружен успешно")
                
            except Exception as e:
                logger.log_error(f"Ошибка загрузки файла периода {i+1}: {str(e)}")
                raise
    
    def create_clients_base(self) -> pd.DataFrame:
        """
        Создание базы клиентов из уникальных идентификаторов
        Объединяет данные из всех периодов
        
        Returns:
            pd.DataFrame: База клиентов с данными по всем периодам
        """
        logger.debug("Создание базы клиентов")
        
        # Сбор всех уникальных идентификаторов клиентов
        all_client_ids = set()
        for period_key, df in self.data_frames.items():
            all_client_ids.update(df['client_id'].unique())
        
        # Создание базового DataFrame с клиентами
        clients_base = pd.DataFrame({'client_id': list(all_client_ids)})
        
        # Добавление данных по каждому периоду
        for i, (period_key, df) in enumerate(self.data_frames.items()):
            period_num = i + 1
            
            # Создание словаря для маппинга клиент -> данные
            period_data = df.groupby('client_id').agg({
                'tab_number': 'first',
                'fio': 'first',
                'tb': 'first',
                'gosb': 'first',
                'client_name': 'first',
                'value': 'sum'
            }).reset_index()
            
            # Добавление колонок с данными периода
            clients_base = clients_base.merge(
                period_data[['client_id', 'tab_number', 'fio', 'tb', 'gosb', 'client_name', 'value']],
                on='client_id',
                how='left',
                suffixes=('', f'_period_{period_num}')
            )
            
            # Переименование колонок для ясности
            clients_base = clients_base.rename(columns={
                'tab_number': f'tab_number_period_{period_num}',
                'fio': f'fio_period_{period_num}',
                'tb': f'tb_period_{period_num}',
                'gosb': f'gosb_period_{period_num}',
                'client_name': f'client_name_period_{period_num}',
                'value': f'value_period_{period_num}'
            })
        
        # Заполнение пропущенных значений
        for i in range(1, self.file_count + 1):
            clients_base[f'tab_number_period_{i}'] = clients_base[f'tab_number_period_{i}'].fillna(0).astype(int)
            clients_base[f'value_period_{i}'] = clients_base[f'value_period_{i}'].fillna(0)
            clients_base[f'fio_period_{i}'] = clients_base[f'fio_period_{i}'].fillna('')
            clients_base[f'tb_period_{i}'] = clients_base[f'tb_period_{i}'].fillna('')
            clients_base[f'gosb_period_{i}'] = clients_base[f'gosb_period_{i}'].fillna('')
            clients_base[f'client_name_period_{i}'] = clients_base[f'client_name_period_{i}'].fillna('')
        
        # Определение итогового табельного номера
        clients_base['final_tab_number'] = 0
        clients_base['final_fio'] = ''
        clients_base['final_tb'] = ''
        clients_base['final_gosb'] = ''
        
        # Логика выбора итогового табельного: приоритет последним периодам
        for i in range(self.file_count, 0, -1):
            mask = (clients_base[f'tab_number_period_{i}'] != 0) & (clients_base['final_tab_number'] == 0)
            clients_base.loc[mask, 'final_tab_number'] = clients_base.loc[mask, f'tab_number_period_{i}']
            clients_base.loc[mask, 'final_fio'] = clients_base.loc[mask, f'fio_period_{i}']
            clients_base.loc[mask, 'final_tb'] = clients_base.loc[mask, f'tb_period_{i}']
            clients_base.loc[mask, 'final_gosb'] = clients_base.loc[mask, f'gosb_period_{i}']
        
        logger.debug(f"База клиентов создана: {len(clients_base)} уникальных клиентов")
        return clients_base
    
    def calculate_growth(self, clients_base: pd.DataFrame) -> pd.DataFrame:
        """
        Расчет приростов согласно формуле
        
        Args:
            clients_base (pd.DataFrame): База клиентов с данными по периодам
            
        Returns:
            pd.DataFrame: База клиентов с рассчитанными приростами
        """
        logger.log_calculation_start()
        
        # Расчет прироста в зависимости от количества периодов
        if self.file_count == 2:
            # Прирост = показатель в файле 2 - показатель в файле 1
            clients_base['growth'] = (
                clients_base['value_period_2'] - clients_base['value_period_1']
            )
        elif self.file_count == 3:
            # Прирост = (файл 3 - файл 2) - (файл 2 - файл 1)
            period_2_1 = clients_base['value_period_2'] - clients_base['value_period_1']
            period_3_2 = clients_base['value_period_3'] - clients_base['value_period_2']
            clients_base['growth'] = period_3_2 - period_2_1
        else:
            raise ValueError(f"Неподдерживаемое количество периодов: {self.file_count}")
        
        logger.log_calculation_end()
        logger.debug(f"Приросты рассчитаны для {len(clients_base)} клиентов")
        
        return clients_base
    
    def create_managers_summary(self, clients_base: pd.DataFrame) -> pd.DataFrame:
        """
        Создание сводки по менеджерам
        
        Args:
            clients_base (pd.DataFrame): База клиентов с приростами
            
        Returns:
            pd.DataFrame: Сводка по менеджерам
        """
        logger.debug("Создание сводки по менеджерам")
        
        # Группировка по итоговому табельному номеру
        managers_summary = clients_base.groupby('final_tab_number').agg({
            'final_fio': 'first',
            'final_tb': 'first',
            'final_gosb': 'first',
            'value_period_1': 'sum',
            'value_period_2': 'sum',
            'value_period_3': 'sum' if self.file_count == 3 else 'sum',
            'growth': 'sum'
        }).reset_index()
        
        # Переименование колонок для соответствия выходному формату
        managers_summary = managers_summary.rename(columns={
            'final_tab_number': 'tab_number',
            'final_fio': 'fio',
            'final_tb': 'tb',
            'final_gosb': 'gosb',
            'value_period_1': 'value_1',
            'value_period_2': 'value_2',
            'value_period_3': 'value_3' if self.file_count == 3 else None,
            'growth': 'total_growth'
        })
        
        # Удаление колонки value_3 если только 2 периода
        if self.file_count == 2:
            managers_summary = managers_summary.drop(columns=['value_3'], errors='ignore')
        
        logger.debug(f"Сводка по менеджерам создана: {len(managers_summary)} уникальных менеджеров")
        return managers_summary
    
    def create_output_file(self, clients_base: pd.DataFrame, managers_summary: pd.DataFrame) -> None:
        """
        Создание выходного Excel файла с результатами
        
        Args:
            clients_base (pd.DataFrame): База клиентов с приростами
            managers_summary (pd.DataFrame): Сводка по менеджерам
        """
        output_file = self.output_config['file_name']
        logger.log_output_creation(output_file)
        
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # Подготовка данных для листа клиентов
                clients_output = clients_base[[
                    'client_id',
                    'client_name_period_1',
                    'value_period_1',
                    'value_period_2',
                    'value_period_3' if self.file_count == 3 else None,
                    'growth',
                    'final_tab_number',
                    'final_fio',
                    'final_gosb',
                    'final_tb'
                ]].copy()
                
                # Удаление None колонок
                clients_output = clients_output.dropna(axis=1, how='all')
                
                # Переименование колонок для читаемости
                column_mapping = {
                    'client_id': 'Идентификатор клиента',
                    'client_name_period_1': 'Наименование клиента',
                    'value_period_1': 'Показатель 1',
                    'value_period_2': 'Показатель 2',
                    'value_period_3': 'Показатель 3',
                    'growth': 'Прирост',
                    'final_tab_number': 'Табельный итоговый',
                    'final_fio': 'ФИО итогового сотрудника',
                    'final_gosb': 'ГОСБ',
                    'final_tb': 'ТБ'
                }
                
                clients_output = clients_output.rename(columns=column_mapping)
                
                # Запись листа клиентов
                clients_output.to_excel(
                    writer,
                    sheet_name=self.output_config['sheets']['clients'],
                    index=False
                )
                
                # Подготовка данных для листа менеджеров
                managers_output = managers_summary.copy()
                
                # Переименование колонок для читаемости
                managers_column_mapping = {
                    'tab_number': 'Табельный уникальный',
                    'fio': 'ФИО',
                    'tb': 'ТБ',
                    'gosb': 'ГОСБ',
                    'value_1': 'Показатель 1',
                    'value_2': 'Показатель 2',
                    'value_3': 'Показатель 3',
                    'total_growth': 'Суммарный прирост'
                }
                
                managers_output = managers_output.rename(columns=managers_column_mapping)
                
                # Удаление колонки value_3 если только 2 периода
                if self.file_count == 2:
                    managers_output = managers_output.drop(columns=['Показатель 3'], errors='ignore')
                
                # Запись листа менеджеров
                managers_output.to_excel(
                    writer,
                    sheet_name=self.output_config['sheets']['managers'],
                    index=False
                )
            
            # Применение форматирования после записи всех данных
            self._apply_formatting_to_file(output_file)
            
            logger.log_output_created(output_file)
            logger.debug(f"Выходной файл создан: {output_file}")
            
        except Exception as e:
            error_msg = f"Ошибка создания выходного файла: {str(e)}"
            logger.log_error(error_msg)
            raise Exception(error_msg)
    
    def _apply_formatting_to_file(self, file_path: str) -> None:
        """
        Применение форматирования к созданному Excel файлу
        
        Args:
            file_path: Путь к Excel файлу
        """
        logger.debug("Применение форматирования к файлу")
        
        try:
            from openpyxl import load_workbook
            
            # Загрузка файла
            wb = load_workbook(file_path)
            
            # Получение настроек форматирования
            formatting_config = self.output_config.get('formatting', {})
            
            # Форматирование листа клиентов
            if 'clients' in formatting_config:
                clients_sheet = wb[self.output_config['sheets']['clients']]
                self._format_sheet_columns(clients_sheet, formatting_config['clients'])
            
            # Форматирование листа менеджеров
            if 'managers' in formatting_config:
                managers_sheet = wb[self.output_config['sheets']['managers']]
                self._format_sheet_columns(managers_sheet, formatting_config['managers'])
            
            # Сохранение файла с форматированием
            wb.save(file_path)
            wb.close()
            
            logger.debug("Форматирование применено успешно")
            
        except Exception as e:
            logger.log_error(f"Ошибка применения форматирования: {str(e)}")
            # Не прерываем выполнение, так как форматирование не критично
    
    def _format_sheet_columns(self, sheet, column_formats: dict) -> None:
        """
        Форматирование колонок конкретного листа
        
        Args:
            sheet: Лист Excel файла
            column_formats: Словарь с настройками форматирования колонок
        """
        from openpyxl.styles import Font, Alignment
        from openpyxl.utils import get_column_letter
        
        # Создание стилей для разных типов данных
        number_font = Font(name='Arial', size=10)
        number_alignment = Alignment(horizontal='right')
        
        text_font = Font(name='Arial', size=10)
        text_alignment = Alignment(horizontal='left')
        
        # Получение заголовков для поиска колонок
        headers = []
        for cell in sheet[1]:
            headers.append(cell.value)
        
        # Применение форматирования к каждой колонке
        for col_name, format_config in column_formats.items():
            # Поиск индекса колонки по имени
            col_idx = None
            for i, header in enumerate(headers, 1):
                if header == col_name:
                    col_idx = i
                    break
            
            if col_idx is None:
                logger.debug(f"Колонка '{col_name}' не найдена в листе")
                continue
                
            col_letter = get_column_letter(col_idx)
            logger.debug(f"Форматирование колонки {col_name} ({col_letter})")
            
            # Применение стиля и формата
            if format_config['type'] == 'number':
                # Применение числового формата
                for row in range(2, sheet.max_row + 1):
                    cell = sheet[f"{col_letter}{row}"]
                    cell.number_format = format_config.get('format', '#,##0.00')
                    cell.font = number_font
                    cell.alignment = number_alignment
            elif format_config['type'] == 'text':
                # Применение текстового стиля
                for row in range(2, sheet.max_row + 1):
                    cell = sheet[f"{col_letter}{row}"]
                    cell.font = text_font
                    cell.alignment = text_alignment
        
        # Автоподбор ширины колонок
        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            adjusted_width = min(max_length + 2, 50)  # Максимальная ширина 50
            sheet.column_dimensions[column_letter].width = adjusted_width
    
    def run_analysis(self) -> None:
        """
        Основной метод для запуска анализа
        Выполняет полный цикл обработки данных
        """
        try:
            logger.info("Начало анализа периодов")
            
            # Загрузка всех файлов
            self.load_all_files()
            
            # Создание базы клиентов
            clients_base = self.create_clients_base()
            
            # Расчет приростов
            clients_base = self.calculate_growth(clients_base)
            
            # Создание сводки по менеджерам
            managers_summary = self.create_managers_summary(clients_base)
            
            # Создание выходного файла
            self.create_output_file(clients_base, managers_summary)
            
            logger.info("Анализ завершен успешно")
            
        except Exception as e:
            logger.log_error(f"Ошибка в процессе анализа: {str(e)}")
            raise


def main():
    """
    Главная функция программы
    Запускает анализ периодов
    """
    try:
        # Создание экземпляра класса анализа
        analyzer = PeriodComparison()
        
        # Запуск анализа
        analyzer.run_analysis()
        
        logger.log_program_end()
        print("Анализ завершен успешно. Результаты сохранены в файл comparison_result.xlsx")
        
    except Exception as e:
        logger.log_error(f"Критическая ошибка: {str(e)}")
        print(f"Произошла ошибка: {str(e)}")
        return 1
    
    return 0


if __name__ == "__main__":
    exit(main())
