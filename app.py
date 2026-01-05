# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import win32com.client as win32
import pythoncom
from datetime import datetime, timedelta
import json
import os
import re
import requests
from functools import lru_cache
from io import BytesIO
import tempfile
import shutil
from docxtpl import DocxTemplate
import time
import sys
import logging
import threading
from html import unescape

# --- Настройка логирования ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('app.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# --- Конфигурация ---
CONFIG_FILE = "requirements.txt"
BIDS_FILE = "bids.json"
CARRIERS_FILE = "carriers.json"
OFFERS_FILE = "offers.json"
CONTRACTS_FILE = "contracts.json"
CARRIERS_INFO_FILE = "carriers_info.xlsx"

# --- Блокировка для потокобезопасной работы с файлами ---
file_lock = threading.Lock()

# --- Инициализация файлов ---
def init_files():
    """Создает необходимые файлы, если они отсутствуют"""
    try:
        for file in [BIDS_FILE, CARRIERS_FILE, OFFERS_FILE, CONTRACTS_FILE]:
            if not os.path.exists(file):
                with open(file, 'w', encoding='utf-8') as f:
                    json.dump([], f, ensure_ascii=False, indent=2)
                logger.info(f"Файл {file} успешно создан")
    except Exception as e:
        st.error(f"Ошибка при создании файла {file}: {str(e)}")
        logger.error(f"Ошибка при создании файла {file}: {str(e)}")

    # Инициализация файла с информацией о перевозчиках
    if not os.path.exists(CARRIERS_INFO_FILE):
        try:
            df_example = pd.DataFrame(columns=[
                'name', 'email', 'legal_name', 'inn', 'kpp', 'ogrn', 'address',
                'bank_name', 'bik', 'rs', 'ks', 'contract_number', 'contract_date'
            ])
            df_example.to_excel(CARRIERS_INFO_FILE, index=False)
            logger.info(f"Файл {CARRIERS_INFO_FILE} успешно создан")
        except Exception as e:
            st.error(f"Ошибка при создании файла {CARRIERS_INFO_FILE}: {str(e)}")
            logger.error(f"Ошибка при создании файла {CARRIERS_INFO_FILE}: {str(e)}")

init_files()

# --- Функции для работы с данными ---
def load_json_file(filename: str) -> list:
    """Загружает данные из JSON файла с блокировкой"""
    with file_lock:
        try:
            if not os.path.exists(filename):
                logger.warning(f"Файл {filename} не найден, создаем пустой список")
                return []
            with open(filename, 'r', encoding='utf-8') as f:
                data = json.load(f)
                logger.info(f"Успешно загружено {len(data)} записей из {filename}")
                return data
        except json.JSONDecodeError as e:
            st.error(f"Ошибка парсинга JSON в файле {filename}: {str(e)}")
            logger.error(f"Ошибка парсинга JSON в файле {filename}: {str(e)}")
            return []
        except Exception as e:
            st.error(f"Ошибка загрузки файла {filename}: {str(e)}")
            logger.error(f"Ошибка загрузки файла {filename}: {str(e)}")
            return []

def save_json_file(filename: str, data: list) -> bool:
    """Сохраняет данные в JSON файл с блокировкой и проверкой целостности"""
    with file_lock:
        try:
            # Создаем резервную копию перед сохранением
            backup_path = f"{filename}.backup"
            if os.path.exists(filename):
                shutil.copy2(filename, backup_path)
                logger.info(f"Создана резервная копия {backup_path}")
            # Сохраняем данные
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            # Проверяем, что файл действительно изменился
            if os.path.getsize(filename) > 0:
                logger.info(f"Успешно сохранено {len(data)} записей в {filename}")
                return True
            else:
                raise Exception("Файл сохранен, но имеет нулевой размер")
        except Exception as e:
            st.error(f"Ошибка сохранения файла {filename}: {str(e)}")
            logger.error(f"Ошибка сохранения файла {filename}: {str(e)}")
            # Восстанавливаем из резервной копии при ошибке
            if os.path.exists(backup_path):
                try:
                    shutil.copy2(backup_path, filename)
                    st.warning(f"Восстановлено состояние из резервной копии {backup_path}")
                    logger.info(f"Восстановлено состояние из резервной копии {backup_path}")
                except Exception as restore_error:
                    st.error(f"Не удалось восстановить из резервной копии: {str(restore_error)}")
                    logger.error(f"Не удалось восстановить из резервной копии: {str(restore_error)}")
            return False

# --- Функция инициализации Outlook ---
def init_outlook():
    """Инициализирует соединение с Outlook"""
    pythoncom.CoInitialize()
    return win32.Dispatch("Outlook.Application")

# --- Работа с Outlook ---
def send_email(to: str, subject: str, body_text: str, attachments=None) -> bool:
    """Отправляет email через Outlook с возможностью вложений"""
    try:
        outlook = init_outlook()
        mail = outlook.CreateItem(0)
        mail.To = to
        mail.Subject = subject
        mail.HTMLBody = body_text
        # Добавление вложений
        if attachments:
            for file in attachments:
                if os.path.exists(file):  # Проверка существования файла
                    mail.Attachments.Add(file)
                else:
                    st.warning(f"Файл {file} не найден и не будет прикреплен")
                    logger.warning(f"Файл {file} не найден и не будет прикреплен")
        mail.Send()
        logger.info(f"Email отправлен на {to} с темой '{subject}'")
        return True
    except Exception as e:
        st.error(f"Ошибка отправки: {str(e)}")
        logger.error(f"Ошибка отправки на {to}: {str(e)}")
        return False
    finally:
        pythoncom.CoUninitialize()

# --- Кэширование курсов валют ---
@lru_cache(maxsize=1)
def get_currency_rates():
    """Получает текущие курсы валют с кэшированием"""
    try:
        response = requests.get("https://www.cbr-xml-daily.ru/daily_json.js", timeout=10)
        response.raise_for_status()  # Проверка статуса ответа
        data = response.json()
        rates = {
            "USD": data["Valute"]["USD"]["Value"],
            "EUR": data["Valute"]["EUR"]["Value"],
            "date": data["Date"][:10]
        }
        logger.info(f"Курсы валют получены: USD={rates['USD']}, EUR={rates['EUR']}")
        return rates
    except requests.RequestException as e:
        st.warning(f"Не удалось получить курсы валют с API: {str(e)}. Используем резервные значения.")
        logger.warning(f"Не удалось получить курсы валют с API: {str(e)}")
        return {"USD": 90.0, "EUR": 100.0, "date": "N/A"}
    except KeyError as e:
        st.warning(f"Некорректный формат данных от API: {str(e)}. Используем резервные значения.")
        logger.warning(f"Некорректный формат данных от API: {str(e)}")
        return {"USD": 90.0, "EUR": 100.0, "date": "N/A"}
    except Exception as e:
        st.error(f"Неизвестная ошибка при получении курсов валют: {str(e)}")
        logger.error(f"Неизвестная ошибка при получении курсов валют: {str(e)}")
        return {"USD": 90.0, "EUR": 100.0, "date": "N/A"}

# --- Виджет курсов валют ---
def currency_rates_widget():
    """Отображает виджет курсов валют в сайдбаре с дельтой и подробной информацией"""
    st.sidebar.title("💰 Курсы валют")
    rates_today = get_currency_rates()
    # Получаем курсы за вчерашний день
    try:
        yesterday = datetime.now() - timedelta(days=1)
        yesterday_str = yesterday.strftime("%Y/%m/%d")
        url_yesterday = f"https://www.cbr-xml-daily.ru/archive/{yesterday_str}/daily_json.js"
        response = requests.get(url_yesterday, timeout=10)
        response.raise_for_status()
        data_yesterday = response.json()
        rates_yesterday = {
            "USD": data_yesterday["Valute"]["USD"]["Value"],
            "EUR": data_yesterday["Valute"]["EUR"]["Value"],
        }
    except Exception:
        # Если не удалось получить данные за вчера, используем приближенные значения
        rates_yesterday = {"USD": rates_today["USD"] * 0.995, "EUR": rates_today["EUR"] * 1.003}

    # Рассчитываем дельты
    try:
        usd_delta_value = rates_today["USD"] - rates_yesterday["USD"]
        usd_delta_percent = (usd_delta_value / rates_yesterday["USD"]) * 100
        usd_delta = f"{'+' if usd_delta_value >= 0 else ''}{usd_delta_value:.2f} ({'+' if usd_delta_percent >= 0 else ''}{usd_delta_percent:.2f}%)"
    except:
        usd_delta = "N/A"
    try:
        eur_delta_value = rates_today["EUR"] - rates_yesterday["EUR"]
        eur_delta_percent = (eur_delta_value / rates_yesterday["EUR"]) * 100
        eur_delta = f"{'+' if eur_delta_value >= 0 else ''}{eur_delta_value:.2f} ({'+' if eur_delta_percent >= 0 else ''}{eur_delta_percent:.2f}%)"
    except:
        eur_delta = "N/A"

    # Основные курсы с дельтами
    col1, col2 = st.sidebar.columns(2)
    with col1:
        st.metric("USD/RUB", f"{rates_today['USD']:.2f} ₽", delta=usd_delta)
    with col2:
        st.metric("EUR/RUB", f"{rates_today['EUR']:.2f} ₽", delta=eur_delta)

    # Блок с подробной информацией
    with st.sidebar.expander("ℹ️ Подробности"):
        st.write(f"Последнее обновление: {rates_today['date']}")
        st.write(f"Курс USD вчера: {rates_yesterday['USD']:.2f}")
        st.write(f"Курс EUR вчера: {rates_yesterday['EUR']:.2f}")
        st.write("""
        **Источник:** [ЦБ РФ API](https://www.cbr-xml-daily.ru )  
        **Кэширование:** 1 час  
        **Резервные значения:** USD=90.0, EUR=100.0
        """)

    # Кнопка обновления
    if st.sidebar.button("🔄 Обновить данные", type="secondary"):
        get_currency_rates.cache_clear()
        st.rerun()

# --- Получение информации о перевозчике из Excel ---
def get_carrier_info(carrier_name: str) -> dict:
    """Получает информацию о перевозчике из Excel файла"""
    try:
        if os.path.exists(CARRIERS_INFO_FILE):
            df = pd.read_excel(CARRIERS_INFO_FILE)
            # Используем точное совпадение по имени
            carrier_info = df[df['name'] == carrier_name]
            if not carrier_info.empty:
                info_dict = carrier_info.iloc[0].to_dict()
                logger.info(f"Информация о перевозчике {carrier_name} успешно получена")
                return info_dict
            else:
                logger.info(f"Информация о перевозчике {carrier_name} не найдена в базе")
        else:
            logger.warning(f"Файл {CARRIERS_INFO_FILE} не найден")
    except Exception as e:
        st.error(f"Ошибка при загрузке информации о перевозчике: {str(e)}")
        logger.error(f"Ошибка при загрузке информации о перевозчике {carrier_name}: {str(e)}")
    return {}

# --- Генерация договора ---
def generate_contract(bid_data: dict, offer_data: dict) -> str | None:
    """Генерирует договор на основе данных заявки и предложения"""
    try:
        # Создаем папку templates если ее нет
        os.makedirs("templates", exist_ok=True)
        template_path = os.path.join("templates", "template.docx")
        if not os.path.exists(template_path):
            raise FileNotFoundError("Шаблон договора не найден")
        doc = DocxTemplate(template_path)

        # Получаем информацию о перевозчике
        carrier_info = get_carrier_info(offer_data['sender'])
        # Используем email из carriers_info, если он там есть, иначе - из предложения
        carrier_email_for_template = carrier_info.get('email', offer_data.get('sender_email', ''))

        # Подготовка данных
        context = {
            'id': bid_data['id'],
            'date_created': datetime.now().strftime('%d.%m.%Y'),
            'carrier_name': offer_data['sender'],
            'carrier_email': carrier_email_for_template,
            'country_from': bid_data['details']['country_from'],
            'loading_address': bid_data['details']['loading_address'],
            'cargo_type': bid_data['details']['cargo_type'],
            'cargo_description': bid_data['details'].get('cargo_description', ''),
            'container_type': bid_data['details']['container_type'],
            'hs_code': bid_data['details']['hs_code'],
            'incoterm': bid_data['details']['incoterm'],
            'ready_date': bid_data['details']['ready_date'],
            'payment_terms': bid_data['details']['payment_terms'],
            'notes': bid_data['details']['notes'],

            # Инициализация всех полей стоимости
            'pre_carriage_cost': 0,
            'pre_carriage_currency': '',
            'othc_cost': 0,
            'othc_currency': '',
            'sea_freight_cost': 0,
            'sea_freight_currency': '',

            # Добавляем информацию о перевозчике
            'name': carrier_info.get('name', ''),
            'email': carrier_info.get('email', ''),
            'legal_name': carrier_info.get('legal_name', ''),
            'inn': carrier_info.get('inn', ''),
            'kpp': carrier_info.get('kpp', ''),
            'ogrn': carrier_info.get('ogrn', ''),
            'address': carrier_info.get('address', ''),
            'bank_name': carrier_info.get('bank_name', ''),
            'bik': carrier_info.get('bik', ''),
            'rs': carrier_info.get('rs', ''),
            'ks': carrier_info.get('ks', ''),
            'contract_number': carrier_info.get('contract_number', ''),
            'contract_date': carrier_info.get('contract_date', ''),
        }

        # Заполняем данные о стоимости из предложения
        for cost in offer_data.get('costs', []):
            item_name = cost.get('ITEM', '')
            if 'Pre-carriage' in item_name:
                context['pre_carriage_cost'] = cost.get('COST', 0)
                context['pre_carriage_currency'] = cost.get('CURRENCY', 'USD')
            elif 'OTHC' in item_name:
                context['othc_cost'] = cost.get('COST', 0)
                context['othc_currency'] = cost.get('CURRENCY', 'USD')
            elif 'Sea freight' in item_name:
                context['sea_freight_cost'] = cost.get('COST', 0)
                context['sea_freight_currency'] = cost.get('CURRENCY', 'USD')

        doc.render(context)

        # Сохранение во временный файл
        temp_dir = tempfile.gettempdir()
        contract_filename = f"contract_{bid_data['id']}_{offer_data['sender']}.docx"
        contract_path = os.path.join(temp_dir, contract_filename)
        doc.save(contract_path)

        # Сохраняем копию в папку contracts
        os.makedirs("contracts", exist_ok=True)
        final_contract_path = os.path.join("contracts", contract_filename)
        shutil.copy(contract_path, final_contract_path)

        # Сохраняем информацию о договоре
        contracts = load_json_file(CONTRACTS_FILE)
        new_contract = {
            "bid_id": bid_data['id'],
            "offer_id": offer_data.get('bid_id', ''),
            "carrier": offer_data['sender'],
            "date": datetime.now().isoformat(),
            "file_path": final_contract_path,
            "status": "generated"
        }
        contracts.append(new_contract)
        save_json_file(CONTRACTS_FILE, contracts)

        logger.info(f"Договор успешно сгенерирован для заявки {bid_data['id']} и перевозчика {offer_data['sender']}")
        return contract_path
    except Exception as e:
        st.error(f"Ошибка при генерации договора: {str(e)}")
        logger.error(f"Ошибка при генерации договора для заявки {bid_data.get('id', 'unknown')} и перевозчика {offer_data.get('sender', 'unknown')}: {str(e)}")
        return None

def format_bid_email(bid: dict) -> str:
    """Генерирует HTML-письмо совместимое со всеми почтовыми клиентами"""
    bid_id = bid['id']
    
    # Создаем таблицу стоимости с простыми полями
    cost_table_rows = ""
    cost_items = [
        "Pre-carriage",
        "OTHC",
        "Sea freight",
        "ЖД перевозка",
        "Если отправка прямым ЖД из Китая - (Прямое ЖД)",
        "Станционные затраты",
        "Доставка со станции"
    ]
    
    for item in cost_items:
        cost_table_rows += f"""
        <tr>
            <td style="border: 1px solid #cccccc; padding: 8px; text-align: left; font-family: Arial, sans-serif; font-size: 12px; background-color: #f9f9f9;">{item}</td>
            <td style="border: 1px solid #cccccc; padding: 8px; text-align: center; font-family: Arial, sans-serif; font-size: 12px; background-color: #f0f0f0; height: 25px;"></td>
            <td style="border: 1px solid #cccccc; padding: 8px; text-align: center; font-family: Arial, sans-serif; font-size: 12px; background-color: #f9f9f9;">USD / RUB / EUR</td>
        </tr>
        """
    
    html_body = f"""
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            /* Только базовые стили, которые поддерживаются везде */
        </style>
    </head>
    <body style="margin: 0; padding: 0; font-family: Arial, sans-serif; font-size: 14px; line-height: 1.4; color: #333333; background-color: #f6f6f6;">
        <!-- Контейнер с табличной версткой для совместимости -->
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: #f6f6f6;">
            <tr>
                <td align="center" style="padding: 20px 0;">
                    <!-- Основной контент -->
                    <table width="100%" max-width="600" cellpadding="0" cellspacing="0" border="0" style="background-color: #ffffff; border: 1px solid #dddddd;">
                        
                        <!-- Шапка -->
                        <tr>
                            <td style="background-color: #2C3E50; padding: 20px; text-align: center;">
                                <h1 style="color: #ffffff; font-size: 20px; font-weight: bold; margin: 0 0 5px 0;">Заявка на перевозку {bid_id}</h1>
                                <p style="color: #ffffff; font-size: 12px; margin: 0; opacity: 0.9;">Срок ответа: до 15:00 следующего дня</p>
                            </td>
                        </tr>
                        
                        <!-- Приветствие -->
                        <tr>
                            <td style="padding: 20px;">
                                <p style="margin: 0 0 10px 0;"><strong>Уважаемый партнер,</strong></p>
                                <p style="margin: 0; font-size: 13px;">Просим предоставить коммерческое предложение по заявке ниже:</p>
                            </td>
                        </tr>
                        
                        <!-- Основные данные -->
                        <tr>
                            <td style="padding: 0 20px;">
                                <p style="font-size: 14px; font-weight: bold; color: #2C3E50; margin: 0 0 10px 0; border-left: 3px solid #3498db; padding-left: 8px;">📌 Основные данные</p>
                                
                                <table width="100%" cellpadding="0" cellspacing="0" border="0">
                                    <tr>
                                        <td width="50%" style="padding: 2px 0;">
                                            <table width="100%" cellpadding="0" cellspacing="0" border="0">
                                                <tr>
                                                    <td width="40%" style="font-size: 12px; font-weight: bold; color: #555555;">ID заявки:</td>
                                                    <td style="font-size: 12px;">{bid_id}</td>
                                                </tr>
                                                <tr>
                                                    <td style="font-size: 12px; font-weight: bold; color: #555555;">Страна отправки:</td>
                                                    <td style="font-size: 12px;">{bid['details']['country_from']}</td>
                                                </tr>
                                                <tr>
                                                    <td style="font-size: 12px; font-weight: bold; color: #555555;">Incoterm:</td>
                                                    <td style="font-size: 12px;">{bid['details']['incoterm']}</td>
                                                </tr>
                                                <tr>
                                                    <td style="font-size: 12px; font-weight: bold; color: #555555;">Тип контейнера:</td>
                                                    <td style="font-size: 12px;">{bid['details']['container_type']}</td>
                                                </tr>
                                            </table>
                                        </td>
                                        <td width="50%" style="padding: 2px 0;">
                                            <table width="100%" cellpadding="0" cellspacing="0" border="0">
                                                <tr>
                                                    <td width="40%" style="font-size: 12px; font-weight: bold; color: #555555;">Номер заказа:</td>
                                                    <td style="font-size: 12px;">{bid.get('order_number', '—')}</td>
                                                </tr>
                                                <tr>
                                                    <td style="font-size: 12px; font-weight: bold; color: #555555;">Порт отправки:</td>
                                                    <td style="font-size: 12px;">{bid['details']['port_from']}</td>
                                                </tr>
                                                <tr>
                                                    <td style="font-size: 12px; font-weight: bold; color: #555555;">Готовность груза:</td>
                                                    <td style="font-size: 12px;">{bid['details']['ready_date']}</td>
                                                </tr>
                                                <tr>
                                                    <td style="font-size: 12px; font-weight: bold; color: #555555;">Способ доставки:</td>
                                                    <td style="font-size: 12px;">{bid['details']['delivery_method']}</td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        
                        <!-- Информация о грузе -->
                        <tr>
                            <td style="padding: 15px 20px 0 20px;">
                                <p style="font-size: 14px; font-weight: bold; color: #2C3E50; margin: 0 0 10px 0; border-left: 3px solid #3498db; padding-left: 8px;">📦 Информация о грузе</p>
                                
                                <table width="100%" cellpadding="0" cellspacing="0" border="0">
                                    <tr>
                                        <td width="50%" style="padding: 2px 0;">
                                            <table width="100%" cellpadding="0" cellspacing="0" border="0">
                                                <tr>
                                                    <td width="40%" style="font-size: 12px; font-weight: bold; color: #555555;">Тип груза:</td>
                                                    <td style="font-size: 12px;">{bid['details']['cargo_type']}</td>
                                                </tr>
                                            </table>
                                        </td>
                                        <td width="50%" style="padding: 2px 0;">
                                            <table width="100%" cellpadding="0" cellspacing="0" border="0">
                                                <tr>
                                                    <td width="40%" style="font-size: 12px; font-weight: bold; color: #555555;">Код ТНВЭД:</td>
                                                    <td style="font-size: 12px;">{bid['details'].get('hs_code', '—')}</td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                                
                                <p style="font-size: 12px; margin: 10px 0 0 0;">
                                    <strong>Адрес погрузки:</strong><br>
                                    {bid['details']['loading_address']}
                                </p>
                            </td>
                        </tr>
                        
                        <!-- Расчет стоимости -->
                        <tr>
                            <td style="padding: 15px 20px 0 20px;">
                                <p style="font-size: 14px; font-weight: bold; color: #2C3E50; margin: 0 0 10px 0; border-left: 3px solid #3498db; padding-left: 8px;">💰 Расчет стоимости</p>
                                <p style="font-size: 12px; margin: 0 0 10px 0;">Заполните стоимость по позициям:</p>
                                
                                <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse: collapse;">
                                    <thead>
                                        <tr>
                                            <th style="border: 1px solid #cccccc; padding: 8px; text-align: left; font-family: Arial, sans-serif; font-size: 12px; font-weight: bold; background-color: #f2f2f2;">Статья затрат</th>
                                            <th style="border: 1px solid #cccccc; padding: 8px; text-align: center; font-family: Arial, sans-serif; font-size: 12px; font-weight: bold; background-color: #f2f2f2;">Сумма</th>
                                            <th style="border: 1px solid #cccccc; padding: 8px; text-align: center; font-family: Arial, sans-serif; font-size: 12px; font-weight: bold; background-color: #f2f2f2;">Валюта</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {cost_table_rows}
                                    </tbody>
                                </table>
                            </td>
                        </tr>
                        
                        <!-- Условия оплаты -->
                        <tr>
                            <td style="padding: 15px 20px 0 20px;">
                                <p style="font-size: 14px; font-weight: bold; color: #2C3E50; margin: 0 0 10px 0; border-left: 3px solid #3498db; padding-left: 8px;">💳 Условия оплаты</p>
                                <p style="font-size: 12px; margin: 0;">{bid['details']['payment_terms']}</p>
                            </td>
                        </tr>
                        
                        <!-- Примечания -->
                        <tr>
                            <td style="padding: 15px 20px 0 20px;">
                                <p style="font-size: 14px; font-weight: bold; color: #2C3E50; margin: 0 0 10px 0; border-left: 3px solid #3498db; padding-left: 8px;">📝 Примечания</p>
                                <p style="font-size: 12px; margin: 0;">{bid['details'].get('notes', '—')}</p>
                            </td>
                        </tr>
                        
                        <!-- Важная информация -->
                        <tr>
                            <td style="padding: 15px 20px;">
                                <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: #fff8e1; border-left: 3px solid #ffc107;">
                                    <tr>
                                        <td style="padding: 12px;">
                                            <p style="font-size: 11px; color: #856404; margin: 0;">
                                                <strong>❗ Важно:</strong><br>
                                                • Укажите стоимость и валюту для каждой позиции (например: 1500 USD)<br>
                                                • Ответьте на это письмо, сохранив структуру<br>
                                                • Заявки, полученные после 15:00, не рассматриваются
                                            </p>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        
                        <!-- Футер -->
                        <tr>
                            <td style="padding: 15px 20px; text-align: center; border-top: 1px solid #eeeeee;">
                                <p style="font-size: 11px; color: #7f8c8d; margin: 0;">
                                    С уважением,<br>
                                    <strong>Логистический отдел</strong><br>
                                    📧 contact@logistics.company | 📞 +7 (495) 123-45-67
                                </p>
                            </td>
                        </tr>
                        
                    </table>
                </td>
            </tr>
        </table>
    </body>
    </html>
    """
    return html_body

# -------------------------
#  НОВЫЕ/УЛУЧШЕННЫЕ УТИЛИТЫ ДЛЯ ПАРСЕРА
# -------------------------
CURRENCY_SYMBOLS_MAP = {
    '$': 'USD',
    'USD': 'USD',
    'usd': 'USD',
    '€': 'EUR',
    'EUR': 'EUR',
    'eur': 'EUR',
    '₽': 'RUB',
    'RUB': 'RUB',
    'RUR': 'RUB',
    'rub': 'RUB',
    'руб': 'RUB',
    'CNY': 'CNY',
    '¥': 'CNY',
    'CN¥': 'CNY',
    'GBP': 'GBP',
    '£': 'GBP'
}

_amount_currency_regexes = [
    # 1 234,56 USD  or 1,234.56 USD or 1234.56USD
    re.compile(r'(?P<amt>\d{1,3}(?:[ ,.\u202f]\d{3})*(?:[.,]\d+)?|\d+(?:[.,]\d+)?)[ ]?(?P<cur>[A-Za-z₽$€¥£]{1,4})', re.IGNORECASE),
    # USD 1 234,56
    re.compile(r'(?P<cur>[A-Za-z₽$€¥£]{1,4})[ ]?(?P<amt>\d{1,3}(?:[ ,.\u202f]\d{3})*(?:[.,]\d+)?|\d+(?:[.,]\d+)?)', re.IGNORECASE),
    # 1500$ or 1500€
    re.compile(r'(?P<amt>[\d\.,\s]+)[ ]?(?P<cur_symbol>[$€₽¥£])'),
    # $1500 or €1500
    re.compile(r'(?P<cur_symbol>[$€₽¥£])[ ]?(?P<amt>[\d\.,\s]+)')
]

def clean_html_to_text(html: str) -> str:
    """Простейшая очистка HTML -> plain text (удаляет теги и декодирует сущности)"""
    if not html:
        return ""
    # убираем теги
    text = re.sub(r'(?s)<style.*?>.*?</style>', '', html)
    text = re.sub(r'(?s)<script.*?>.*?</script>', '', text)
    text = re.sub(r'<[^>]+>', ' ', text)
    # декодируем html сущности и сокращаем пробелы
    return unescape(re.sub(r'\s+', ' ', text)).strip()

def normalize_amount_str(amt_str: str) -> float | None:
    """Нормализует строку с числом (разные разделители тысяч/десятых) и возвращает float"""
    if not amt_str:
        return None
    s = amt_str.strip()
    # заменяем NBSP and unicode spaces
    s = s.replace('\u202f', '').replace('\xa0', '').replace(' ', '')
    # если есть оба '.', ',' — определяем разделитель десятичных
    if '.' in s and ',' in s:
        # если запятая стоит в конце как десятичный разделитель (e.g. '1.234,56'), считаем ',' десятичным
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    else:
        # если только запятая — считаем её десятичной
        if ',' in s and '.' not in s:
            s = s.replace(',', '.')
        # если только точки — оставляем как есть; если нет разделителей — ok
    try:
        return float(s)
    except Exception:
        try:
            # на крайний случай удалим все не-цифры и попробуем
            digits = re.sub(r'[^\d\.]', '', s)
            return float(digits) if digits else None
        except Exception:
            return None

def extract_amount_currency_from_text(text: str) -> list:
    """
    Возвращает список найденных (amount, currency) кортежей в тексте.
    Нормализует валюту в ISO-формат (USD/EUR/RUB/CNY/GBP).
    """
    results = []
    if not text:
        return results
    for rx in _amount_currency_regexes:
        for m in rx.finditer(text):
            if not m:
                continue
            amt = None
            cur = None
            if 'amt' in m.groupdict() and m.group('amt'):
                amt = normalize_amount_str(m.group('amt'))
            if 'cur' in m.groupdict() and m.group('cur'):
                cur_raw = m.group('cur').strip()
                cur = CURRENCY_SYMBOLS_MAP.get(cur_raw, cur_raw.upper())
            elif 'cur_symbol' in m.groupdict() and m.group('cur_symbol'):
                cur_sym = m.group('cur_symbol').strip()
                cur = CURRENCY_SYMBOLS_MAP.get(cur_sym, None)
            if amt is None:
                continue
            if cur is None:
                # Попробуем найти текст рядом (буквы после/до)
                # взять 3 символа справа/слева
                span_right = text[m.end():m.end()+6]
                span_left = text[max(0, m.start()-6):m.start()]
                look = (span_left + " " + span_right).strip()
                cur_search = re.search(r'([A-Za-z₽$€¥£]{1,4})', look)
                if cur_search:
                    cur_guess = cur_search.group(1)
                    cur = CURRENCY_SYMBOLS_MAP.get(cur_guess, cur_guess.upper())
                else:
                    # по умолчанию USD, но лучше вернуть None и позволить вызывающему коду использовать дефолты
                    cur = None
            results.append((amt, cur))
    # удаляем дубликаты (приблизительно)
    uniq = []
    seen = set()
    for amt, cur in results:
        key = (round(amt, 2), cur)
        if key not in seen:
            seen.add(key)
            uniq.append((amt, cur))
    return uniq

# -------------------------
#  УЛУЧШЕННАЯ ФУНКЦИЯ ПАРСЕРА ПРЕДЛОЖЕНИЙ
# -------------------------
def parse_offers_from_outlook(folder_name: str = "Предложения") -> list:
    """Парсит непрочитанные письма с предложениями из Outlook
    Возвращает список предложений с полями:
    - sender, sender_email, email_date, subject, bid_id, order_number, costs (список dict ITEM/COST/CURRENCY), original_body
    """
    try:
        outlook = init_outlook()
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)  # Inbox
        folder = inbox
        if folder_name != "Входящие":
            for i in range(1, inbox.Folders.Count + 1):
                try:
                    if inbox.Folders.Item(i).Name == folder_name:
                        folder = inbox.Folders.Item(i)
                        break
                except Exception:
                    continue
        messages = folder.Items
        # Сортируем по дате (по возрастанию) — чтобы старые обрабатывались первыми
        try:
            messages.Sort("[ReceivedTime]", True)
        except Exception:
            pass

        new_offers = []

        for msg in messages:
            try:
                # Берём только непрочитанные
                if not getattr(msg, "UnRead", False):
                    continue

                body = ""
                try:
                    body = msg.Body or ""
                except Exception:
                    body = ""

                # Если нет plain body, пробуем HTMLBody и очищаем теги
                if not body:
                    try:
                        html = msg.HTMLBody or ""
                        body = clean_html_to_text(html)
                    except Exception:
                        body = ""

                # Если HTMLBody и Body есть — добавим короткую версию HTML для контекста
                if not body and getattr(msg, "HTMLBody", None):
                    body = clean_html_to_text(getattr(msg, "HTMLBody"))

                sender_email = getattr(msg, "SenderEmailAddress", "") or ""
                if not (sender_email and "@" in sender_email and "." in sender_email.split("@")[-1]):
                    # пробуем получить SMTP через Sender
                    try:
                        sender_obj = getattr(msg, "Sender", None)
                        if sender_obj:
                            pa = sender_obj.PropertyAccessor
                            smtp_address = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
                            if smtp_address and "@" in smtp_address:
                                sender_email = smtp_address
                    except Exception:
                        pass

                offer_data = {
                    "sender": getattr(msg, "SenderName", "") or "",
                    "sender_email": sender_email,
                    "email_date": getattr(msg, "ReceivedTime", datetime.now()).strftime("%Y-%m-%d %H:%M:%S"),
                    "subject": getattr(msg, "Subject", "") or "",
                    "bid_id": "",
                    "order_number": "—",
                    "rate": "",
                    "currency": "",
                    "conditions": "",
                    "status": "Новое",
                    "costs": [],
                    "original_body": body
                }

                # Пытаемся спарсить ID заявки из subject и из тела
                id_patterns = [
                    r"(SHIP[-_ ]\d{8}[-_]\d{3,7})",
                    r"Заявка\s*на\s*перевозку\s*[:\s]*([A-Za-z0-9\-_]+)",
                    r"ID\s*заявки[^\w]*([A-Za-z0-9\-_]+)",
                    r"Заявка\s*№?\s*([A-Za-z0-9\-_]+)"
                ]
                # сначала subject
                for pattern in id_patterns:
                    m = re.search(pattern, offer_data["subject"], re.IGNORECASE)
                    if m:
                        offer_data["bid_id"] = m.group(1).strip()
                        break
                # если не в subject — ищем в теле
                if not offer_data["bid_id"]:
                    for pattern in id_patterns:
                        m = re.search(pattern, body, re.IGNORECASE)
                        if m:
                            offer_data["bid_id"] = m.group(1).strip()
                            break

                # Номер заказа
                order_patterns = [
                    r"Номер\s*заказа[^\w]*([A-Z0-9]{2,}\d*[-_]\d{3,})",
                    r"Номер\s*заказа[^\w]*([A-Za-z0-9\-_\s/]+)",
                    r"Order\s*Number[^\w]*([A-Za-z0-9\-_\s/]+)"
                ]
                for pattern in order_patterns:
                    m = re.search(pattern, offer_data["subject"], re.IGNORECASE)
                    if m:
                        offer_data["order_number"] = m.group(1).strip()
                        break
                if offer_data["order_number"] == "—":
                    for pattern in order_patterns:
                        m = re.search(pattern, body, re.IGNORECASE)
                        if m:
                            offer_data["order_number"] = m.group(1).strip()
                            break

                # Находим секцию "Расчет стоимости" (если есть)
                cost_section_start = -1
                for marker in ["💰 Расчет стоимости", "Расчет стоимости", "Расчёт стоимости", "Cost calculation", "Costs"]:
                    idx = body.find(marker)
                    if idx != -1:
                        cost_section_start = idx
                        break
                if cost_section_start == -1:
                    cost_section = body
                else:
                    # определяем конец секции (по следующему блоку или футеру)
                    cost_section_end = len(body)
                    for end_marker in ["💳 Условия оплаты", "Условия оплаты", "Примечания", "Notes", "Payment terms"]:
                        idx = body.find(end_marker, cost_section_start)
                        if idx != -1 and idx > cost_section_start:
                            cost_section_end = idx
                            break
                    cost_section = body[cost_section_start:cost_section_end]

                # Позиции затрат которые мы ожидаем
                cost_items = [
                    "Pre-carriage",
                    "OTHC",
                    "Sea freight",
                    "ЖД перевозка",
                    "Прямое ЖД",
                    "Прямое ЖД из Китая",
                    "Прямое Ж/Д",
                    "Прямое ЖД",
                    "Станционные затраты",
                    "Доставка со станции",
                    "Доставка от станции",
                    "Delivery from station",
                    "Station charges",
                    "Port Handling",
                    "Port charges"
                ]

                # Разбиваем секцию по строкам для локального анализа
                cost_lines = [ln.strip() for ln in re.split(r'[\r\n]+', cost_section) if ln.strip()]

                # Для каждой ожидаемой позиции — ищем в тексте и рядом с ней цену
                for item in cost_items:
                    # ищем индекс строки, где встречается item
                    found = False
                    for i, line in enumerate(cost_lines):
                        if item.lower() in line.lower():
                            # анализируем саму строку и соседние 2 строки на предмет суммы/валюты
                            window = " ".join(cost_lines[max(0, i-1): min(len(cost_lines), i+3)])
                            amounts = extract_amount_currency_from_text(window)
                            if amounts:
                                # возьмём первое подходящее значение (amt, cur)
                                amt, cur = amounts[0]
                                # если валюта отсутствует — попробуем искать возле строки (еще раз)
                                if not cur:
                                    # попытка: искать триггерные коды в самом окне
                                    cc = re.search(r'\b(USD|EUR|RUB|CNY|RUR|GBP|руб|₽)\b', window, re.IGNORECASE)
                                    if cc:
                                        cur = CURRENCY_SYMBOLS_MAP.get(cc.group(1), cc.group(1).upper())
                                if amt is not None:
                                    offer_data["costs"].append({
                                        "ITEM": item,
                                        "COST": round(amt, 2),
                                        "CURRENCY": cur if cur else ""
                                    })
                                    found = True
                                    break
                            # если ничего не найдено — пробуем взять любые числа на линии
                            else:
                                amounts_line = extract_amount_currency_from_text(line)
                                if amounts_line:
                                    amt, cur = amounts_line[0]
                                    offer_data["costs"].append({
                                        "ITEM": item,
                                        "COST": round(amt, 2),
                                        "CURRENCY": cur if cur else ""
                                    })
                                    found = True
                                    break
                    # end for lines
                    if not found:
                        # дополнительная попытка: найти пару "Item: 1500 USD" вблизи по всему блокe
                        pattern_any = re.compile(rf"{re.escape(item)}[^\\d]{{0,30}}([\\d\\s\.,]+)\\s*([A-Za-z₽$€¥£]{{1,4}})", re.IGNORECASE)
                        m = pattern_any.search(cost_section)
                        if m:
                            amt = normalize_amount_str(m.group(1))
                            cur_raw = m.group(2).strip()
                            cur = CURRENCY_SYMBOLS_MAP.get(cur_raw, cur_raw.upper())
                            if amt is not None:
                                offer_data["costs"].append({
                                    "ITEM": item,
                                    "COST": round(amt, 2),
                                    "CURRENCY": cur
                                })

                # Если не найдено ни одной позиции через секцию, пробуем глобальный поиск по всему письму:
                if not offer_data["costs"]:
                    # ищем все пары amount+currency в теле, затем попытаемся ассоциировать их с ближайшим item по вхождению
                    global_amounts = extract_amount_currency_from_text(body)
                    # ищем все строки, где есть сумма+валюта
                    body_lines = [ln.strip() for ln in re.split(r'[\r\n]+', body) if ln.strip()]
                    for line in body_lines:
                        amounts = extract_amount_currency_from_text(line)
                        if not amounts:
                            continue
                        # сопоставляем со словами позиции
                        matched_item = None
                        for item in cost_items:
                            if item.lower() in line.lower():
                                matched_item = item
                                break
                        if not matched_item:
                            # попытка сопоставить ближайшую позицию по контексту (ищем слово "Sea", "OTHC", "Pre")
                            if re.search(r'\bsea\b', line, re.IGNORECASE):
                                matched_item = "Sea freight"
                            elif re.search(r'\bothc\b', line, re.IGNORECASE):
                                matched_item = "OTHC"
                            elif re.search(r'pre[-\s]?carriage', line, re.IGNORECASE):
                                matched_item = "Pre-carriage"
                            elif re.search(r'жд|ж/д|rail', line, re.IGNORECASE):
                                matched_item = "ЖД перевозка"
                        for amt, cur in amounts:
                            if matched_item:
                                offer_data["costs"].append({
                                    "ITEM": matched_item,
                                    "COST": round(amt, 2),
                                    "CURRENCY": cur if cur else ""
                                })
                            else:
                                # если вообще не сопоставить — добавим как 'Other'
                                offer_data["costs"].append({
                                    "ITEM": "Other",
                                    "COST": round(amt, 2),
                                    "CURRENCY": cur if cur else ""
                                })

                # Если у нас есть bid_id — считаем это полноценным предложением
                if offer_data["bid_id"]:
                    new_offers.append(offer_data)
                    # Помечаем письмо как прочитанное
                    try:
                        msg.UnRead = False
                    except Exception:
                        pass

            except Exception as e_msg:
                logger.error(f"Ошибка при обработке письма: {str(e_msg)}")
                continue

        # Сохраняем новые предложения в OFFERS_FILE
        if new_offers:
            try:
                existing_offers = load_json_file(OFFERS_FILE)
                all_offers = existing_offers + new_offers
                # Удаляем дубликаты по комбинации bid_id и sender
                unique_offers = []
                seen_keys = set()
                for offer in all_offers:
                    key = (offer.get("bid_id", ""), offer.get("sender", ""), offer.get("sender_email", ""))
                    if key not in seen_keys:
                        seen_keys.add(key)
                        unique_offers.append(offer)
                success = save_json_file(OFFERS_FILE, unique_offers)
                if success:
                    logger.info(f"Успешно добавлено {len(new_offers)} новых предложений")
                    return new_offers
                else:
                    st.error("Не удалось сохранить новые предложения")
                    return []
            except Exception as e:
                st.error(f"Ошибка сохранения предложений: {str(e)}")
                logger.error(f"Ошибка сохранения предложений: {str(e)}")
                return []
        else:
            st.info("Новых предложений не найдено")
            logger.info("Новых предложений не найдено")
            return []
    except Exception as e:
        st.error(f"Ошибка при парсинге писем: {str(e)}")
        logger.error(f"Ошибка при парсинге писем: {str(e)}")
        return []
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

# --- Форма заявки ---
def create_bid_form():
    """Форма создания новой заявки на перевозку"""
    with st.form("new_bid", clear_on_submit=True):
        st.subheader("🆕 Новая заявка на перевозку")
        col1, col2 = st.columns(2)
        with col1:
            bid_id = st.text_input("ID заявки*", value=f"SHIP-{datetime.now().strftime('%Y%m%d-%H%M')}")
            order_number = st.text_input("Номер заказа*", "IN00-000")
            country_from = st.text_input("Страна отправки*", "Китай")
            incoterm = st.selectbox("Условие отгрузки*", ["FOB", "CIF", "EXW", "DAP", "DDP", "CFR", "FCA", "CPT"])
        with col2:
            port_from = st.text_input("Порт отправки*", "Shanghai")
            ready_date = st.date_input("Дата готовности груза*", datetime.now())
            container_type = st.selectbox("Тип контейнера*", ["20 фут", "40 фут", "40 фут HQ", "Авто 20тн", "Сборный груз"])
            delivery_method = st.selectbox("Способ доставки*", ["Море+ЖД", "Прямое ЖД", "Авто", "Авиа"], key="delivery_method")

        st.subheader("📦 Детали груза")
        cargo_type = st.text_input("Тип груза*", "не опасный")
        hs_code = st.text_input("Код ТНВЭД", "")
        cargo_description = st.text_area(
            "Описание груза*",
            '''- наименование груза :
- вес (нетто/брутто):
- объём, количество грузовых мест: 
- вид упаковки:
- стоимость груза:''',
            height=150,
            key="cargo_description"
        )
        loading_address = st.text_area("Адрес погрузки*", 
                                     "16F, No.839, Sec.4, Taiwan Blvd., Xitun Dist., 407 Taichung, TAIWAN")

        st.subheader("📎 Файлы")
        uploaded_files = st.file_uploader(
            "Прикрепите файлы (например, спецификации груза)", 
            type=["pdf", "docx", "xlsx", "jpg", "png"], 
            accept_multiple_files=True
        )

        valid_files = []
        total_size = 0
        if uploaded_files:
            for file in uploaded_files:
                if file.size > 15 * 1024 * 1024:
                    st.warning(f"Файл {file.name} больше 15 МБ и будет проигнорирован.")
                else:
                    total_size += file.size
                    valid_files.append(file)
            if total_size > 15 * 1024 * 1024:
                st.error("Суммарный размер всех файлов превышает 15 МБ. Некоторые файлы будут игнорированы.")

        st.subheader("💳 Финансовые условия")
        payment_terms = st.text_area("Условия оплаты*", 
                                   "50% TT in advance / 50% 14 days after delivery")

        st.subheader("📊 Расчет стоимости")
        costs = []
        cost_items = [
            "Pre-carriage",
            "OTHC",
            "Sea freight",
            "ЖД перевозка",
            "Прямое ЖД",
            "Станционные затраты",
            "Доставка со станции"
        ]
        for item in cost_items:
            col1, col2, col3 = st.columns([3, 1, 1])
            with col1:
                st.text(item)
            with col2:
                cost = st.number_input(f"Сумма {item}", 
                                     key=f"cost_{item}", 
                                     min_value=0.0,
                                     value=0.0)
            with col3:
                currency = st.selectbox("Валюта", 
                                      ["USD", "RUB", "EUR"], 
                                      key=f"curr_{item}")
            costs.append({"ITEM": item, "COST": cost, "CURRENCY": currency})

        notes = st.text_area("Примечания", '''Простой на выгрузке оплачивается отдельно
прочие затраты включены в стоимость, требуется отметка СКК''')

        if st.form_submit_button("📤 Отправить заявку"):
            if not all([bid_id, country_from, port_from, cargo_type, loading_address, payment_terms]):
                st.error("Заполните обязательные поля (помечены *)")
            else:
                bid_data = {
                    "id": bid_id,
                    "order_number": order_number,
                    "date_created": datetime.now().isoformat(),
                    "status": "Новая",
                    "details": {
                        "country_from": country_from,
                        "incoterm": incoterm,
                        "port_from": port_from,
                        "ready_date": str(ready_date),
                        "container_type": container_type,
                        "cargo_type": cargo_type,
                        "cargo_description": cargo_description,
                        "delivery_method": delivery_method,
                        "hs_code": hs_code,
                        "loading_address": loading_address,
                        "payment_terms": payment_terms,
                        "notes": notes
                    },
                    "costs": costs
                }

                try:
                    bids = load_json_file(BIDS_FILE)
                    bids.append(bid_data)
                    save_json_file(BIDS_FILE, bids)

                    carriers = load_json_file(CARRIERS_FILE)
                    html_body = format_bid_email(bid_data)

                    success_count = 0
                    attachments = []
                    if valid_files:
                        for file in valid_files:
                            temp_file_path = os.path.join(tempfile.gettempdir(), file.name)
                            with open(temp_file_path, "wb") as f:
                                f.write(file.getbuffer())
                            attachments.append(temp_file_path)

                    for carrier in carriers:
                        email_list = []
                        if carrier['email']:
                            normalized_emails = carrier['email'].replace(';', ',').replace(':', ',')
                            email_list = [e.strip() for e in normalized_emails.split(',') if e.strip()]
                        for email in email_list:
                            if email and '@' in email:
                                if send_email(email, f"Заявка на перевозку {bid_id}", html_body, attachments):
                                    success_count += 1
                                else:
                                    st.error(f"Ошибка отправки для {carrier['name']} ({email})")
                            elif email:
                                st.warning(f"Некорректный email для {carrier['name']}: {email}")

                    # Удаление временных файлов
                    if attachments:
                        for file in attachments:
                            try:
                                os.remove(file)
                            except Exception as e:
                                st.warning(f"Не удалось удалить временный файл {file}: {e}")

                    if success_count > 0:
                        st.success(f"✅ Заявка {bid_id} создана! Уведомления отправлены {success_count} перевозчикам")
                        time.sleep(3)
                    else:
                        st.warning("⚠️ Заявка создана, но не удалось отправить уведомления перевозчикам")
                        time.sleep(3)

                except Exception as e:
                    st.error(f"Ошибка при сохранении заявки: {str(e)}")
                    logger.error(f"Ошибка при сохранении заявки {bid_id}: {str(e)}")
                    time.sleep(3)

# --- Просмотр предложений ---
def view_offers():
    """Отображает и управляет предложениями от перевозчиков"""
    st.subheader("📬 Поступившие предложения")
    
    # --- Стили для таблицы ---
    st.markdown("""
    <style>
    .stDataFrame {
        width: 100% !important;
        font-size: 13px;
    }
    .stDataFrame th, .stDataFrame td {
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        padding: 8px 10px;
        vertical-align: top;
    }
    .stDataFrame th {
        background-color: #2C3E50;
        color: white;
        font-weight: bold;
        text-align: center;
        font-size: 13px;
    }
    .stDataFrame tr:nth-child(even) {
        background-color: #ECF0F1;
    }
    .stDataFrame tr:hover {
        background-color: #D5DBDB;
        transition: background-color 0.3s ease;
    }
    .stDataFrame td {
        border: 1px solid #BDC3C7;
    }
    .stDataFrame th:nth-child(5), .stDataFrame th:nth-child(6), .stDataFrame th:nth-child(7),
    .stDataFrame th:nth-child(8), .stDataFrame th:nth-child(9), .stDataFrame th:nth-child(10),
    .stDataFrame th:nth-child(11) {
        min-width: 120px;
        max-width: 150px;
        word-wrap: break-word;
    }
    .stDataFrame th:nth-child(12) {
        min-width: 140px;
        max-width: 180px;
        word-wrap: break-word;
    }
    </style>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns([3, 1])
    with col1:
        if st.button("🔄 Обновить список предложений"):
            new_offers = parse_offers_from_outlook()
            if new_offers:
                st.success(f"🔍 Найдено {len(new_offers)} новых предложений")
                time.sleep(3)
            else:
                st.info("📭 Новых предложений не найдено")
                time.sleep(3)
            st.rerun()

    offers = load_json_file(OFFERS_FILE)
    if offers:
        rates = get_currency_rates()
        comparison_data = []
        for offer in offers:
            total_rub = 0.0
            costs_dict = {}
            for cost in offer.get('costs', []):
                if cost.get('ITEM'):
                    cost_value = cost.get('COST', 0)
                    currency = cost.get('CURRENCY', '')
                    if currency == "USD":
                        converted = cost_value * rates["USD"]
                    elif currency == "EUR":
                        converted = cost_value * rates["EUR"]
                    elif currency == "CNY":
                        # если нужен курс CNY — можно добавить
                        converted = cost_value * 13.0
                    else:
                        converted = cost_value
                    total_rub += converted
                    costs_dict[cost['ITEM']] = {
                        'COST': cost_value,
                        'CURRENCY': currency,
                        'COST_RUB': converted
                    }

            def format_cost(value, currency):
                try:
                    int_value = int(round(value))
                    formatted_value = f"{int_value:,}".replace(',', ' ')
                    return f"{formatted_value} {currency}" if currency else f"{formatted_value}"
                except Exception:
                    return f"{value} {currency}"

            total_rub_int = int(round(total_rub))
            total_rub_formatted = f"{total_rub_int:,}".replace(',', ' ') + " ₽"

            comparison_data.append({
                "Дата получения": offer["email_date"],
                "Перевозчик": offer["sender"],
                "ID заявки": offer["bid_id"],
                "Номер заказа": offer.get("order_number", "—"),
                "Pre-carriage": format_cost(costs_dict.get('Pre-carriage', {}).get('COST', 0), costs_dict.get('Pre-carriage', {}).get('CURRENCY', '')),
                "OTHC": format_cost(costs_dict.get('OTHC', {}).get('COST', 0), costs_dict.get('OTHC', {}).get('CURRENCY', '')),
                "Sea freight": format_cost(costs_dict.get('Sea freight', {}).get('COST', 0), costs_dict.get('Sea freight', {}).get('CURRENCY', '')),
                "ЖД перевозка": format_cost(costs_dict.get('ЖД перевозка', {}).get('COST', 0), costs_dict.get('ЖД перевозка', {}).get('CURRENCY', '')),
                "Прямое ЖД": format_cost(costs_dict.get('Прямое ЖД', {}).get('COST', 0), costs_dict.get('Прямое ЖД', {}).get('CURRENCY', '')),
                "Станционные затраты": format_cost(costs_dict.get('Станционные затраты', {}).get('COST', 0), costs_dict.get('Станционные затраты', {}).get('CURRENCY', '')),
                "Доставка со станции": format_cost(costs_dict.get('Доставка со станции', {}).get('COST', 0), costs_dict.get('Доставка со станции', {}).get('CURRENCY', '')),
                "Итого (RUB)": total_rub_formatted,
                "Статус": offer.get("status", "Новое"),
                "original_index": offers.index(offer)
            })

        df_comparison = pd.DataFrame(comparison_data)

        bid_id_filter = st.text_input("🔍 Фильтр по ID заявки")
        if bid_id_filter:
            df_comparison = df_comparison[df_comparison['ID заявки'].str.contains(bid_id_filter, case=False)]

        df_display = df_comparison.copy()
        df_display["Перевозчик"] = df_display["Перевозчик"].apply(lambda x: x[:20] if isinstance(x, str) else x)
        df_display["Дата получения"] = df_display["Дата получения"].apply(lambda x: x[:10] if isinstance(x, str) else x)

        for bid_id in df_display['ID заявки'].unique():
            min_value_str = df_display[df_display['ID заявки'] == bid_id]['Итого (RUB)'].min()
            mask = (df_display['ID заявки'] == bid_id) & (df_display['Итого (RUB)'] == min_value_str)
            df_display.loc[mask, 'Итого (RUB)'] = "⭐ " + df_display.loc[mask, 'Итого (RUB)']

        df_display_with_checkbox = df_display.copy()
        df_display_with_checkbox.insert(0, "Выбрать", False)

        edited_df = st.data_editor(
            df_display_with_checkbox,
            use_container_width=True,
            hide_index=True,
            disabled=[
                "Дата получения", "Перевозчик", "ID заявки", "Номер заказа", 
                "Pre-carriage", "OTHC", "Sea freight", "ЖД перевозка", 
                "Прямое ЖД", "Станционные затраты", "Доставка со станции", 
                "Итого (RUB)"
            ],
            column_config={
                "Статус": st.column_config.SelectboxColumn(
                    "Статус",
                    options=["Новое", "В работе", "Отклонено", "Принято"],
                    required=True
                ),
                "Выбрать": st.column_config.CheckboxColumn(
                    "Выбрать",
                    help="Выберите строки для удаления",
                    default=False
                )
            }
        )

        selected_rows = edited_df[edited_df["Выбрать"] == True]
        if not selected_rows.empty:
            if st.button("🗑️ Удалить выбранные строки"):
                try:
                    keys_to_remove = set(selected_rows['original_index'])
                    filtered_offers = [offer for idx, offer in enumerate(offers) if idx not in keys_to_remove]
                    success = save_json_file(OFFERS_FILE, filtered_offers)
                    if success:
                        st.success(f"✅ Выбранные предложения удалены! ({len(selected_rows)} шт.)")
                        time.sleep(3)
                        st.rerun()
                    else:
                        st.error("❌ Не удалось сохранить изменения после удаления")
                except Exception as e:
                    st.error(f"❌ Ошибка при удалении предложений: {str(e)}")
                    logger.error(f"Ошибка при удалении предложений: {str(e)}")
                    time.sleep(3)

        if st.button("🗑️ Удалить отклоненные"):
            try:
                rejected_indices = set()
                for _, row in edited_df.iterrows():
                    if row["Статус"] == "Отклонено":
                        rejected_indices.add(row["original_index"])
                filtered_offers = [offer for idx, offer in enumerate(offers) if idx not in rejected_indices]
                success = save_json_file(OFFERS_FILE, filtered_offers)
                if success:
                    st.success("✅ Отклоненные предложения удалены!")
                    time.sleep(3)
                    st.rerun()
                else:
                    st.error("❌ Не удалось сохранить изменения после удаления отклоненных")
            except Exception as e:
                st.error(f"❌ Ошибка при удалении отклоненных предложений: {str(e)}")
                logger.error(f"Ошибка при удалении отклоненных предложений: {str(e)}")
                time.sleep(3)

        st.markdown("### 📝 Генерация договора")
        if len(edited_df) > 0:
            selected_offer_idx = st.selectbox(
                "Выберите предложение для генерации договора",
                range(len(edited_df)),
                format_func=lambda x: f"{edited_df.iloc[x]['Перевозчик']} - {edited_df.iloc[x]['ID заявки']}"
            )

            col1, col2, col3 = st.columns(3)
            with col1:
                if st.button("🖨️ Сгенерировать договор"):
                    selected_row = edited_df.iloc[selected_offer_idx]
                    offers = load_json_file(OFFERS_FILE)
                    selected_offer = next(
                        (o for o in offers 
                         if o["bid_id"] == selected_row["ID заявки"] 
                         and o["sender"] == selected_row["Перевозчик"]),
                        None
                    )
                    if selected_offer:
                        bids = load_json_file(BIDS_FILE)
                        selected_bid = next((b for b in bids if b["id"] == selected_offer["bid_id"]), None)
                        if selected_bid:
                            contract_path = generate_contract(selected_bid, selected_offer)
                            if contract_path:
                                st.success("✅ Договор успешно сгенерирован!")
                                time.sleep(3)
                                with open(contract_path, "rb") as f:
                                    st.download_button(
                                        label="📥 Скачать договор",
                                        data=f,
                                        file_name=f"Договор_{selected_bid['id']}_{selected_offer['sender']}.docx",
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                    )
                            else:
                                st.error("❌ Ошибка при генерации договора")
                                time.sleep(3)
                        else:
                            st.error("❌ Не найдена соответствующая заявка")
                            time.sleep(3)
                    else:
                        st.error("❌ Не найдено соответствующее предложение")
                        time.sleep(3)

            with col2:
                if st.button("📤 Отправить поставщику"):
                    selected_row = edited_df.iloc[selected_offer_idx]
                    offers = load_json_file(OFFERS_FILE)
                    selected_offer = next(
                        (o for o in offers 
                         if o["bid_id"] == selected_row["ID заявки"] 
                         and o["sender"] == selected_row["Перевозчик"]),
                        None
                    )
                    if selected_offer:
                        bids = load_json_file(BIDS_FILE)
                        selected_bid = next((b for b in bids if b["id"] == selected_offer["bid_id"]), None)
                        if selected_bid:
                            contract_path = generate_contract(selected_bid, selected_offer)
                            if contract_path:
                                subject = f"Договор по заявке {selected_bid['id']}"
                                body = f"""
                                Уважаемый {selected_offer['sender']},
                                
                                Прикрепляем договор по заявке {selected_bid['id']}.
                                Просим подтвердить получение и согласие с условиями.
                                
                                С уважением,
                                Логистический отдел
                                """
                                if send_email(selected_offer['sender_email'], subject, body, attachments=[contract_path]):
                                    st.success("✅ Договор успешно отправлен!")
                                    time.sleep(3)
                                else:
                                    st.error("❌ Ошибка при отправке договора")
                                    time.sleep(3)
                            else:
                                st.error("❌ Ошибка при генерации договора")
                                time.sleep(3)
                        else:
                            st.error("❌ Не найдена соответствующая заявка")
                            time.sleep(3)
                    else:
                        st.error("❌ Не найдено соответствующее предложение")
                        time.sleep(3)

            with col3:
                if st.button("💾 Сохранить изменения статусов"):
                    try:
                        offers = load_json_file(OFFERS_FILE)
                        updated_offers = []
                        sent_notifications = set()
                        success_count = 0
                        now = datetime.now()
                        new_statuses = {}
                        for idx, row in edited_df.iterrows():
                            key = (row["ID заявки"], row["Перевозчик"])
                            new_statuses[key] = row["Статус"]
                        for offer in offers:
                            key = (offer["bid_id"], offer["sender"])
                            if key in new_statuses:
                                old_status = offer.get("status", "Новое")
                                new_status = new_statuses[key]
                                offer["status"] = new_status
                                if new_status in ["Отклонено", "Принято"] and old_status != new_status:
                                    last_change = offer.get("last_status_change")
                                    should_send = True
                                    if last_change:
                                        try:
                                            time_diff = (now - datetime.fromisoformat(last_change)).total_seconds()
                                            if time_diff < 60:
                                                st.warning(f"⚠️ Слишком частое обновление для {offer['sender']}")
                                                should_send = False
                                        except Exception:
                                            pass
                                    if offer['sender'] in sent_notifications:
                                        should_send = False
                                    if should_send:
                                        subject = f"Обновление статуса заявки {offer['bid_id']}"
                                        body = f"""Здравствуйте, {offer['sender']}!
Статус вашей заявки с ID {offer['bid_id']} изменён на "{new_status}".
Подробности:
- Перевозчик: {offer['sender']}
- ID заявки: {offer['bid_id']}
- Статус: {new_status}
С уважением,
Логистическая система
"""
                                        if send_email(offer['sender_email'], subject, body):
                                            st.success(f"✅ Уведомление отправлено {offer['sender']} (статус: {new_status})")
                                            success_count += 1
                                            offer["last_status_change"] = now.isoformat()
                                            sent_notifications.add(offer['sender'])
                                        else:
                                            st.warning(f"⚠️ Не удалось отправить уведомление для {offer['sender']}")
                            updated_offers.append(offer)
                        success = save_json_file(OFFERS_FILE, updated_offers)
                        if success:
                            st.success(f"✅ Статусы предложений обновлены! Уведомления отправлены {success_count} перевозчикам")
                            time.sleep(2)
                            st.rerun()
                        else:
                            st.error("❌ Не удалось сохранить изменения статусов")
                    except Exception as e:
                        st.error(f"❌ Ошибка при обновлении статусов: {str(e)}")
                        logger.error(f"Ошибка при обновлении статусов: {str(e)}")
                        time.sleep(3)

        def to_excel(df):
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='Предложения')
            writer.close()
            processed_data = output.getvalue()
            return processed_data

        excel_data = to_excel(edited_df)
        st.download_button(
            label="📊 Экспорт в Excel",
            data=excel_data,
            file_name=f"offers_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # --- Подпись под таблицей ---
        st.markdown("""
        <div style="text-align: center; margin-top: 20px; font-size: 12px; color: #7F8C8D; font-style: italic;">
            📌 Система управления логистикой • Версия 1.0 • © 2025 Логистический отдел
        </div>
        """, unsafe_allow_html=True)

    else:
        st.info("ℹ️ Нет данных о предложениях")

# --- Управление перевозчиками ---
def manage_carriers():
    """Управление списком перевозчиков"""
    st.subheader("🚛 Управление перевозчиками")
    try:
        carriers = load_json_file(CARRIERS_FILE)
        df = pd.DataFrame(carriers if carriers else [{"name": "", "email": "", "notes": ""}])
        with st.expander("📋 Текущий список перевозчиков"):
            edited_df = st.data_editor(
                df,
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "name": "Название компании",
                    "email": st.column_config.TextColumn(
                        "Email",
                        help="Должен быть валидный email адрес",
                        validate="^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
                    ),
                    "notes": st.column_config.TextColumn(
                        "Примечания",
                        help="Дополнительная информация"
                    )
                }
            )
            if st.button("💾 Сохранить изменения"):
                invalid_rows = []
                for idx, row in edited_df.iterrows():
                    if not row['name'] or not row['email'] or '@' not in row['email']:
                        invalid_rows.append(idx + 1)
                if not invalid_rows:
                    success = save_json_file(CARRIERS_FILE, edited_df.to_dict('records'))
                    if success:
                        st.success("✅ Список перевозчиков обновлен!")
                        time.sleep(3)
                        st.rerun()
                    else:
                        st.error("❌ Не удалось сохранить изменения списка перевозчиков")
                else:
                    st.error(f"Проверьте данные в строках {', '.join(map(str, invalid_rows))}: все поля должны быть заполнены, email должен быть корректным")
                    time.sleep(3)

        with st.expander("📤 Импорт/экспорт"):
            col1, col2 = st.columns(2)
            with col1:
                uploaded_file = st.file_uploader("Импорт из CSV", type=["csv"])
                if uploaded_file:
                    try:
                        import_df = pd.read_csv(uploaded_file)
                        if set(import_df.columns) >= {"name", "email"}:
                            success = save_json_file(CARRIERS_FILE, import_df.to_dict('records'))
                            if success:
                                st.success("✅ Данные успешно импортированы!")
                                time.sleep(3)
                                st.rerun()
                            else:
                                st.error("❌ Не удалось сохранить импортированные данные")
                        else:
                            st.error("CSV должен содержать колонки 'name' и 'email'")
                            time.sleep(3)
                    except Exception as e:
                        st.error(f"Ошибка импорта: {str(e)}")
                        logger.error(f"Ошибка импорта из CSV: {str(e)}")
                        time.sleep(3)

                uploaded_excel = st.file_uploader("Импорт из Excel", type=["xlsx"])
                if uploaded_excel:
                    try:
                        import_df = pd.read_excel(uploaded_excel)
                        if set(import_df.columns) >= {"name", "email", "notes"}:
                            success = save_json_file(CARRIERS_FILE, import_df.to_dict('records'))
                            if success:
                                st.success("✅ Данные успешно импортированы из Excel!")
                                time.sleep(3)
                                st.rerun()
                            else:
                                st.error("❌ Не удалось сохранить импортированные данные из Excel")
                        else:
                            st.error("Excel должен содержать колонки 'name', 'email', 'notes'")
                            time.sleep(3)
                    except Exception as e:
                        st.error(f"Ошибка импорта из Excel: {str(e)}")
                        logger.error(f"Ошибка импорта из Excel: {str(e)}")
                        time.sleep(3)

            with col2:
                csv = df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Экспорт в CSV",
                    data=csv,
                    file_name="carriers.csv",
                    mime="text/csv"
                )

                def to_excel_carriers(df):
                    output = BytesIO()
                    writer = pd.ExcelWriter(output, engine='xlsxwriter')
                    df.to_excel(writer, index=False, sheet_name='Перевозчики')
                    writer.close()
                    processed_data = output.getvalue()
                    return processed_data

                excel_data_carriers = to_excel_carriers(df)
                st.download_button(
                    label="📥 Экспорт в Excel",
                    data=excel_data_carriers,
                    file_name="carriers.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Ошибка при работе с перевозчиками: {str(e)}")
        logger.error(f"Ошибка при работе с перевозчиками: {str(e)}")
        time.sleep(3)

# --- README ---
def show_readme():
    """Отображает инструкцию по использованию системы"""
    st.subheader("📖 Руководство пользователя")
    with st.expander("1. Управление перевозчиками"):
        st.markdown("""
        **Функции блока "Управление перевозчиками":**
        - Добавление новых перевозчиков
        - Редактирование существующих записей
        - Удаление перевозчиков
        - Сохранение списка в Excel/CSV
        - Корректировка данных: "Название компании", "Email", "Примечания"

        **Как управлять списком:**
        1. Перейдите в раздел "Управление перевозчиками".
        2. Добавьте/удалите перевозчиков.
        3. Нажмите "Сохранить изменения".
        4. Для экспорта — кнопки "Экспорт в CSV" / "Экспорт в Excel".
        5. Для импорта — загрузите файл и нажмите "Сохранить изменения".
        """)
    with st.expander("2. Создание заявки"):
        st.markdown("""
        **Как создать новую заявку:**
        1. Перейдите в раздел "Создать заявку".
        2. Заполните обязательные поля (*).
        3. Укажите стоимость по каждому пункту.
        4. Нажмите "Отправить заявку".
        Система автоматически разошлет уведомления всем перевозчикам.
        """)
    with st.expander("3. Работа с предложениями"):
        st.markdown("""
        **Как работать с предложениями:**
        1. В разделе "Просмотр предложений" — нажмите "Обновить список".
        2. Фильтруйте по ID заявки.
        3. Изменяйте статусы: Новое / В работе / Отклонено / Принято.
        4. Для выбранного предложения — сгенерируйте договор.
        5. Сохраните изменения кнопкой "Сохранить изменения статусов".
        6. Удалите отклоненные предложения.
        7. Экспортируйте данные в Excel.
        """)

# --- Главный интерфейс ---
def main():
    """Основная функция приложения"""
    st.set_page_config(
        page_title="Логистическая система",
        layout="wide",
        page_icon="🚢"
    )

    if 'user' not in st.session_state:
        st.session_state.user = "admin"

    if st.session_state.user == "admin":
        # Логотип в шапке
        if os.path.exists("logo.png"):
            st.sidebar.image("logo.png", use_container_width=False, width=150)
        # Виджет курсов валют
        currency_rates_widget()

        st.sidebar.title(f"👤 {st.session_state.user}")
        menu = st.sidebar.radio(
            "Меню",
            ["README", "Управление перевозчиками", "Создать заявку", "Просмотр предложений"],
            index=0
        )
        if st.sidebar.button("🚪 Выйти"):
            pass

        # Подпись в сайдбаре
        if os.path.exists("Shi_Py.png"):
            st.sidebar.image("Shi_Py.png", use_container_width=False, width=60)
        st.sidebar.markdown("© 2025 Логистический отдел")

        # Основная навигация
        if menu == "README":
            show_readme()
        elif menu == "Создать заявку":
            create_bid_form()
        elif menu == "Просмотр предложений":
            view_offers()
        elif menu == "Управление перевозчиками":
            manage_carriers()
    else:
        st.error("Ошибка авторизации. Пользователь не определен.")

if __name__ == "__main__":
    main()
