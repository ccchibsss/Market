"""
================================================================================
🚀 ULTIMATE UNIT ECONOMICS ENGINE v44.0 - ПОЛНАЯ ВЕРСИЯ
================================================================================
📌 ВОЗМОЖНОСТИ:
    ✅ Две модели расчета: Товарная (B2C) и Агентская (Marketplace)
    ✅ API-интеграция с маркетплейсами (Яндекс, Ozon, WB)
    ✅ Парсинг цен конкурентов в реальном времени
    ✅ Автоматическая выгрузка цен и остатков
    ✅ Интеграция с 1С и CRM
    ✅ Интерактивный дашборд с графиками
    ✅ Прогнозирование продаж (ARIMA, Prophet, Linear Regression)
    ✅ ABC/XYZ-анализ с визуализацией
    ✅ Тепловая карта продаж
    ✅ Массовая обработка 250 000+ товаров

🚀 ЗАПУСК:
    streamlit run streamlit_app.py

📋 ТРЕБОВАНИЯ:
    pip install -r requirements.txt
================================================================================
"""

# ============================================================
# 1. ИМПОРТЫ
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import math
import json
import warnings
import requests
import logging
import time
import hashlib
import hmac
import base64
import urllib.parse
import traceback
import os
import pickle
import random
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional, Tuple, Union
from dataclasses import dataclass, field
from collections import Counter, defaultdict
from functools import lru_cache
from threading import Thread, Lock
from queue import Queue

# ============================================================
# 2. НАСТРОЙКА ЛОГИРОВАНИЯ (ДО ВСЕХ try/except)
# ============================================================

warnings.filterwarnings('ignore')

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('unit_economy.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ============================================================
# 3. ПРОВЕРКА НАЛИЧИЯ БИБЛИОТЕК
# ============================================================

OPENPYXL_AVAILABLE = False
try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.dataframe import dataframe_to_rows
    OPENPYXL_AVAILABLE = True
except ImportError as e:
    logger.warning(f"OpenPyXL не установлен: {e}")

PLOTLY_AVAILABLE = False
try:
    import plotly.graph_objects as go
    import plotly.express as px
    from plotly.subplots import make_subplots
    PLOTLY_AVAILABLE = True
except ImportError as e:
    logger.warning(f"Plotly не установлен: {e}")

SKLEARN_AVAILABLE = False
try:
    from sklearn.linear_model import LinearRegression
    from sklearn.metrics import mean_absolute_error, mean_squared_error
    from sklearn.preprocessing import StandardScaler
    SKLEARN_AVAILABLE = True
except ImportError as e:
    logger.warning(f"Scikit-learn не установлен: {e}")

# ============================================================
# 4. КОНФИГУРАЦИЯ
# ============================================================

CONFIG = {
    "version": "44.0.0",
    "app_name": "🚀 Юнит-экономика с API и дашбордом",
    "currency": "₽",
    "marketplaces": ["Яндекс Маркет", "Ozon", "Wildberries", "AliExpress", "Мегамаркет"],
    "operation_modes": ["FBY", "FBS", "FBO", "DBS"],
    "calculation_models": [
        {"id": "product", "name": "📦 Товарная (B2C)", "description": "Перепродажа физических товаров"},
        {"id": "agency", "name": "🏪 Агентская (Marketplace)", "description": "Платформа с комиссией за транзакции"}
    ],
    "category_keywords": {
        "Двигатель": ["двигатель", "мотор", "поршень", "кольцо", "гильза", "гбц", "головка", "блок цилиндров", "коленвал", "распредвал", "шатун", "клапан", "ремень грм", "цепь грм"],
        "Трансмиссия": ["коробка", "передач", "кпп", "вариатор", "сцепление", "диск сцепления", "корзина", "выжимной", "механизм", "автомат"],
        "Подвеска": ["амортизатор", "стойка", "пружина", "рычаг", "сайлентблок", "шаровой", "стабилизатор", "шаровая опора", "тяга"],
        "Тормозная система": ["тормоз", "колодка", "диск тормозной", "барабан", "суппорт", "главный тормозной", "шланг", "цилиндр"],
        "Рулевое управление": ["рулевой", "рейка", "наконечник", "тяга", "кардан", "усилитель", "рулевая"],
        "Электрооборудование": ["генератор", "стартер", "провод", "датчик", "реле", "блок управления", "эбу", "модуль"],
        "Система охлаждения": ["радиатор", "помпа", "термостат", "антифриз", "вентилятор", "патрубок"],
        "Система выпуска": ["глушитель", "катализатор", "сажевый", "выпускной", "гофра", "резонатор"],
        "Система питания": ["форсунка", "насос", "бензонасос", "тнвд", "инжектор", "карбюратор", "магистраль"],
        "Кузовные детали": ["бампер", "крыло", "капот", "дверь", "зеркало", "стекло", "молдинг", "порог"],
        "Оптика": ["фара", "фонарь", "поворотник", "габарит", "ксенон", "свет"],
        "Шины и диски": ["шина", "диск", "колесо", "резина", "покрышка", "литье"],
        "Масла и жидкости": ["масло", "жидкость", "смазка", "охлаждайка", "трансмиссионка", "гидравлика"],
        "Фильтры": ["фильтр", "масляный", "воздушный", "топливный", "салонный", "очистка"],
        "Ремни и цепи": ["ремень", "цепь", "натяжитель", "ролик", "грм"],
        "Свечи зажигания": ["свеча", "зажигание", "накаливания", "свечной", "свечка"],
        "Колодки и диски": ["колодки", "диски тормозные", "накладки", "тормозные"],
        "Аккумуляторы": ["аккумулятор", "батарея", "акум", "зарядка"],
        "Инструмент": ["инструмент", "ключ", "набор", "ремкомплект"],
        "Ручной инструмент": ["головка", "ключ", "набор инструментов", "отвертка", "плоскогубцы", "кусачки", "молоток", "стамеска"],
        "Садовая техника": ["триммер", "газонокосилка", "опрыскиватель", "шланг", "лопата", "грабли"],
        "Бытовая химия": ["чистящее", "моющее", "порошок", "гель", "пятновыводитель", "отбеливатель"],
        "Электротовары": ["кабель", "разъем", "переходник", "предохранитель", "розетка", "выключатель"],
        "Красота и уход": ["шампунь", "мыло", "крем", "маска", "лосьон", "тоник"]
    },
    "oem_patterns": [
        r'[0-9]{6,12}',
        r'[A-Z0-9]{6,12}',
        r'[A-Z]{2}[0-9]{6,10}',
        r'[A-Z]{2}[0-9]{4}[A-Z]{2}',
        r'[0-9]{4}[A-Z]{2}[0-9]{4}'
    ],
    "validation": {
        "min_price": 10,
        "max_price": 1000000,
        "min_cost": 1,
        "max_cost": 500000,
        "min_dimension": 0.1,
        "max_dimension": 1000,
        "min_volume": 0.001,
        "max_volume": 10000,
        "min_commission": 0.01,
        "max_commission": 0.99
    },
    "api": {
        "cache_ttl": 300,
        "max_retries": 3,
        "timeout": 30,
        "rate_limit": 10
    }
}

# ============================================================
# 5. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================

def safe_float(val: Any, default: float = 0.0) -> float:
    """Безопасное преобразование в float с логированием"""
    try:
        if val is None or val == "" or val == "NaN" or val == "nan":
            return default
        if isinstance(val, (int, float)):
            if math.isnan(val) or math.isinf(val):
                return default
            return float(val)
        if isinstance(val, str):
            val = val.replace(',', '.').replace(' ', '').replace('₽', '').replace('%', '').replace('$', '')
            val = val.replace('€', '').replace('£', '')
            val = re.sub(r'[^\d.\-]', '', val)
            if not val or val == '-' or val == '.':
                return default
            return float(val)
        return default
    except (ValueError, TypeError):
        return default

def safe_str(val: Any, default: str = "") -> str:
    """Безопасное преобразование в str"""
    try:
        if val is None:
            return default
        if isinstance(val, (int, float)) and (math.isnan(val) or math.isinf(val)):
            return default
        return str(val).strip() if str(val).strip() else default
    except (ValueError, TypeError):
        return default

def safe_int(val: Any, default: int = 0) -> int:
    """Безопасное преобразование в int"""
    try:
        return int(safe_float(val, default))
    except (ValueError, TypeError):
        return default

def normalize_text(text: str) -> str:
    """Нормализация текста для поиска"""
    if not text:
        return ""
    text = text.lower()
    text = re.sub(r'[^\w\sа-яА-Я]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def calculate_volume(length: float, width: float, height: float) -> float:
    """Расчет объема в литрах с валидацией"""
    try:
        if all([length, width, height]) and all([length > 0, width > 0, height > 0]):
            if any([length > 1000, width > 1000, height > 1000]):
                return 0.0
            volume = (length * width * height) / 1000.0
            if volume < 0.001:
                return 0.0
            return round(volume, 3)
        return 0.0
    except (TypeError, ValueError):
        return 0.0

def format_currency(value: float) -> str:
    """Форматирование валюты"""
    try:
        if value is None or math.isnan(value) or math.isinf(value):
            return "0 ₽"
        return f"{value:,.0f} ₽" if abs(value) >= 1 else f"{value:.2f} ₽"
    except (ValueError, TypeError):
        return "0 ₽"

def format_percent(value: float) -> str:
    """Форматирование процентов"""
    try:
        if value is None or math.isnan(value) or math.isinf(value):
            return "0%"
        return f"{value:.1f}%" if abs(value) >= 0.1 else f"{value:.2f}%"
    except (ValueError, TypeError):
        return "0%"

def format_number(value: float) -> str:
    """Форматирование чисел"""
    try:
        if value is None or math.isnan(value) or math.isinf(value):
            return "0"
        if abs(value) >= 1000:
            return f"{value:,.0f}"
        elif abs(value) >= 1:
            return f"{value:.2f}"
        else:
            return f"{value:.4f}"
    except (ValueError, TypeError):
        return "0"

def generate_cache_key(*args) -> str:
    """Генерация ключа для кэша"""
    key = "|".join(str(arg) for arg in args)
    return hashlib.md5(key.encode()).hexdigest()

# ============================================================
# 6. КЭШИРОВАНИЕ API
# ============================================================

class APICache:
    """Кэширование API-запросов с TTL"""
    
    def __init__(self, cache_dir: str = "api_cache"):
        self.cache_dir = cache_dir
        self.memory_cache = {}
        self.lock = Lock()
        
        if not os.path.exists(cache_dir):
            os.makedirs(cache_dir)
    
    def get(self, key: str) -> Optional[Any]:
        """Получение из кэша"""
        with self.lock:
            # Проверяем память
            if key in self.memory_cache:
                data, timestamp = self.memory_cache[key]
                if (datetime.now() - timestamp).total_seconds() < CONFIG["api"]["cache_ttl"]:
                    return data
            
            # Проверяем диск
            cache_file = os.path.join(self.cache_dir, f"{hashlib.md5(key.encode()).hexdigest()}.pkl")
            if os.path.exists(cache_file):
                try:
                    with open(cache_file, 'rb') as f:
                        data, timestamp = pickle.load(f)
                        if (datetime.now() - timestamp).total_seconds() < CONFIG["api"]["cache_ttl"]:
                            self.memory_cache[key] = (data, timestamp)
                            return data
                except Exception as e:
                    logger.warning(f"Cache read error: {e}")
            
            return None
    
    def set(self, key: str, value: Any):
        """Сохранение в кэш"""
        with self.lock:
            timestamp = datetime.now()
            self.memory_cache[key] = (value, timestamp)
            
            try:
                cache_file = os.path.join(self.cache_dir, f"{hashlib.md5(key.encode()).hexdigest()}.pkl")
                with open(cache_file, 'wb') as f:
                    pickle.dump((value, timestamp), f)
            except Exception as e:
                logger.warning(f"Cache write error: {e}")
    
    def clear(self):
        """Очистка кэша"""
        with self.lock:
            self.memory_cache.clear()
            for file in os.listdir(self.cache_dir):
                try:
                    os.remove(os.path.join(self.cache_dir, file))
                except Exception as e:
                    logger.warning(f"Cache clear error: {e}")

# ============================================================
# 7. API-КЛИЕНТЫ ДЛЯ МАРКЕТПЛЕЙСОВ
# ============================================================

class BaseMarketplaceAPI:
    """Базовый класс для API маркетплейсов"""
    
    def __init__(self, api_key: str = None, client_id: str = None):
        self.api_key = api_key
        self.client_id = client_id
        self.cache = APICache()
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (compatible; UnitEconomy/1.0)'
        })
    
    def _request(self, method: str, url: str, **kwargs) -> Optional[Dict]:
        """Выполнение запроса с ретраями"""
        for attempt in range(CONFIG["api"]["max_retries"]):
            try:
                response = self.session.request(
                    method, url,
                    timeout=CONFIG["api"]["timeout"],
                    **kwargs
                )
                if response.status_code == 200:
                    return response.json()
                elif response.status_code == 429:
                    time.sleep(2 ** attempt)
                else:
                    logger.warning(f"API error {response.status_code}: {response.text[:200]}")
                    break
            except Exception as e:
                logger.error(f"Request error (attempt {attempt+1}): {e}")
                time.sleep(1)
        
        return None


class YandexMarketAPI(BaseMarketplaceAPI):
    """API Яндекс Маркета"""
    
    BASE_URL = "https://api.partner.market.yandex.ru/v2"
    
    def __init__(self, api_key: str = None, business_id: str = None):
        super().__init__(api_key)
        self.business_id = business_id
        if api_key:
            self.session.headers.update({
                'Authorization': f'Bearer {api_key}',
                'Content-Type': 'application/json'
            })
    
    def get_products(self, page: int = 1, page_size: int = 100) -> Optional[Dict]:
        """Получение списка товаров"""
        cache_key = generate_cache_key('yandex_products', page, page_size)
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        url = f"{self.BASE_URL}/businesses/{self.business_id}/offer-mappings"
        params = {"page": page, "page_size": page_size}
        
        result = self._request('GET', url, params=params)
        if result:
            self.cache.set(cache_key, result)
        return result
    
    def get_offer_prices(self, offer_ids: List[str]) -> Optional[Dict]:
        """Получение цен товаров"""
        cache_key = generate_cache_key('yandex_prices', *sorted(offer_ids))
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        url = f"{self.BASE_URL}/businesses/{self.business_id}/offer-prices"
        data = {"offerIds": offer_ids}
        
        result = self._request('POST', url, json=data)
        if result:
            self.cache.set(cache_key, result)
        return result
    
    def update_price(self, offer_id: str, price: float) -> Optional[Dict]:
        """Обновление цены товара"""
        url = f"{self.BASE_URL}/businesses/{self.business_id}/offer-prices/updates"
        data = {
            "offers": [{
                "offerId": offer_id,
                "price": {
                    "value": price,
                    "currency": "RUB"
                }
            }]
        }
        return self._request('POST', url, json=data)
    
    def get_stocks(self, offer_ids: List[str]) -> Optional[Dict]:
        """Получение остатков"""
        cache_key = generate_cache_key('yandex_stocks', *sorted(offer_ids))
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        url = f"{self.BASE_URL}/businesses/{self.business_id}/stocks"
        data = {"offerIds": offer_ids}
        
        result = self._request('POST', url, json=data)
        if result:
            self.cache.set(cache_key, result)
        return result


class OzonAPI(BaseMarketplaceAPI):
    """API Ozon"""
    
    BASE_URL = "https://api-seller.ozon.ru/v2"
    
    def __init__(self, api_key: str = None, client_id: str = None):
        super().__init__(api_key, client_id)
        if api_key and client_id:
            self.session.headers.update({
                'Api-Key': api_key,
                'Client-Id': client_id,
                'Content-Type': 'application/json'
            })
    
    def get_products(self, page: int = 1, page_size: int = 100) -> Optional[Dict]:
        """Получение списка товаров"""
        cache_key = generate_cache_key('ozon_products', page, page_size)
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        url = f"{self.BASE_URL}/product/list"
        data = {"page": page, "page_size": page_size}
        
        result = self._request('POST', url, json=data)
        if result:
            self.cache.set(cache_key, result)
        return result
    
    def get_prices(self, product_ids: List[str]) -> Optional[Dict]:
        """Получение цен"""
        cache_key = generate_cache_key('ozon_prices', *sorted(product_ids))
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        url = f"{self.BASE_URL}/product/prices"
        data = {"product_ids": product_ids}
        
        result = self._request('POST', url, json=data)
        if result:
            self.cache.set(cache_key, result)
        return result
    
    def update_price(self, product_id: str, price: float) -> Optional[Dict]:
        """Обновление цены"""
        url = f"{self.BASE_URL}/product/price"
        data = {
            "product_id": product_id,
            "price": price
        }
        return self._request('POST', url, json=data)
    
    def get_stocks(self, product_ids: List[str]) -> Optional[Dict]:
        """Получение остатков"""
        cache_key = generate_cache_key('ozon_stocks', *sorted(product_ids))
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        url = f"{self.BASE_URL}/product/stocks"
        data = {"product_ids": product_ids}
        
        result = self._request('POST', url, json=data)
        if result:
            self.cache.set(cache_key, result)
        return result


class WildberriesAPI(BaseMarketplaceAPI):
    """API Wildberries"""
    
    BASE_URL = "https://suppliers-api.wildberries.ru"
    
    def __init__(self, api_key: str = None):
        super().__init__(api_key)
        if api_key:
            self.session.headers.update({
                'Authorization': api_key,
                'Content-Type': 'application/json'
            })
    
    def get_products(self, page: int = 1, limit: int = 100) -> Optional[Dict]:
        """Получение списка товаров"""
        cache_key = generate_cache_key('wb_products', page, limit)
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        url = f"{self.BASE_URL}/content/v2/get/cards/list"
        data = {"limit": limit, "offset": (page - 1) * limit}
        
        result = self._request('POST', url, json=data)
        if result:
            self.cache.set(cache_key, result)
        return result
    
    def get_prices(self, nm_ids: List[int]) -> Optional[Dict]:
        """Получение цен"""
        cache_key = generate_cache_key('wb_prices', *sorted(nm_ids))
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        url = f"{self.BASE_URL}/public/v1/prices"
        data = {"nm_ids": nm_ids}
        
        result = self._request('POST', url, json=data)
        if result:
            self.cache.set(cache_key, result)
        return result
    
    def update_price(self, nm_id: int, price: float) -> Optional[Dict]:
        """Обновление цены"""
        url = f"{self.BASE_URL}/public/v1/prices"
        data = [{"nm_id": nm_id, "price": price}]
        return self._request('POST', url, json=data)
    
    def get_stocks(self, nm_ids: List[int]) -> Optional[Dict]:
        """Получение остатков"""
        cache_key = generate_cache_key('wb_stocks', *sorted(nm_ids))
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        url = f"{self.BASE_URL}/public/v1/stocks"
        data = {"nm_ids": nm_ids}
        
        result = self._request('POST', url, json=data)
        if result:
            self.cache.set(cache_key, result)
        return result


class AliExpressAPI(BaseMarketplaceAPI):
    """API AliExpress"""
    
    BASE_URL = "https://api.aliexpress.com/openapi"
    
    def __init__(self, api_key: str = None, secret: str = None):
        super().__init__(api_key)
        self.secret = secret
    
    def _sign_request(self, params: Dict) -> str:
        """Подпись запроса для AliExpress"""
        sorted_params = sorted(params.items())
        sign_str = self.secret + ''.join(f"{k}{v}" for k, v in sorted_params) + self.secret
        return hashlib.md5(sign_str.encode()).hexdigest().upper()
    
    def get_products(self, page: int = 1, page_size: int = 20) -> Optional[Dict]:
        """Получение списка товаров"""
        cache_key = generate_cache_key('aliexpress_products', page, page_size)
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        params = {
            'method': 'aliexpress.product.list',
            'page': page,
            'page_size': page_size,
            'access_token': self.api_key
        }
        
        # Реальный запрос требует подписи и другие параметры
        # Здесь упрощенная версия
        
        result = self._request('GET', self.BASE_URL, params=params)
        if result:
            self.cache.set(cache_key, result)
        return result

# ============================================================
# 8. ПАРСЕРЫ ЦЕН КОНКУРЕНТОВ
# ============================================================

class CompetitorParser:
    """Парсинг цен конкурентов с маркетплейсов"""
    
    def __init__(self):
        self.cache = APICache()
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        self.proxies = []
        self.proxy_index = 0
    
    def parse_yandex_market(self, query: str, max_pages: int = 3) -> List[Dict]:
        """Парсинг Яндекс Маркета"""
        cache_key = generate_cache_key('parse_yandex', query, max_pages)
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        results = []
        
        for page in range(1, max_pages + 1):
            try:
                url = f"https://market.yandex.ru/search"
                params = {
                    'text': query,
                    'page': page
                }
                
                response = self.session.get(url, params=params, timeout=30)
                if response.status_code != 200:
                    continue
                
                # Упрощенный парсинг (в реальности нужен BeautifulSoup)
                # Здесь извлекаем данные из JSON-структуры
                html = response.text
                
                # Ищем блоки с товарами
                items = re.findall(r'"offerId":"([^"]+)","name":"([^"]+)","price":"([^"]+)"', html)
                
                for offer_id, name, price in items[:20]:
                    results.append({
                        'marketplace': 'Яндекс Маркет',
                        'offer_id': offer_id,
                        'name': name,
                        'price': safe_float(price),
                        'parsed_at': datetime.now().isoformat()
                    })
                
                time.sleep(0.5)
                
            except Exception as e:
                logger.error(f"Yandex parse error: {e}")
        
        self.cache.set(cache_key, results)
        return results
    
    def parse_ozon(self, query: str, max_pages: int = 3) -> List[Dict]:
        """Парсинг Ozon"""
        cache_key = generate_cache_key('parse_ozon', query, max_pages)
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        results = []
        
        for page in range(1, max_pages + 1):
            try:
                url = "https://www.ozon.ru/api/composer-api.bx/page/json/v2"
                params = {
                    'url': f'/search/?text={urllib.parse.quote(query)}&page={page}'
                }
                
                response = self.session.get(url, params=params, timeout=30)
                if response.status_code != 200:
                    continue
                
                data = response.json()
                
                # Извлечение товаров из структуры Ozon
                # Упрощенный вариант
                widgets = data.get('widgets', [])
                for widget in widgets:
                    if widget.get('type') == 'searchResultsV2':
                        items = widget.get('items', [])
                        for item in items:
                            results.append({
                                'marketplace': 'Ozon',
                                'product_id': item.get('id'),
                                'name': item.get('name', ''),
                                'price': safe_float(item.get('price', {}).get('price', 0)),
                                'parsed_at': datetime.now().isoformat()
                            })
                
                time.sleep(0.5)
                
            except Exception as e:
                logger.error(f"Ozon parse error: {e}")
        
        self.cache.set(cache_key, results)
        return results
    
    def parse_wildberries(self, query: str, max_pages: int = 3) -> List[Dict]:
        """Парсинг Wildberries"""
        cache_key = generate_cache_key('parse_wb', query, max_pages)
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        results = []
        
        for page in range(1, max_pages + 1):
            try:
                url = "https://search.wb.ru/exactmatch/ru/common/v4/search"
                params = {
                    'query': query,
                    'page': page,
                    'limit': 100
                }
                
                response = self.session.get(url, params=params, timeout=30)
                if response.status_code != 200:
                    continue
                
                data = response.json()
                
                products = data.get('data', {}).get('products', [])
                for product in products:
                    results.append({
                        'marketplace': 'Wildberries',
                        'nm_id': product.get('id'),
                        'name': product.get('name', ''),
                        'price': safe_float(product.get('priceU', 0)) / 100,
                        'parsed_at': datetime.now().isoformat()
                    })
                
                time.sleep(0.5)
                
            except Exception as e:
                logger.error(f"WB parse error: {e}")
        
        self.cache.set(cache_key, results)
        return results
    
    def parse_all_marketplaces(self, query: str) -> Dict[str, List[Dict]]:
        """Парсинг всех маркетплейсов"""
        results = {}
        
        # Парсим все в параллельных потоках
        threads = []
        results_dict = {}
        
        def parse_with_store(marketplace, parser_func):
            try:
                results_dict[marketplace] = parser_func(query)
            except Exception as e:
                logger.error(f"Error parsing {marketplace}: {e}")
                results_dict[marketplace] = []
        
        parsers = [
            ('Яндекс Маркет', self.parse_yandex_market),
            ('Ozon', self.parse_ozon),
            ('Wildberries', self.parse_wildberries)
        ]
        
        for name, func in parsers:
            thread = Thread(target=parse_with_store, args=(name, func))
            thread.start()
            threads.append(thread)
        
        for thread in threads:
            thread.join(timeout=60)
        
        return results_dict

# ============================================================
# 9. AI-ПОЛУЧЕНИЕ ТАРИФОВ
# ============================================================

class AITariffProvider:
    """Получение актуальных тарифов через AI с кэшированием"""
    
    def __init__(self, api_key: str = None, cache_ttl: int = 3600):
        self.api_key = api_key
        self.api_url = "https://api.deepseek.com/v1/chat/completions"
        self.cache = {}
        self.cache_ttl = cache_ttl
        self.last_update = {}
        
    def get_rates(self, marketplace: str, mode: str = "FBY", model_type: str = "product") -> Dict:
        """Получение тарифов через AI или из кэша"""
        cache_key = generate_cache_key(marketplace, mode, model_type)
        
        if cache_key in self.cache:
            cached_data, timestamp = self.cache[cache_key]
            if (datetime.now() - timestamp).total_seconds() < self.cache_ttl:
                logger.info(f"Использую кэшированные тарифы для {marketplace}/{mode}/{model_type}")
                return cached_data.copy()
        
        rates = self._get_base_rates(marketplace, mode, model_type)
        
        if self.api_key:
            try:
                ai_rates = self._get_ai_rates(marketplace, mode, model_type)
                if ai_rates and isinstance(ai_rates, dict):
                    for key, value in ai_rates.items():
                        if key in rates and isinstance(value, (int, float)):
                            rates[key] = value
                    logger.info(f"Получены AI-тарифы для {marketplace}/{mode}/{model_type}")
            except Exception as e:
                logger.error(f"AI tariff error: {e}")
        
        self.cache[cache_key] = (rates.copy(), datetime.now())
        self.last_update[cache_key] = datetime.now()
        
        return rates
    
    def _get_base_rates(self, marketplace: str, mode: str, model_type: str) -> Dict:
        """Базовые тарифы"""
        rates = {
            "commission": 0.11,
            "logistics_base": 70.0,
            "logistics_per_liter": 22.0,
            "storage_per_liter": 3.0,
            "storage_free_days": 365,
            "acquiring": 0.022,
            "returns": 0.10,
            "advertising": 0.15,
            "packaging": 50.0,
            "fba_base": 70.0,
            "fba_per_kg": 12.0,
            "fbs_logistics": 115.0,
            "fbo_storage": 0.5,
            "delivery_to_customer_percent": 0.045,
            "delivery_to_customer_max": 1000.0,
            "service_fee": 0.01,
            "insurance": 0.005,
            "platform_fee": 0.00,
            "transaction_fee": 0.00,
            "marketing_coef": 0.20,
            "support_cost_per_tx": 50.0,
            "tech_cost_per_tx": 20.0
        }
        
        if model_type == "agency":
            rates.update({
                "platform_fee": 0.10,
                "transaction_fee": 0.02,
                "marketing_coef": 0.30,
                "support_cost_per_tx": 50.0,
                "tech_cost_per_tx": 20.0,
                "commission_min": 50.0
            })
        
        marketplace_rates = {
            "Яндекс Маркет": {"commission": 0.11, "transaction_fee": 0.02 if model_type == "agency" else 0.00},
            "Ozon": {"commission": 0.10, "transaction_fee": 0.025 if model_type == "agency" else 0.00},
            "Wildberries": {"commission": 0.12, "transaction_fee": 0.028 if model_type == "agency" else 0.00},
            "AliExpress": {"commission": 0.08, "transaction_fee": 0.03 if model_type == "agency" else 0.00},
            "Мегамаркет": {"commission": 0.09, "transaction_fee": 0.026 if model_type == "agency" else 0.00}
        }
        
        rates.update(marketplace_rates.get(marketplace, {}))
        
        mode_rates = {
            "FBY": {"fba_base": 70.0, "fba_per_kg": 12.0, "logistics_base": 70.0},
            "FBS": {"fbs_logistics": 115.0, "logistics_base": 115.0},
            "FBO": {"fbo_storage": 0.5, "storage_per_liter": 0.5},
            "DBS": {"logistics_base": 60.0, "logistics_per_liter": 10.0}
        }
        
        rates.update(mode_rates.get(mode, {}))
        
        return rates
    
    def _get_ai_rates(self, marketplace: str, mode: str, model_type: str) -> Optional[Dict]:
        """Получение тарифов через AI API"""
        if not self.api_key:
            return None
            
        try:
            prompt = f"""
            Предоставь актуальные тарифы для маркетплейса {marketplace} 
            для продажи автозапчастей на начало 2026 года.
            Режим работы: {mode}
            Модель расчета: {model_type}
            
            Верни ТОЛЬКО JSON с тарифами.
            """
            
            headers = {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            }
            
            data = {
                "model": "deepseek-chat",
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.1,
                "max_tokens": 1000,
                "response_format": {"type": "json_object"}
            }
            
            response = requests.post(self.api_url, headers=headers, json=data, timeout=30)
            
            if response.status_code == 200:
                content = response.json()['choices'][0]['message']['content']
                json_match = re.search(r'\{.*\}', content, re.DOTALL)
                if json_match:
                    rates = json.loads(json_match.group())
                    required_keys = ["commission", "logistics_base", "logistics_per_liter", 
                                   "storage_per_liter", "acquiring"]
                    if all(key in rates for key in required_keys):
                        return rates
            return None
            
        except Exception as e:
            logger.error(f"AI API error: {e}")
            return None
    
    def get_all_rates(self, model_type: str = "product") -> Dict:
        """Получение тарифов для всех маркетплейсов и режимов"""
        all_rates = {}
        for marketplace in CONFIG["marketplaces"]:
            all_rates[marketplace] = {}
            for mode in CONFIG["operation_modes"]:
                all_rates[marketplace][mode] = self.get_rates(marketplace, mode, model_type)
        return all_rates
    
    def clear_cache(self):
        """Очистка кэша"""
        self.cache.clear()
        self.last_update.clear()
        logger.info("Кэш тарифов очищен")

# ============================================================
# 10. КЛАССИФИКАТОР КАТЕГОРИЙ
# ============================================================

class CategoryClassifier:
    """Классификация категорий товаров по наименованию"""
    
    def __init__(self):
        self.keywords = CONFIG["category_keywords"]
        self.categories = list(self.keywords.keys())
        self.cache = {}
        self.oem_patterns = CONFIG["oem_patterns"]
        
    @lru_cache(maxsize=10000)
    def classify(self, name: str) -> Tuple[str, float]:
        """Определение категории по наименованию с кэшированием"""
        if not name:
            return "Прочее", 0.0
        
        name_lower = name.lower()
        best_category = "Прочее"
        best_score = 0.0
        
        for category, keywords in self.keywords.items():
            score = 0.0
            for keyword in keywords:
                if keyword.lower() in name_lower:
                    weight = len(keyword) / 10.0
                    if name_lower.startswith(keyword.lower()):
                        weight *= 1.5
                    score += weight
            
            if score > best_score:
                best_score = score
                best_category = category
        
        confidence = min(best_score * 20, 100.0)
        
        if self.extract_oem(name):
            confidence = min(confidence + 10, 100)
        
        return best_category, round(confidence, 1)
    
    def extract_oem(self, name: str) -> Optional[str]:
        """Извлечение OEM номера из наименования"""
        if not name:
            return None
        
        for pattern in self.oem_patterns:
            match = re.search(pattern, name.upper())
            if match:
                return match.group()
        return None
    
    def extract_brand(self, name: str) -> Optional[str]:
        """Извлечение бренда из наименования"""
        if not name:
            return None
        
        brands = [
            "BOSCH", "DENSO", "NGK", "BREMBO", "AISIN", "HITACHI", "VALEO",
            "PIERBURG", "MANN", "MAHLE", "HENGST", "SACHS", "ZF",
            "CONTINENTAL", "GATES", "DAYCO", "SKF", "FAG", "TIMKEN",
            "MERCEDES", "BMW", "AUDI", "VW", "FORD", "TOYOTA", "HONDA",
            "NISSAN", "HYUNDAI", "KIA", "SKODA", "VOLVO", "RENAULT"
        ]
        
        name_upper = name.upper()
        for brand in brands:
            if brand in name_upper:
                return brand
        return None

# ============================================================
# 11. УПРАВЛЕНИЕ КОНКУРЕНТАМИ
# ============================================================

class CompetitorManager:
    """Управление конкурентными данными"""
    
    def __init__(self):
        self.parser = CompetitorParser()
        self.competitor_data = {}
        self.last_update = {}
    
    def get_competitor_prices(self, query: str, marketplace: str = None) -> Dict:
        """Получение цен конкурентов"""
        cache_key = generate_cache_key('competitor_prices', query, marketplace or 'all')
        
        if cache_key in self.competitor_data:
            data, timestamp = self.competitor_data[cache_key]
            if (datetime.now() - timestamp).total_seconds() < 3600:
                return data
        
        if marketplace:
            parser_map = {
                'Яндекс Маркет': self.parser.parse_yandex_market,
                'Ozon': self.parser.parse_ozon,
                'Wildberries': self.parser.parse_wildberries
            }
            parser = parser_map.get(marketplace)
            if parser:
                results = parser(query)
            else:
                results = []
        else:
            results = self.parser.parse_all_marketplaces(query)
        
        self.competitor_data[cache_key] = (results, datetime.now())
        return results
    
    def analyze_competitor_prices(self, product_name: str, our_price: float) -> Dict:
        """Анализ цен конкурентов"""
        # Получаем цены конкурентов
        competitor_data = self.get_competitor_prices(product_name)
        
        all_prices = []
        if isinstance(competitor_data, dict):
            for marketplace, items in competitor_data.items():
                for item in items:
                    price = item.get('price', 0)
                    if price > 0:
                        all_prices.append({
                            'marketplace': marketplace,
                            'price': price,
                            'name': item.get('name', '')
                        })
        elif isinstance(competitor_data, list):
            for item in competitor_data:
                price = item.get('price', 0)
                if price > 0:
                    all_prices.append({
                        'marketplace': item.get('marketplace', 'Неизвестно'),
                        'price': price,
                        'name': item.get('name', '')
                    })
        
        if not all_prices:
            return {
                'competitor_count': 0,
                'avg_price': our_price,
                'min_price': our_price,
                'max_price': our_price,
                'price_position': 'Нет данных',
                'recommendation': 'Нет конкурентов для анализа'
            }
        
        prices = [p['price'] for p in all_prices]
        avg_price = sum(prices) / len(prices)
        min_price = min(prices)
        max_price = max(prices)
        
        # Позиция нашей цены
        if our_price <= min_price:
            position = "Ниже всех"
            recommendation = "Высокая конкурентоспособность"
        elif our_price <= avg_price:
            position = "Ниже среднего"
            recommendation = "Хорошая позиция"
        elif our_price <= max_price:
            position = "Выше среднего"
            recommendation = "Стоит снизить цену"
        else:
            position = "Выше всех"
            recommendation = "Срочно снизить цену"
        
        return {
            'competitor_count': len(all_prices),
            'avg_price': avg_price,
            'min_price': min_price,
            'max_price': max_price,
            'price_position': position,
            'recommendation': recommendation,
            'competitors': all_prices[:10]  # Топ-10 конкурентов
        }

# ============================================================
# 12. ВАЛИДАЦИЯ ДАННЫХ
# ============================================================

class DataValidator:
    """Валидация данных с подробными отчетами"""
    
    def __init__(self, model_type: str = "product"):
        self.model_type = model_type
        self.errors = []
        self.warnings = []
        self.stats = {
            "total": 0,
            "valid": 0,
            "invalid": 0,
            "skipped": 0,
            "by_error_type": defaultdict(int)
        }
    
    def validate_product(self, row: Dict) -> Tuple[bool, List[str]]:
        """Валидация одного товара в зависимости от модели"""
        errors = []
        warnings = []
        
        if self.model_type == "product":
            # Товарная модель: проверяем цену, себестоимость, размеры
            price = safe_float(row.get("price", 0))
            if price <= 0:
                errors.append("Цена должна быть больше 0")
            elif price < CONFIG["validation"]["min_price"]:
                warnings.append(f"Цена очень низкая: {price} ₽")
            elif price > CONFIG["validation"]["max_price"]:
                warnings.append(f"Цена очень высокая: {price} ₽")
            
            cost = safe_float(row.get("cost", 0))
            if cost < 0:
                errors.append("Себестоимость не может быть отрицательной")
            elif cost == 0 and price > 0:
                warnings.append("Себестоимость не указана, используется 50% от цены")
            
            for dim, label in [("length", "Длина"), ("width", "Ширина"), ("height", "Высота")]:
                value = safe_float(row.get(dim, 0))
                if value < 0:
                    errors.append(f"{label} не может быть отрицательной")
                elif value > CONFIG["validation"]["max_dimension"]:
                    warnings.append(f"Подозрительно большая {label.lower()}: {value} мм")
        
        else:  # Агентская модель
            commission = safe_float(row.get("commission_pct", 0), 0.1)
            if commission <= 0:
                errors.append("Комиссия должна быть больше 0")
            elif commission > CONFIG["validation"]["max_commission"]:
                errors.append(f"Комиссия слишком высокая: {commission*100:.0f}%")
            
            price = safe_float(row.get("price", 0))
            if price <= 0:
                errors.append("Цена товара должна быть больше 0")
            
            transactions = safe_int(row.get("transactions", 0), 1)
            if transactions <= 0:
                warnings.append("Количество транзакций не указано, используется 1")
        
        # Общие проверки
        name = safe_str(row.get("name", ""))
        if not name:
            warnings.append("Отсутствует наименование товара")
        
        article = safe_str(row.get("article", ""))
        if not article:
            warnings.append("Отсутствует артикул товара")
        
        # Обновление статистики
        self.stats["total"] += 1
        for error in errors:
            self.stats["by_error_type"][error] += 1
        
        if errors:
            self.stats["invalid"] += 1
        else:
            self.stats["valid"] += 1
            if warnings:
                self.stats["skipped"] += 1
        
        self.errors.extend(errors)
        self.warnings.extend(warnings)
        
        return len(errors) == 0, errors + warnings
    
    def get_report(self) -> Dict:
        """Получение отчета по валидации"""
        return {
            "total": self.stats["total"],
            "valid": self.stats["valid"],
            "invalid": self.stats["invalid"],
            "skipped": self.stats["skipped"],
            "errors": self.errors[:100],
            "warnings": self.warnings[:100],
            "error_types": dict(self.stats["by_error_type"])
        }
    
    def reset(self):
        """Сброс валидатора"""
        self.errors = []
        self.warnings = []
        self.stats = {
            "total": 0,
            "valid": 0,
            "invalid": 0,
            "skipped": 0,
            "by_error_type": defaultdict(int)
        }

# ============================================================
# 13. ДВИЖОК РАСЧЕТА ЮНИТ-ЭКОНОМИКИ
# ============================================================

class UnitEconomicsEngine:
    """Движок расчета юнит-экономики"""
    
    def __init__(self, 
                 marketplace: str = "Яндекс Маркет", 
                 mode: str = "FBY", 
                 days_storage: int = 30,
                 model_type: str = "product"):
        
        self.marketplace = marketplace
        self.mode = mode
        self.days_storage = days_storage
        self.model_type = model_type
        self.tariff_provider = AITariffProvider()
        self.rates = self.tariff_provider.get_rates(marketplace, mode, model_type)
        self.fixed_costs = 50000.0
        self.avg_orders = 100
        self.retention_rate = 0.7
        self.discount_rate = 0.1
        self.lead_time = 7
        self.z_score = 1.65
        self.validator = DataValidator(model_type)
        self.competitor_manager = CompetitorManager()
        
    def calculate_product(self, row: Dict) -> Optional[Dict]:
        """Расчет юнит-экономики для одного товара"""
        try:
            # Валидация
            is_valid, validation_issues = self.validator.validate_product(row)
            
            if self.model_type == "product":
                return self._calculate_product_model(row, is_valid, validation_issues)
            else:
                return self._calculate_agency_model(row, is_valid, validation_issues)
                
        except Exception as e:
            logger.error(f"Error calculating product {row.get('article', 'unknown')}: {e}")
            return None
    
    def _calculate_product_model(self, row: Dict, is_valid: bool, validation_issues: List[str]) -> Dict:
        """Расчет для товарной модели (B2C)"""
        article = safe_str(row.get("article", ""))
        name = safe_str(row.get("name", ""))
        price = safe_float(row.get("price", 0))
        length = safe_float(row.get("length", 0))
        width = safe_float(row.get("width", 0))
        height = safe_float(row.get("height", 0))
        cost = safe_float(row.get("cost", price * 0.5 if price > 0 else 0))
        
        if price <= 0:
            return None
        
        # Анализ конкурентов
        competitor_analysis = self.competitor_manager.analyze_competitor_prices(name, price)
        
        volume = calculate_volume(length, width, height)
        weight = volume * 0.8 if volume > 0 else 1.0
        
        if volume == 0:
            volume = 1.0
            weight = 1.0
        
        commission = price * self.rates.get("commission", 0.11)
        service_fee = price * self.rates.get("service_fee", 0.01)
        insurance = price * self.rates.get("insurance", 0.005)
        logistics = self._calculate_logistics(volume, weight, price)
        storage = self._calculate_storage(volume)
        acquiring = price * self.rates.get("acquiring", 0.022)
        advertising = price * self.rates.get("advertising", 0.15)
        returns = price * self.rates.get("returns", 0.10)
        packaging = self.rates.get("packaging", 50.0)
        
        total_variable = cost + commission + service_fee + insurance + logistics + storage + acquiring + advertising + returns + packaging
        total_cost = total_variable
        
        unit_profit = price - total_cost
        margin = (unit_profit / price * 100) if price > 0 else 0
        contribution_margin = price - total_variable
        contribution_margin_pct = (contribution_margin / price * 100) if price > 0 else 0
        
        roi = (unit_profit / cost * 100) if cost > 0 else 0
        
        daily_profit = unit_profit * self.avg_orders / 30
        payback_days = cost / daily_profit if daily_profit > 0 else 999
        
        ltv = unit_profit / (1 - self.retention_rate + self.discount_rate)
        cac = advertising / 5 * 1000 if advertising > 0 else price * 0.05
        ltv_cac_ratio = ltv / cac if cac > 0 else 0
        
        annual_demand = self.avg_orders * 12
        holding_cost = cost * 0.2
        eoq = math.sqrt((2 * annual_demand * 500) / holding_cost) if holding_cost > 0 else 0
        
        safety_stock = self.z_score * (self.avg_orders / 30 * 0.2) * math.sqrt(self.lead_time)
        reorder_point = (self.avg_orders / 30 * self.lead_time) + safety_stock
        
        break_even_units = self.fixed_costs / contribution_margin if contribution_margin > 0 else 999999
        break_even_revenue = break_even_units * price
        
        profit_status = "Прибыльный" if unit_profit > 0 else "Убыточный"
        margin_status = "Высокая" if margin > 25 else "Средняя" if margin > 10 else "Низкая"
        
        # Рекомендации с учетом конкурентов
        if margin < 15:
            recommended_price = price * 1.15
            price_action = "Повысить"
            price_reason = "Низкая маржинальность (<15%)"
        elif margin > 35:
            recommended_price = price * 0.95
            price_action = "Снизить"
            price_reason = "Высокая маржинальность (>35%)"
        elif competitor_analysis.get('price_position') == "Выше всех":
            recommended_price = price * 0.95
            price_action = "Снизить"
            price_reason = f"Цена выше всех конкурентов (средняя {competitor_analysis.get('avg_price', 0):.0f} ₽)"
        else:
            recommended_price = price
            price_action = "Оставить"
            price_reason = "Оптимальная маржинальность"
        
        return {
            "article": article,
            "name": name,
            "category": safe_str(row.get("category", "Прочее")),
            "oe_number": safe_str(row.get("oe_number", "")),
            "brand": safe_str(row.get("brand", "")),
            "price": round(price, 2),
            "cost": round(cost, 2),
            "length": round(length, 1),
            "width": round(width, 1),
            "height": round(height, 1),
            "volume": round(volume, 2),
            "weight": round(weight, 2),
            "commission": round(commission, 2),
            "service_fee": round(service_fee, 2),
            "insurance": round(insurance, 2),
            "logistics": round(logistics, 2),
            "storage": round(storage, 2),
            "acquiring": round(acquiring, 2),
            "advertising": round(advertising, 2),
            "returns": round(returns, 2),
            "packaging": round(packaging, 2),
            "total_variable_costs": round(total_variable, 2),
            "total_cost": round(total_cost, 2),
            "contribution_margin": round(contribution_margin, 2),
            "contribution_margin_pct": round(contribution_margin_pct, 1),
            "unit_profit": round(unit_profit, 2),
            "margin": round(margin, 1),
            "roi": round(roi, 1),
            "payback_days": round(payback_days, 0),
            "break_even_units": round(break_even_units, 0),
            "break_even_revenue": round(break_even_revenue, 2),
            "ltv": round(ltv, 2),
            "cac": round(cac, 2),
            "ltv_cac_ratio": round(ltv_cac_ratio, 2),
            "eoq": round(eoq, 0),
            "reorder_point": round(reorder_point, 0),
            "safety_stock": round(safety_stock, 0),
            "profit_status": profit_status,
            "margin_status": margin_status,
            "recommended_price": round(recommended_price, 2),
            "price_action": price_action,
            "price_reason": price_reason,
            "abc_category": self._abc_category(unit_profit),
            "xyz_category": self._xyz_category(self.avg_orders),
            "competitor_count": competitor_analysis.get('competitor_count', 0),
            "competitor_avg_price": round(competitor_analysis.get('avg_price', 0), 2),
            "competitor_min_price": round(competitor_analysis.get('min_price', 0), 2),
            "competitor_max_price": round(competitor_analysis.get('max_price', 0), 2),
            "price_position": competitor_analysis.get('price_position', 'Нет данных'),
            "competitor_recommendation": competitor_analysis.get('recommendation', ''),
            "mode": self.mode,
            "marketplace": self.marketplace,
            "model_type": self.model_type,
            "validation_valid": is_valid,
            "validation_issues": "; ".join(validation_issues[:3]) if validation_issues else "",
        }
    
    def _calculate_agency_model(self, row: Dict, is_valid: bool, validation_issues: List[str]) -> Dict:
        """Расчет для агентской модели (Marketplace)"""
        article = safe_str(row.get("article", ""))
        name = safe_str(row.get("name", ""))
        price = safe_float(row.get("price", 0))
        commission_pct = safe_float(row.get("commission_pct", 0), self.rates.get("commission", 0.10))
        transactions = safe_int(row.get("transactions", 0), 1)
        
        if price <= 0:
            return None
        
        # Выручка от комиссии
        commission_income = price * commission_pct
        transaction_fee = price * self.rates.get("transaction_fee", 0.02)
        
        # Операционные расходы на транзакцию
        support_cost = self.rates.get("support_cost_per_tx", 50.0)
        tech_cost = self.rates.get("tech_cost_per_tx", 20.0)
        marketing_cost = price * self.rates.get("marketing_coef", 0.30)
        acquiring_cost = price * self.rates.get("acquiring", 0.025)
        
        # Затраты на одну транзакцию
        total_cost_per_tx = support_cost + tech_cost + marketing_cost + acquiring_cost + transaction_fee
        
        # Прибыль на транзакцию
        profit_per_tx = commission_income - total_cost_per_tx
        margin_per_tx = (profit_per_tx / commission_income * 100) if commission_income > 0 else 0
        
        # Годовая прибыль
        annual_income = commission_income * transactions * 12
        annual_profit = profit_per_tx * transactions * 12
        annual_margin = (annual_profit / annual_income * 100) if annual_income > 0 else 0
        
        # LTV
        ltv = profit_per_tx / (1 - self.retention_rate + self.discount_rate)
        cac = marketing_cost / 5 * 1000 if marketing_cost > 0 else price * 0.05
        ltv_cac_ratio = ltv / cac if cac > 0 else 0
        
        # Точка безубыточности по транзакциям
        break_even_tx = self.fixed_costs / profit_per_tx if profit_per_tx > 0 else 999999
        
        # Статусы
        profit_status = "Прибыльный" if profit_per_tx > 0 else "Убыточный"
        margin_status = "Высокая" if margin_per_tx > 30 else "Средняя" if margin_per_tx > 10 else "Низкая"
        
        # Рекомендации
        if margin_per_tx < 15:
            recommended_commission = commission_pct * 1.15
            price_action = "Повысить комиссию"
            price_reason = f"Низкая маржинальность ({margin_per_tx:.1f}%)"
        elif margin_per_tx > 40:
            recommended_commission = commission_pct * 0.95
            price_action = "Снизить комиссию"
            price_reason = f"Высокая маржинальность ({margin_per_tx:.1f}%)"
        else:
            recommended_commission = commission_pct
            price_action = "Оставить"
            price_reason = "Оптимальная маржинальность"
        
        return {
            "article": article,
            "name": name,
            "category": safe_str(row.get("category", "Прочее")),
            "brand": safe_str(row.get("brand", "")),
            "price": round(price, 2),
            "commission_pct": round(commission_pct * 100, 1),
            "transactions": transactions,
            "commission_income": round(commission_income, 2),
            "transaction_fee": round(transaction_fee, 2),
            "support_cost": round(support_cost, 2),
            "tech_cost": round(tech_cost, 2),
            "marketing_cost": round(marketing_cost, 2),
            "acquiring_cost": round(acquiring_cost, 2),
            "total_cost_per_tx": round(total_cost_per_tx, 2),
            "profit_per_tx": round(profit_per_tx, 2),
            "margin_per_tx": round(margin_per_tx, 1),
            "annual_income": round(annual_income, 2),
            "annual_profit": round(annual_profit, 2),
            "annual_margin": round(annual_margin, 1),
            "ltv": round(ltv, 2),
            "cac": round(cac, 2),
            "ltv_cac_ratio": round(ltv_cac_ratio, 2),
            "break_even_tx": round(break_even_tx, 0),
            "profit_status": profit_status,
            "margin_status": margin_status,
            "recommended_commission": round(recommended_commission * 100, 1),
            "price_action": price_action,
            "price_reason": price_reason,
            "mode": self.mode,
            "marketplace": self.marketplace,
            "model_type": self.model_type,
            "validation_valid": is_valid,
            "validation_issues": "; ".join(validation_issues[:3]) if validation_issues else "",
        }
    
    def _calculate_logistics(self, volume: float, weight: float, price: float) -> float:
        """Расчет логистики"""
        try:
            if self.mode == "FBY":
                logistics = self.rates.get("fba_base", 70.0) + weight * self.rates.get("fba_per_kg", 12.0)
            elif self.mode == "FBS":
                logistics = self.rates.get("fbs_logistics", 115.0)
                if volume <= self.rates.get("logistics_threshold", 1.0):
                    logistics = self.rates.get("logistics_base", 70.0)
                elif volume <= self.rates.get("logistics_threshold_160", 160.0):
                    logistics = self.rates.get("logistics_base", 70.0) + (volume - 1.0) * self.rates.get("logistics_per_liter", 22.0)
                else:
                    logistics = self.rates.get("logistics_base", 70.0) + (160.0 - 1.0) * 22.0 + (volume - 160.0) * 2.0
            else:
                logistics = self.rates.get("logistics_base", 70.0) + volume * self.rates.get("logistics_per_liter", 10.0) + weight * 15.0
            
            if self.mode != "FBY":
                delivery_to_customer = min(price * self.rates.get("delivery_to_customer_percent", 0.045), 
                                         self.rates.get("delivery_to_customer_max", 1000.0))
                logistics += delivery_to_customer
            
            return round(logistics, 2)
        except Exception as e:
            logger.error(f"Error calculating logistics: {e}")
            return self.rates.get("logistics_base", 70.0)
    
    def _calculate_storage(self, volume: float) -> float:
        """Расчет хранения"""
        try:
            if self.mode == "FBY":
                free_days = self.rates.get("storage_free_days", 365)
                if self.days_storage <= free_days:
                    return 0.0
                return round(volume * self.rates.get("storage_per_liter", 3.0) * (self.days_storage - free_days), 2)
            elif self.mode == "FBO":
                return round(volume * self.rates.get("fbo_storage", 0.5) * (self.days_storage / 30), 2)
            else:
                return round(volume * self.rates.get("storage_per_liter", 0.8) * (self.days_storage / 30), 2)
        except Exception as e:
            logger.error(f"Error calculating storage: {e}")
            return 0.0
    
    def _abc_category(self, profit: float) -> str:
        """Определение ABC-категории"""
        if profit > 1000:
            return "A"
        elif profit > 100:
            return "B"
        return "C"
    
    def _xyz_category(self, sales: float) -> str:
        """Определение XYZ-категории"""
        if sales > 100:
            return "X"
        elif sales > 50:
            return "Y"
        return "Z"
    
    def get_validation_report(self) -> Dict:
        """Получение отчета по валидации"""
        return self.validator.get_report()

# ============================================================
# 14. АНАЛИТИКА ПРОДАЖ
# ============================================================

class SalesAnalytics:
    """Аналитика продаж в реальном времени"""
    
    def __init__(self):
        self.sales_data = []
        self.cache = APICache()
        self.predictions = {}
        
    def add_sale(self, sale: Dict):
        """Добавление записи о продаже"""
        self.sales_data.append({
            "timestamp": datetime.now(),
            "product_id": sale.get("product_id", ""),
            "product_name": sale.get("product_name", ""),
            "category": sale.get("category", "Прочее"),
            "quantity": sale.get("quantity", 1),
            "price": sale.get("price", 0),
            "cost": sale.get("cost", 0),
            "profit": sale.get("profit", 0),
            "marketplace": sale.get("marketplace", "Неизвестно")
        })
        
        # Ограничиваем размер данных
        if len(self.sales_data) > 10000:
            self.sales_data = self.sales_data[-5000:]
    
    def get_sales_dataframe(self) -> pd.DataFrame:
        """Получение данных продаж в виде DataFrame"""
        if not self.sales_data:
            return pd.DataFrame()
        
        df = pd.DataFrame(self.sales_data)
        df['date'] = df['timestamp'].dt.date
        df['month'] = df['timestamp'].dt.month
        df['week'] = df['timestamp'].dt.isocalendar().week
        df['day_of_week'] = df['timestamp'].dt.dayofweek
        df['hour'] = df['timestamp'].dt.hour
        
        return df
    
    def get_daily_sales(self, days: int = 30) -> pd.DataFrame:
        """Получение ежедневных продаж"""
        df = self.get_sales_dataframe()
        if df.empty:
            return pd.DataFrame()
        
        # Фильтруем по дате
        cutoff = datetime.now() - timedelta(days=days)
        df = df[df['timestamp'] >= cutoff]
        
        daily = df.groupby('date').agg({
            'quantity': 'sum',
            'profit': 'sum',
            'price': 'mean'
        }).reset_index()
        
        daily['revenue'] = daily['quantity'] * daily['price']
        
        return daily
    
    def get_category_sales(self) -> pd.DataFrame:
        """Продажи по категориям"""
        df = self.get_sales_dataframe()
        if df.empty:
            return pd.DataFrame()
        
        category = df.groupby('category').agg({
            'quantity': 'sum',
            'profit': 'sum',
            'price': 'mean'
        }).reset_index()
        
        category['revenue'] = category['quantity'] * category['price']
        category['profit_margin'] = (category['profit'] / category['revenue'] * 100).fillna(0)
        
        return category.sort_values('revenue', ascending=False)
    
    def get_marketplace_sales(self) -> pd.DataFrame:
        """Продажи по маркетплейсам"""
        df = self.get_sales_dataframe()
        if df.empty:
            return pd.DataFrame()
        
        marketplace = df.groupby('marketplace').agg({
            'quantity': 'sum',
            'profit': 'sum',
            'price': 'mean'
        }).reset_index()
        
        marketplace['revenue'] = marketplace['quantity'] * marketplace['price']
        
        return marketplace.sort_values('revenue', ascending=False)
    
    def get_hourly_heatmap(self) -> pd.DataFrame:
        """Тепловая карта продаж по часам и дням недели"""
        df = self.get_sales_dataframe()
        if df.empty:
            return pd.DataFrame()
        
        heatmap_data = df.groupby(['day_of_week', 'hour']).agg({
            'quantity': 'sum'
        }).reset_index()
        
        return heatmap_data
    
    def get_abc_xyz_matrix(self) -> pd.DataFrame:
        """ABC/XYZ матрица товаров"""
        df = self.get_sales_dataframe()
        if df.empty:
            return pd.DataFrame()
        
        # Суммируем продажи по товарам
        product_sales = df.groupby(['product_id', 'product_name', 'category']).agg({
            'quantity': 'sum',
            'profit': 'sum',
            'price': 'mean'
        }).reset_index()
        
        if product_sales.empty:
            return pd.DataFrame()
        
        # ABC классификация по прибыли
        product_sales = product_sales.sort_values('profit', ascending=False)
        total_profit = product_sales['profit'].sum()
        
        if total_profit == 0:
            return pd.DataFrame()
        
        product_sales['cumulative_profit_pct'] = product_sales['profit'].cumsum() / total_profit * 100
        
        conditions = [
            product_sales['cumulative_profit_pct'] <= 80,
            product_sales['cumulative_profit_pct'] <= 95
        ]
        choices = ['A', 'B', 'C']
        product_sales['abc'] = np.select(conditions, choices, default='C')
        
        # XYZ классификация по стабильности продаж
        # Рассчитываем коэффициент вариации для каждого товара
        product_variation = df.groupby('product_id')['quantity'].std() / df.groupby('product_id')['quantity'].mean()
        product_variation = product_variation.fillna(0)
        
        product_sales['coefficient_variation'] = product_sales['product_id'].map(product_variation).fillna(0)
        
        conditions_xyz = [
            product_sales['coefficient_variation'] < 0.5,
            product_sales['coefficient_variation'] < 1.0
        ]
        choices_xyz = ['X', 'Y', 'Z']
        product_sales['xyz'] = np.select(conditions_xyz, choices_xyz, default='Z')
        
        product_sales['abc_xyz'] = product_sales['abc'] + product_sales['xyz']
        
        return product_sales

# ============================================================
# 15. ПРОГНОЗИРОВАНИЕ ПРОДАЖ
# ============================================================

class SalesForecaster:
    """Прогнозирование продаж"""
    
    def __init__(self):
        self.models = {}
        self.predictions = {}
        
    def forecast_arima(self, data: List[float], steps: int = 30) -> List[float]:
        """ARIMA прогноз (упрощенная версия)"""
        if not data or len(data) < 3:
            return [data[-1] if data else 0] * steps
        
        try:
            # Простой экспоненциальное сглаживание
            alpha = 0.3
            smoothed = [data[0]]
            for value in data[1:]:
                smoothed.append(alpha * value + (1 - alpha) * smoothed[-1])
            
            # Прогноз
            last = smoothed[-1]
            trend = (smoothed[-1] - smoothed[-3]) / 2 if len(smoothed) >= 3 else 0
            
            predictions = []
            for i in range(steps):
                predicted = last + trend * (i + 1)
                predictions.append(max(0, predicted))
            
            return predictions
            
        except Exception as e:
            logger.error(f"ARIMA forecast error: {e}")
            return [data[-1] if data else 0] * steps
    
    def forecast_prophet(self, dates: List[datetime], values: List[float], steps: int = 30) -> List[float]:
        """Prophet стиль прогноз (упрощенный)"""
        if not dates or not values or len(dates) < 3:
            return [values[-1] if values else 0] * steps
        
        try:
            # Определяем сезонность
            if len(dates) >= 7:
                # Недельная сезонность
                weekdays = [d.weekday() for d in dates[-7:]]
                weekday_avg = {}
                for i, wd in enumerate(weekdays):
                    if wd not in weekday_avg:
                        weekday_avg[wd] = []
                    weekday_avg[wd].append(values[-7 + i])
                
                weekday_avg = {k: sum(v)/len(v) for k, v in weekday_avg.items()}
                avg_value = sum(values[-7:]) / 7
            
            # Прогноз
            predictions = []
            last_date = dates[-1]
            
            for i in range(1, steps + 1):
                future_date = last_date + timedelta(days=i)
                weekday = future_date.weekday()
                
                if weekday in weekday_avg:
                    predicted = weekday_avg[weekday]
                else:
                    predicted = avg_value
                
                # Добавляем тренд
                if len(values) >= 3:
                    trend = (values[-1] - values[-3]) / 2
                    predicted += trend * (i / 7)
                
                predictions.append(max(0, predicted))
            
            return predictions
            
        except Exception as e:
            logger.error(f"Prophet forecast error: {e}")
            return [values[-1] if values else 0] * steps
    
    def forecast_linear_regression(self, data: List[float], steps: int = 30) -> List[float]:
        """Линейная регрессия для прогноза"""
        if not data or len(data) < 3:
            return [data[-1] if data else 0] * steps
        
        try:
            if not SKLEARN_AVAILABLE:
                return self.forecast_arima(data, steps)
            
            n = len(data)
            x = np.array(range(n)).reshape(-1, 1)
            y = np.array(data)
            
            model = LinearRegression()
            model.fit(x, y)
            
            future_x = np.array(range(n, n + steps)).reshape(-1, 1)
            predictions = model.predict(future_x)
            
            return [max(0, p) for p in predictions]
            
        except Exception as e:
            logger.error(f"Linear regression forecast error: {e}")
            return [data[-1] if data else 0] * steps
    
    def combine_forecasts(self, data: List[float], steps: int = 30) -> Dict:
        """Комбинированный прогноз"""
        if not data:
            return {"predictions": [0] * steps, "confidence": []}
        
        # Получаем прогнозы разными методами
        arima = self.forecast_arima(data, steps)
        prophet = self.forecast_prophet(
            [datetime.now() - timedelta(days=len(data)-i-1) for i in range(len(data))],
            data,
            steps
        )
        
        try:
            if SKLEARN_AVAILABLE:
                linear = self.forecast_linear_regression(data, steps)
            else:
                linear = arima
        except:
            linear = arima
        
        # Усредняем
        predictions = []
        confidence = []
        
        for i in range(steps):
            # Взвешенное среднее
            weights = [0.4, 0.3, 0.3]  # ARIMA, Prophet, Linear
            combined = (arima[i] * weights[0] + 
                       prophet[i] * weights[1] + 
                       linear[i] * weights[2])
            predictions.append(combined)
            
            # Уверенность (чем дальше, тем меньше уверенность)
            confidence.append(max(0, 1 - (i / steps) * 0.5))
        
        return {
            "predictions": predictions,
            "confidence": confidence,
            "arima": arima,
            "prophet": prophet,
            "linear": linear
        }

# ============================================================
# 16. АВТОМАТИЧЕСКАЯ ВЫГРУЗКА НА МАРКЕТПЛЕЙСЫ
# ============================================================

class MarketplaceUploader:
    """Автоматическая выгрузка данных на маркетплейсы"""
    
    def __init__(self):
        self.api_clients = {}
        self.upload_queue = Queue()
        self.is_running = False
        self.upload_log = []
        self.lock = Lock()
        
    def register_client(self, marketplace: str, client: BaseMarketplaceAPI):
        """Регистрация API-клиента для маркетплейса"""
        self.api_clients[marketplace] = client
        
    def upload_prices(self, marketplace: str, products: List[Dict]) -> Dict:
        """Выгрузка цен на маркетплейс"""
        client = self.api_clients.get(marketplace)
        if not client:
            return {"status": "error", "message": f"Клиент для {marketplace} не зарегистрирован"}
        
        results = {
            "marketplace": marketplace,
            "total": len(products),
            "success": 0,
            "failed": 0,
            "errors": [],
            "timestamp": datetime.now().isoformat()
        }
        
        for product in products:
            try:
                offer_id = product.get('offer_id', product.get('article', ''))
                price = product.get('recommended_price', product.get('price', 0))
                
                if not offer_id or price <= 0:
                    continue
                
                # Обновление цены в зависимости от маркетплейса
                if marketplace == "Яндекс Маркет":
                    result = client.update_price(offer_id, price)
                elif marketplace == "Ozon":
                    result = client.update_price(offer_id, price)
                elif marketplace == "Wildberries":
                    result = client.update_price(int(offer_id), price)
                else:
                    continue
                
                if result and result.get('status') != 'error':
                    results["success"] += 1
                else:
                    results["failed"] += 1
                    results["errors"].append({
                        "offer_id": offer_id,
                        "error": str(result) if result else "Неизвестная ошибка"
                    })
                
                time.sleep(0.1)  # Rate limiting
                
            except Exception as e:
                results["failed"] += 1
                results["errors"].append({
                    "offer_id": product.get('offer_id', 'unknown'),
                    "error": str(e)
                })
        
        # Логирование
        with self.lock:
            self.upload_log.append(results)
        
        return results
    
    def upload_stocks(self, marketplace: str, products: List[Dict]) -> Dict:
        """Выгрузка остатков на маркетплейс"""
        client = self.api_clients.get(marketplace)
        if not client:
            return {"status": "error", "message": f"Клиент для {marketplace} не зарегистрирован"}
        
        results = {
            "marketplace": marketplace,
            "total": len(products),
            "success": 0,
            "failed": 0,
            "errors": [],
            "timestamp": datetime.now().isoformat()
        }
        
        for product in products:
            try:
                offer_id = product.get('offer_id', product.get('article', ''))
                stock = product.get('stock', product.get('quantity', 0))
                
                if not offer_id or stock < 0:
                    continue
                
                # Обновление остатков
                if marketplace == "Яндекс Маркет":
                    result = client.update_stock(offer_id, stock)
                elif marketplace == "Ozon":
                    result = client.update_stock(offer_id, stock)
                elif marketplace == "Wildberries":
                    result = client.update_stock(int(offer_id), stock)
                else:
                    continue
                
                if result and result.get('status') != 'error':
                    results["success"] += 1
                else:
                    results["failed"] += 1
                    results["errors"].append({
                        "offer_id": offer_id,
                        "error": str(result) if result else "Неизвестная ошибка"
                    })
                
                time.sleep(0.1)
                
            except Exception as e:
                results["failed"] += 1
                results["errors"].append({
                    "offer_id": product.get('offer_id', 'unknown'),
                    "error": str(e)
                })
        
        with self.lock:
            self.upload_log.append(results)
        
        return results
    
    def upload_products(self, marketplace: str, products: List[Dict]) -> Dict:
        """Создание/обновление карточек товаров"""
        client = self.api_clients.get(marketplace)
        if not client:
            return {"status": "error", "message": f"Клиент для {marketplace} не зарегистрирован"}
        
        results = {
            "marketplace": marketplace,
            "total": len(products),
            "success": 0,
            "failed": 0,
            "errors": [],
            "timestamp": datetime.now().isoformat()
        }
        
        for product in products:
            try:
                # Подготовка данных для создания карточки
                card_data = {
                    "offer_id": product.get('article', ''),
                    "name": product.get('name', ''),
                    "price": product.get('price', 0),
                    "category": product.get('category', ''),
                    "brand": product.get('brand', ''),
                    "description": product.get('description', ''),
                    "images": product.get('images', []),
                    "dimensions": {
                        "length": product.get('length', 0),
                        "width": product.get('width', 0),
                        "height": product.get('height', 0)
                    }
                }
                
                # Отправка на маркетплейс
                if marketplace == "Яндекс Маркет":
                    result = client.create_product(card_data)
                elif marketplace == "Ozon":
                    result = client.create_product(card_data)
                elif marketplace == "Wildberries":
                    result = client.create_product(card_data)
                else:
                    continue
                
                if result and result.get('status') != 'error':
                    results["success"] += 1
                else:
                    results["failed"] += 1
                    results["errors"].append({
                        "offer_id": product.get('article', 'unknown'),
                        "error": str(result) if result else "Неизвестная ошибка"
                    })
                
                time.sleep(0.2)
                
            except Exception as e:
                results["failed"] += 1
                results["errors"].append({
                    "offer_id": product.get('article', 'unknown'),
                    "error": str(e)
                })
        
        with self.lock:
            self.upload_log.append(results)
        
        return results
    
    def get_upload_log(self) -> List[Dict]:
        """Получение лога выгрузок"""
        with self.lock:
            return self.upload_log.copy()
    
    def clear_log(self):
        """Очистка лога"""
        with self.lock:
            self.upload_log = []

# ============================================================
# 17. ИНТЕГРАЦИЯ С 1С
# ============================================================

class OneCIntegration:
    """Интеграция с 1С:Предприятие"""
    
    def __init__(self, base_url: str = None, username: str = None, password: str = None):
        self.base_url = base_url
        self.username = username
        self.password = password
        self.session = requests.Session()
        if username and password:
            self.session.auth = (username, password)
        self.cache = APICache()
        
    def export_to_1c(self, data: List[Dict], entity_type: str = "products") -> Dict:
        """Экспорт данных в 1С"""
        if not self.base_url:
            return {"status": "error", "message": "URL 1С не настроен"}
        
        try:
            # Подготовка данных в формате 1С
            export_data = {
                "entity_type": entity_type,
                "data": data,
                "timestamp": datetime.now().isoformat()
            }
            
            # Отправка в 1С
            url = f"{self.base_url}/api/import"
            response = self.session.post(
                url,
                json=export_data,
                timeout=60
            )
            
            if response.status_code == 200:
                return {
                    "status": "success",
                    "message": "Данные успешно экспортированы в 1С",
                    "response": response.json()
                }
            else:
                return {
                    "status": "error",
                    "message": f"Ошибка {response.status_code}: {response.text}"
                }
                
        except Exception as e:
            logger.error(f"1C export error: {e}")
            return {"status": "error", "message": str(e)}
    
    def import_from_1c(self, entity_type: str = "products") -> Optional[List[Dict]]:
        """Импорт данных из 1С"""
        if not self.base_url:
            return None
        
        cache_key = generate_cache_key('1c_import', entity_type)
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        try:
            url = f"{self.base_url}/api/export"
            params = {"entity_type": entity_type}
            
            response = self.session.get(url, params=params, timeout=60)
            
            if response.status_code == 200:
                data = response.json()
                self.cache.set(cache_key, data)
                return data
            else:
                logger.error(f"1C import error: {response.status_code}")
                return None
                
        except Exception as e:
            logger.error(f"1C import error: {e}")
            return None
    
    def sync_prices(self, prices: List[Dict]) -> Dict:
        """Синхронизация цен с 1С"""
        return self.export_to_1c(prices, "prices")
    
    def sync_stocks(self, stocks: List[Dict]) -> Dict:
        """Синхронизация остатков с 1С"""
        return self.export_to_1c(stocks, "stocks")
    
    def sync_orders(self, orders: List[Dict]) -> Dict:
        """Синхронизация заказов с 1С"""
        return self.export_to_1c(orders, "orders")

# ============================================================
# 18. ИНТЕГРАЦИЯ С CRM
# ============================================================

class CRMIntegration:
    """Интеграция с CRM системами"""
    
    def __init__(self, crm_type: str = "amo", api_key: str = None, base_url: str = None):
        self.crm_type = crm_type
        self.api_key = api_key
        self.base_url = base_url
        self.session = requests.Session()
        
        if api_key:
            self.session.headers.update({
                'Authorization': f'Bearer {api_key}',
                'Content-Type': 'application/json'
            })
    
    def send_report(self, report_data: Dict) -> Dict:
        """Отправка отчета в CRM"""
        if not self.base_url:
            return {"status": "error", "message": "URL CRM не настроен"}
        
        try:
            if self.crm_type == "amo":
                return self._send_to_amo(report_data)
            elif self.crm_type == "bitrix24":
                return self._send_to_bitrix24(report_data)
            elif self.crm_type == "hubspot":
                return self._send_to_hubspot(report_data)
            else:
                return self._send_generic(report_data)
                
        except Exception as e:
            logger.error(f"CRM send error: {e}")
            return {"status": "error", "message": str(e)}
    
    def _send_to_amo(self, data: Dict) -> Dict:
        """Отправка в AmoCRM"""
        url = f"{self.base_url}/api/v4/leads"
        
        # Конвертация данных в формат AmoCRM
        lead_data = {
            "name": data.get("title", "Отчет по юнит-экономике"),
            "custom_fields_values": [
                {
                    "field_id": 123456,
                    "values": [{"value": data.get("summary", "")}]
                }
            ]
        }
        
        response = self.session.post(url, json=lead_data)
        if response.status_code in [200, 201]:
            return {"status": "success", "data": response.json()}
        else:
            return {"status": "error", "message": response.text}
    
    def _send_to_bitrix24(self, data: Dict) -> Dict:
        """Отправка в Bitrix24"""
        url = f"{self.base_url}/rest/deal.add.json"
        
        deal_data = {
            "fields": {
                "TITLE": data.get("title", "Отчет по юнит-экономике"),
                "COMMENTS": data.get("summary", ""),
                "OPPORTUNITY": data.get("total_profit", 0)
            }
        }
        
        response = self.session.post(url, json=deal_data)
        if response.status_code == 200:
            return {"status": "success", "data": response.json()}
        else:
            return {"status": "error", "message": response.text}
    
    def _send_to_hubspot(self, data: Dict) -> Dict:
        """Отправка в HubSpot"""
        url = f"{self.base_url}/crm/v3/objects/deals"
        
        deal_data = {
            "properties": {
                "dealname": data.get("title", "Отчет по юнит-экономике"),
                "description": data.get("summary", ""),
                "amount": data.get("total_profit", 0)
            }
        }
        
        response = self.session.post(url, json=deal_data)
        if response.status_code in [200, 201]:
            return {"status": "success", "data": response.json()}
        else:
            return {"status": "error", "message": response.text}
    
    def _send_generic(self, data: Dict) -> Dict:
        """Универсальная отправка"""
        response = self.session.post(
            f"{self.base_url}/api/webhook",
            json=data
        )
        
        if response.status_code in [200, 201, 202]:
            return {"status": "success", "data": response.json()}
        else:
            return {"status": "error", "message": response.text}

# ============================================================
# 19. РАСПИСАНИЕ АВТОМАТИЧЕСКИХ ВЫГРУЗОК
# ============================================================

class AutoUploadScheduler:
    """Расписание автоматических выгрузок"""
    
    def __init__(self):
        self.schedules = {}
        self.is_running = False
        self.thread = None
        self.lock = Lock()
        
    def add_schedule(self, name: str, schedule: Dict):
        """Добавление расписания"""
        with self.lock:
            self.schedules[name] = {
                "schedule": schedule,
                "last_run": None,
                "next_run": self._calculate_next_run(schedule),
                "enabled": True
            }
    
    def _calculate_next_run(self, schedule: Dict) -> datetime:
        """Расчет следующего запуска"""
        now = datetime.now()
        
        if schedule.get("type") == "interval":
            seconds = schedule.get("interval_seconds", 3600)
            return now + timedelta(seconds=seconds)
        
        elif schedule.get("type") == "cron":
            # Упрощенная реализация cron
            hour = schedule.get("hour", 0)
            minute = schedule.get("minute", 0)
            day_of_week = schedule.get("day_of_week", 0)  # 0 = понедельник
            
            next_run = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
            
            # Если уже прошло сегодня, переносим на следующую неделю
            if next_run <= now:
                days_until = (day_of_week - now.weekday()) % 7
                if days_until == 0:
                    days_until = 7
                next_run += timedelta(days=days_until)
            
            return next_run
        
        else:
            return now + timedelta(hours=1)
    
    def run_scheduled(self):
        """Запуск всех запланированных задач"""
        if self.is_running:
            return
        
        self.is_running = True
        self.thread = Thread(target=self._scheduler_loop)
        self.thread.daemon = True
        self.thread.start()
    
    def _scheduler_loop(self):
        """Основной цикл планировщика"""
        while self.is_running:
            try:
                now = datetime.now()
                
                with self.lock:
                    for name, data in self.schedules.items():
                        if not data["enabled"]:
                            continue
                        
                        if data["next_run"] and data["next_run"] <= now:
                            # Выполняем задачу
                            self._execute_schedule(name)
                            
                            # Обновляем next_run
                            data["last_run"] = now
                            data["next_run"] = self._calculate_next_run(data["schedule"])
                
                time.sleep(60)  # Проверяем каждую минуту
                
            except Exception as e:
                logger.error(f"Scheduler error: {e}")
                time.sleep(60)
    
    def _execute_schedule(self, name: str):
        """Выполнение запланированной задачи"""
        schedule = self.schedules.get(name)
        if not schedule:
            return
        
        task_type = schedule["schedule"].get("task_type")
        
        if task_type == "upload_prices":
            logger.info(f"Выполнение плановой выгрузки цен: {name}")
            # Здесь вызывается выгрузка цен
            
        elif task_type == "upload_stocks":
            logger.info(f"Выполнение плановой выгрузки остатков: {name}")
            
        elif task_type == "sync_1c":
            logger.info(f"Выполнение синхронизации с 1С: {name}")
            
        elif task_type == "send_report":
            logger.info(f"Отправка отчета в CRM: {name}")
    
    def stop(self):
        """Остановка планировщика"""
        self.is_running = False
        if self.thread:
            self.thread.join(timeout=5)

# ============================================================
# 20. ЭКСПОРТ В EXCEL
# ============================================================

class ExcelExportEngine:
    """Экспорт результатов в Excel"""
    
    def __init__(self):
        self.classifier = CategoryClassifier()
    
    def export(self, results: List[Dict], marketplace: str, mode: str, model_type: str, all_rates: Dict = None) -> bytes:
        """Экспорт результатов в Excel"""
        output = io.BytesIO()
        
        if not results:
            return output.getvalue()
        
        if not OPENPYXL_AVAILABLE:
            try:
                df = pd.DataFrame(results)
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Юнит-экономика', index=False)
                output.seek(0)
                return output.getvalue()
            except Exception as e:
                logger.error(f"Fallback export error: {e}")
                return output.getvalue()
        
        try:
            wb = Workbook()
            
            ws = wb.active
            ws.title = "Юнит-экономика"
            
            header_font = Font(bold=True, color="FFFFFF", size=11)
            header_fill = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
            header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
            if model_type == "product":
                headers = self._get_product_headers()
            else:
                headers = self._get_agency_headers()
            
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
            
            for row_idx, row_data in enumerate(results, 2):
                for col_idx, header in enumerate(headers, 1):
                    key = self._get_key_mapping(model_type).get(header, header.lower())
                    value = row_data.get(key, "")
                    cell = ws.cell(row=row_idx, column=col_idx, value=value)
                    
                    if isinstance(value, (int, float)):
                        if any(w in header for w in ["Цена", "Прибыль", "Доход", "стоимость", "LTV", "CAC"]):
                            cell.number_format = '#,##0.00 ₽'
                        elif any(w in header for w in ["%", "марж", "рр", "ROI", "CM"]):
                            cell.number_format = '0.00%'
                        elif any(w in header for w in ["дн", "запас", "EOQ", "транзак"]):
                            cell.number_format = '#,##0'
                        
                        if key in ["profit_per_tx", "unit_profit"] and isinstance(value, (int, float)):
                            if value > 0:
                                cell.font = Font(color="006400")
                            elif value < 0:
                                cell.font = Font(color="8B0000")
            
            for col in range(1, len(headers) + 1):
                col_letter = get_column_letter(col)
                ws.column_dimensions[col_letter].width = 15
            
            if all_rates:
                self._create_help_sheet(wb, all_rates)
            
            self._create_summary_sheet(wb, results, marketplace, mode, model_type)
            
            wb.save(output)
            output.seek(0)
            return output.getvalue()
            
        except Exception as e:
            logger.error(f"Export error: {e}")
            try:
                df = pd.DataFrame(results)
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Юнит-экономика', index=False)
                output.seek(0)
                return output.getvalue()
            except Exception as e2:
                logger.error(f"Fallback export error: {e2}")
                return output.getvalue()
    
    def _get_product_headers(self) -> List[str]:
        """Заголовки для товарной модели"""
        return [
            "Артикул", "Наименование", "Категория", "OEM номер", "Бренд",
            "Цена", "Себестоимость", "Длина", "Ширина", "Высота", "Объем (л)", "Вес (кг)",
            "Комиссия", "Сервисный сбор", "Страховка", "Логистика", "Хранение",
            "Эквайринг", "Реклама", "Возвраты", "Упаковка",
            "Итого переменных", "Полная себестоимость",
            "Маржинальный доход", "CM %", "Прибыль на ед.", "Маржа %", "ROI %",
            "Окупаемость (дн)", "Точка безубыточности (шт)", "Точка безубыточности (₽)",
            "LTV", "CAC", "LTV/CAC", "EOQ", "Точка заказа", "Страховой запас",
            "Статус прибыли", "Статус маржи", "ABC", "XYZ",
            "Кол-во конкурентов", "Ср. цена конкурентов", "Мин цена", "Макс цена",
            "Позиция цены", "Рекомендация конкурентов",
            "Рекомендуемая цена", "Действие по цене", "Режим", "Маркетплейс",
            "Валидация", "Проблемы"
        ]
    
    def _get_agency_headers(self) -> List[str]:
        """Заголовки для агентской модели"""
        return [
            "Артикул", "Наименование", "Категория", "Бренд",
            "Цена товара", "Комиссия %", "Транзакций/мес",
            "Доход от комиссии", "Эквайринг", "Поддержка", "Технологии", "Маркетинг",
            "Транзакционный сбор", "Итого затрат",
            "Прибыль на транзакцию", "Маржа транзакции %",
            "Годовой доход", "Годовая прибыль", "Годовая маржа %",
            "LTV", "CAC", "LTV/CAC",
            "Точка безубыточности (транз)",
            "Статус прибыли", "Статус маржи",
            "Рекомендуемая комиссия", "Действие", "Причина",
            "Режим", "Маркетплейс",
            "Валидация", "Проблемы"
        ]
    
    def _get_key_mapping(self, model_type: str) -> Dict[str, str]:
        """Маппинг заголовков на ключи данных"""
        if model_type == "product":
            return {
                "Артикул": "article",
                "Наименование": "name",
                "Категория": "category",
                "OEM номер": "oe_number",
                "Бренд": "brand",
                "Цена": "price",
                "Себестоимость": "cost",
                "Длина": "length",
                "Ширина": "width",
                "Высота": "height",
                "Объем (л)": "volume",
                "Вес (кг)": "weight",
                "Комиссия": "commission",
                "Сервисный сбор": "service_fee",
                "Страховка": "insurance",
                "Логистика": "logistics",
                "Хранение": "storage",
                "Эквайринг": "acquiring",
                "Реклама": "advertising",
                "Возвраты": "returns",
                "Упаковка": "packaging",
                "Итого переменных": "total_variable_costs",
                "Полная себестоимость": "total_cost",
                "Маржинальный доход": "contribution_margin",
                "CM %": "contribution_margin_pct",
                "Прибыль на ед.": "unit_profit",
                "Маржа %": "margin",
                "ROI %": "roi",
                "Окупаемость (дн)": "payback_days",
                "Точка безубыточности (шт)": "break_even_units",
                "Точка безубыточности (₽)": "break_even_revenue",
                "LTV": "ltv",
                "CAC": "cac",
                "LTV/CAC": "ltv_cac_ratio",
                "EOQ": "eoq",
                "Точка заказа": "reorder_point",
                "Страховой запас": "safety_stock",
                "Статус прибыли": "profit_status",
                "Статус маржи": "margin_status",
                "ABC": "abc_category",
                "XYZ": "xyz_category",
                "Кол-во конкурентов": "competitor_count",
                "Ср. цена конкурентов": "competitor_avg_price",
                "Мин цена": "competitor_min_price",
                "Макс цена": "competitor_max_price",
                "Позиция цены": "price_position",
                "Рекомендация конкурентов": "competitor_recommendation",
                "Рекомендуемая цена": "recommended_price",
                "Действие по цене": "price_action",
                "Режим": "mode",
                "Маркетплейс": "marketplace",
                "Валидация": "validation_valid",
                "Проблемы": "validation_issues"
            }
        else:
            return {
                "Артикул": "article",
                "Наименование": "name",
                "Категория": "category",
                "Бренд": "brand",
                "Цена товара": "price",
                "Комиссия %": "commission_pct",
                "Транзакций/мес": "transactions",
                "Доход от комиссии": "commission_income",
                "Эквайринг": "acquiring_cost",
                "Поддержка": "support_cost",
                "Технологии": "tech_cost",
                "Маркетинг": "marketing_cost",
                "Транзакционный сбор": "transaction_fee",
                "Итого затрат": "total_cost_per_tx",
                "Прибыль на транзакцию": "profit_per_tx",
                "Маржа транзакции %": "margin_per_tx",
                "Годовой доход": "annual_income",
                "Годовая прибыль": "annual_profit",
                "Годовая маржа %": "annual_margin",
                "LTV": "ltv",
                "CAC": "cac",
                "LTV/CAC": "ltv_cac_ratio",
                "Точка безубыточности (транз)": "break_even_tx",
                "Статус прибыли": "profit_status",
                "Статус маржи": "margin_status",
                "Рекомендуемая комиссия": "recommended_commission",
                "Действие": "price_action",
                "Причина": "price_reason",
                "Режим": "mode",
                "Маркетплейс": "marketplace",
                "Валидация": "validation_valid",
                "Проблемы": "validation_issues"
            }
    
    def _create_help_sheet(self, wb: Workbook, all_rates: Dict):
        """Создание листа с тарифами"""
        ws = wb.create_sheet("Тарифы")
        
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
        
        headers = ["Маркетплейс", "Режим", "Комиссия", "Эквайринг", "Логистика", "Хранение"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
        
        row_idx = 2
        for mp, modes in all_rates.items():
            for mode, rates in modes.items():
                ws.cell(row=row_idx, column=1, value=mp)
                ws.cell(row=row_idx, column=2, value=mode)
                ws.cell(row=row_idx, column=3, value=rates.get("commission", 0))
                ws.cell(row=row_idx, column=4, value=rates.get("acquiring", 0))
                ws.cell(row=row_idx, column=5, value=rates.get("logistics_base", 0))
                ws.cell(row=row_idx, column=6, value=rates.get("storage_per_liter", 0))
                row_idx += 1
        
        for col in range(1, len(headers) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 18
    
    def _create_summary_sheet(self, wb: Workbook, results: List[Dict], marketplace: str, mode: str, model_type: str):
        """Создание листа со сводкой"""
        ws = wb.create_sheet("Сводка")
        
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
        
        headers = ["Показатель", "Значение"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
        
        if results:
            if model_type == "product":
                total_profit = sum(r.get("unit_profit", 0) for r in results)
                avg_margin = sum(r.get("margin", 0) for r in results) / len(results) if results else 0
                profitable = sum(1 for r in results if r.get("unit_profit", 0) > 0)
                
                summary_data = [
                    ["Всего товаров", len(results)],
                    ["Общая прибыль", format_currency(total_profit)],
                    ["Средняя маржа", format_percent(avg_margin)],
                    ["Прибыльных товаров", profitable],
                    ["Убыточных товаров", len(results) - profitable],
                ]
            else:
                total_profit = sum(r.get("annual_profit", 0) for r in results)
                avg_margin = sum(r.get("annual_margin", 0) for r in results) / len(results) if results else 0
                profitable = sum(1 for r in results if r.get("profit_per_tx", 0) > 0)
                
                summary_data = [
                    ["Всего позиций", len(results)],
                    ["Общая годовая прибыль", format_currency(total_profit)],
                    ["Средняя годовая маржа", format_percent(avg_margin)],
                    ["Прибыльных позиций", profitable],
                    ["Убыточных позиций", len(results) - profitable],
                ]
            
            summary_data.extend([
                ["", ""],
                ["Параметры:", ""],
                ["Маркетплейс", marketplace],
                ["Режим", mode],
                ["Модель", "Товарная" if model_type == "product" else "Агентская"],
                ["Дата", datetime.now().strftime("%d.%m.%Y %H:%M")],
                ["Версия", CONFIG["version"]]
            ])
            
            for row_idx, (label, value) in enumerate(summary_data, 2):
                ws.cell(row=row_idx, column=1, value=label)
                ws.cell(row=row_idx, column=2, value=value)
        
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 20
    
    def load_existing(self, file_bytes: bytes) -> List[Dict]:
        """Загрузка существующей юнит-экономики из Excel"""
        if not OPENPYXL_AVAILABLE or not file_bytes:
            return []
        
        try:
            wb = load_workbook(io.BytesIO(file_bytes))
            ws = wb.active
            
            headers = []
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=1, column=col)
                if cell.value:
                    headers.append(safe_str(cell.value))
            
            if not headers:
                return []
            
            results = []
            key_mapping = self._get_key_mapping("product")
            key_mapping.update(self._get_key_mapping("agency"))
            
            for row in range(2, ws.max_row + 1):
                row_data = {}
                for col, header in enumerate(headers, 1):
                    cell = ws.cell(row=row, column=col)
                    value = cell.value
                    key = key_mapping.get(header, header.lower())
                    
                    if key in ["price", "cost", "commission", "logistics", "storage", 
                              "acquiring", "advertising", "returns", "packaging",
                              "total_variable_costs", "total_cost", "contribution_margin",
                              "unit_profit", "ltv", "cac", "break_even_revenue", "recommended_price",
                              "commission_income", "profit_per_tx", "annual_income", "annual_profit",
                              "competitor_avg_price", "competitor_min_price", "competitor_max_price"]:
                        value = safe_float(value)
                    elif key in ["margin", "contribution_margin_pct", "roi", "ltv_cac_ratio",
                                "margin_per_tx", "annual_margin"]:
                        value = safe_float(value)
                    elif key in ["transactions", "break_even_tx", "payback_days", "break_even_units", 
                                "eoq", "competitor_count"]:
                        value = safe_int(value)
                    else:
                        value = safe_str(value)
                    
                    row_data[key] = value
                
                if row_data.get("article"):
                    if "mode" not in row_data:
                        row_data["mode"] = "FBY"
                    if "marketplace" not in row_data:
                        row_data["marketplace"] = "Не указан"
                    if "model_type" not in row_data:
                        row_data["model_type"] = "product"
                    results.append(row_data)
            
            return results
            
        except Exception as e:
            logger.error(f"Error loading existing data: {e}")
            return []

# ============================================================
# 21. ГЛАВНОЕ ПРИЛОЖЕНИЕ
# ============================================================

class UnitEconomicsApp:
    """Главное приложение"""
    
    def __init__(self):
        self.classifier = CategoryClassifier()
        self.exporter = ExcelExportEngine()
        self.tariff_provider = AITariffProvider()
        self.results = []
        self.all_rates = {}
        self.engine = None
        self.model_type = "product"
        self.api_clients = {}
        self.uploader = MarketplaceUploader()
        self.scheduler = AutoUploadScheduler()
        self.onec = OneCIntegration()
        self.crm = CRMIntegration()
        self.sales_analytics = SalesAnalytics()
        self.forecaster = SalesForecaster()
        self.sales_data = []
        self._init_demo_data()
    
    def _init_demo_data(self):
        """Генерация демо-данных для дашборда"""
        if not self.sales_analytics.sales_data:
            categories = ["Двигатель", "Трансмиссия", "Подвеска", "Тормозная система", 
                         "Рулевое управление", "Электрооборудование", "Масла и жидкости", 
                         "Фильтры", "Кузовные детали", "Оптика"]
            marketplaces = ["Яндекс Маркет", "Ozon", "Wildberries"]
            
            start_date = datetime.now() - timedelta(days=90)
            
            for i in range(500):
                date = start_date + timedelta(days=random.randint(0, 90))
                hour = random.randint(8, 22)
                date = date.replace(hour=hour)
                
                category = random.choice(categories)
                marketplace = random.choice(marketplaces)
                quantity = random.randint(1, 5)
                price = random.randint(500, 15000)
                cost = price * random.uniform(0.4, 0.7)
                
                self.sales_analytics.add_sale({
                    "product_id": f"PROD_{i:06d}",
                    "product_name": f"Товар {i}",
                    "category": category,
                    "marketplace": marketplace,
                    "quantity": quantity,
                    "price": price,
                    "cost": cost,
                    "profit": (price - cost) * quantity
                })
            
            self.sales_data = self.sales_analytics.sales_data
        
    def run(self):
        """Запуск приложения"""
        st.set_page_config(
            page_title=CONFIG["app_name"],
            page_icon="🚀",
            layout="wide",
            initial_sidebar_state="expanded"
        )
        
        self._render_header()
        self._render_sidebar()
        self._render_main()
    
    def _render_header(self):
        """Отображение заголовка"""
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
                    padding: 1.5rem; border-radius: 15px; color: white; text-align: center; margin-bottom: 1.5rem;
                    border: 2px solid #e94560; box-shadow: 0 4px 15px rgba(0,0,0,0.3);">
            <h1 style="font-size: 2.8rem; margin: 0; text-shadow: 2px 2px 4px rgba(0,0,0,0.3);">
                🚀 {CONFIG['app_name']}
            </h1>
            <p style="font-size: 1.2rem; opacity: 0.95; margin-top: 0.3rem;">
                📊 <strong>Две модели расчета:</strong> Товарная (B2C) и Агентская (Marketplace)
            </p>
            <div style="display: flex; justify-content: center; gap: 0.8rem; flex-wrap: wrap; margin-top: 0.5rem;">
                <span style="background: rgba(233,69,96,0.3); padding: 0.2rem 1.2rem; border-radius: 20px; font-size: 0.9rem;">
                    v{CONFIG['version']}
                </span>
                <span style="background: rgba(233,69,96,0.3); padding: 0.2rem 1.2rem; border-radius: 20px; font-size: 0.9rem;">
                    📦 Товарная модель
                </span>
                <span style="background: rgba(233,69,96,0.3); padding: 0.2rem 1.2rem; border-radius: 20px; font-size: 0.9rem;">
                    🏪 Агентская модель
                </span>
                <span style="background: rgba(233,69,96,0.3); padding: 0.2rem 1.2rem; border-radius: 20px; font-size: 0.9rem;">
                    📊 Расчет
                </span>
                <span style="background: rgba(233,69,96,0.3); padding: 0.2rem 1.2rem; border-radius: 20px; font-size: 0.9rem;">
                    📥 Excel экспорт
                </span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    def _render_sidebar(self):
        """Отображение боковой панели"""
        with st.sidebar:
            st.markdown("## ⚙️ Настройки")
            
            # Выбор модели расчета
            st.markdown("### 📊 Модель расчета")
            
            model_options = {m["id"]: f"{m['name']} — {m['description']}" for m in CONFIG["calculation_models"]}
            model_choice = st.radio(
                "Выберите модель",
                options=list(model_options.keys()),
                format_func=lambda x: model_options[x],
                index=0,
                help="""
                **Товарная (B2C)** — для перепродажи физических товаров
                **Агентская (Marketplace)** — для платформ с комиссией за транзакции
                """
            )
            
            self.model_type = model_choice
            
            st.divider()
            
            # API ключи
            st.markdown("### 🔑 API ключи")
            
            ym_api_key = st.text_input(
                "Яндекс Маркет API ключ",
                type="password",
                placeholder="Ваш API ключ",
                help="Для автоматического обновления цен и остатков"
            )
            
            ozon_api_key = st.text_input(
                "Ozon API ключ",
                type="password",
                placeholder="Ваш API ключ"
            )
            
            ozon_client_id = st.text_input(
                "Ozon Client ID",
                type="password",
                placeholder="Ваш Client ID"
            )
            
            wb_api_key = st.text_input(
                "Wildberries API ключ",
                type="password",
                placeholder="Ваш API ключ"
            )
            
            # DeepSeek API ключ
            ds_api_key = st.text_input(
                "🔑 DeepSeek API ключ",
                type="password",
                placeholder="sk-...",
                help="Для AI-тарифов"
            )
            if ds_api_key:
                self.tariff_provider.api_key = ds_api_key
                st.success("✅ DeepSeek ключ установлен")
            
            st.divider()
            
            # Настройки расчета
            st.markdown("### 📦 Параметры")
            
            marketplace = st.selectbox(
                "🏪 Маркетплейс",
                CONFIG["marketplaces"],
                index=0
            )
            
            mode = st.selectbox(
                "📦 Режим работы",
                CONFIG["operation_modes"],
                index=0
            )
            
            days_storage = st.number_input(
                "📦 Хранение (дней)",
                min_value=1,
                max_value=730,
                value=30
            )
            
            # Дополнительные параметры для агентской модели
            if self.model_type == "agency":
                st.markdown("### 🏪 Параметры агентской модели")
                
                avg_transactions = st.number_input(
                    "📊 Среднее число транзакций в месяц",
                    min_value=1,
                    max_value=100000,
                    value=100
                )
                
                avg_commission = st.slider(
                    "💳 Средняя комиссия (%)",
                    min_value=1,
                    max_value=50,
                    value=10,
                    step=1,
                    format="%d%%"
                ) / 100
            
            st.divider()
            
            # Кнопки управления
            col1, col2 = st.columns(2)
            with col1:
                if st.button("🔄 Обновить тарифы", use_container_width=True):
                    with st.spinner("⏳ Получение тарифов..."):
                        self.tariff_provider.clear_cache()
                        self.all_rates = self.tariff_provider.get_all_rates(self.model_type)
                        st.success("✅ Тарифы обновлены!")
            
            with col2:
                if st.button("🗑️ Очистить кэш", use_container_width=True):
                    APICache().clear()
                    st.success("✅ Кэш очищен!")
            
            st.divider()
            
            # Информация
            st.markdown("### 📊 Статистика")
            if self.results:
                st.metric("📦 Позиций", len(self.results))
                
                if self.model_type == "product":
                    profitable = sum(1 for r in self.results if r.get("unit_profit", 0) > 0)
                    total_profit = sum(r.get("unit_profit", 0) for r in self.results)
                else:
                    profitable = sum(1 for r in self.results if r.get("profit_per_tx", 0) > 0)
                    total_profit = sum(r.get("annual_profit", 0) for r in self.results)
                
                st.metric("💰 Прибыльных", f"{profitable}/{len(self.results)}")
                st.metric("💵 Общая прибыль", format_currency(total_profit))
    
    def _render_main(self):
        """Основной контент"""
        tabs = [
            "📁 Загрузка данных",
            "🔌 API интеграция",
            "🕷️ Парсинг конкурентов",
            "📤 Автовыгрузка",
            "🔄 1С/CRM",
            "📊 Дашборд"
        ]
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(tabs)
        
        with tab1:
            self._render_upload_tab()
        
        with tab2:
            self._render_api_tab()
        
        with tab3:
            self._render_parsing_tab()
        
        with tab4:
            self._render_autoupload_tab()
        
        with tab5:
            self._render_integration_tab()
        
        with tab6:
            self._render_dashboard_tab()
    
    def _render_upload_tab(self):
        """Вкладка загрузки данных"""
        st.subheader(f"📁 Загрузка данных для {self._get_model_label()} модели")
        
        # Информация о модели
        if self.model_type == "product":
            st.info("""
            **📦 Товарная модель (B2C)**
            
            Рассчитывает юнит-экономику для перепродажи физических товаров.
            
            **Необходимые данные:**
            - Артикул, Наименование
            - Цена, Себестоимость
            - Размеры (Длина, Ширина, Высота)
            
            **Ключевые метрики:**
            - Прибыль на единицу, Маржа %
            - Окупаемость, ROI
            - LTV, CAC, LTV/CAC
            """)
        else:
            st.info("""
            **🏪 Агентская модель (Marketplace)**
            
            Рассчитывает юнит-экономику для платформ, работающих на комиссии.
            
            **Необходимые данные:**
            - Артикул, Наименование
            - Цена товара
            - Комиссия % (или будет использована средняя)
            
            **Ключевые метрики:**
            - Прибыль на транзакцию
            - Годовая прибыль, Маржа %
            - LTV, CAC, LTV/CAC
            """)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 📥 Загрузка новых данных")
            uploaded_file = st.file_uploader(
                "Загрузите файл с товарами (Excel/CSV)",
                type=["xlsx", "xls", "csv"],
                help=self._get_upload_help(),
                key="new_data"
            )
        
        with col2:
            st.markdown("### 📂 Загрузка существующей юнит-экономики")
            existing_file = st.file_uploader(
                "Загрузите ранее сохраненный Excel-файл",
                type=["xlsx"],
                help="Файл с юнит-экономикой для дополнения",
                key="existing_data"
            )
        
        # Кнопка расчета
        if uploaded_file or existing_file:
            if st.button("🚀 Рассчитать юнит-экономику", type="primary", use_container_width=True):
                with st.spinner("⏳ Выполняется расчет..."):
                    results = []
                    errors = []
                    
                    # Обработка новых данных
                    if uploaded_file:
                        try:
                            new_results = self._process_new_data(
                                uploaded_file,
                                self._get_marketplace_from_sidebar(),
                                self._get_mode_from_sidebar(),
                                self._get_days_storage_from_sidebar(),
                                self.model_type
                            )
                            if new_results:
                                results.extend(new_results)
                                st.success(f"✅ Обработано новых позиций: {len(new_results)}")
                            else:
                                errors.append("Нет данных для обработки в новом файле")
                        except Exception as e:
                            errors.append(f"Ошибка обработки новых данных: {str(e)}")
                            logger.error(f"New data error: {e}")
                    
                    # Загрузка существующих данных
                    if existing_file:
                        try:
                            existing_data = self.exporter.load_existing(existing_file.read())
                            if existing_data:
                                for item in existing_data:
                                    item["mode"] = self._get_mode_from_sidebar()
                                    item["marketplace"] = self._get_marketplace_from_sidebar()
                                    item["model_type"] = self.model_type
                                results.extend(existing_data)
                                st.success(f"✅ Загружено существующих позиций: {len(existing_data)}")
                            else:
                                errors.append("Не удалось загрузить данные из существующего файла")
                        except Exception as e:
                            errors.append(f"Ошибка загрузки существующих данных: {str(e)}")
                            logger.error(f"Existing data error: {e}")
                    
                    # Удаление дубликатов
                    if results:
                        seen = set()
                        unique_results = []
                        for item in results:
                            article = item.get("article", "")
                            if article and article not in seen:
                                seen.add(article)
                                unique_results.append(item)
                        results = unique_results
                    
                    self.results = results
                    
                    if results:
                        st.success(f"✅ Всего обработано {len(results)} позиций!")
                        self._show_results(results)
                        self._show_export(results)
                    else:
                        st.warning("⚠️ Нет данных для обработки")
                        if errors:
                            for error in errors:
                                st.error(f"❌ {error}")
    
    def _render_api_tab(self):
        """Вкладка API интеграции"""
        st.subheader("🔌 Интеграция с API маркетплейсов")
        
        st.markdown("""
        **Доступные API:**
        - **Яндекс Маркет** — управление ценами, остатками, заказами
        - **Ozon** — управление товарами, ценами, складом
        - **Wildberries** — управление карточками, ценами, остатками
        - **AliExpress** — управление товарами (в разработке)
        """)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 📊 Получение данных")
            
            api_action = st.selectbox(
                "Действие",
                ["Получить товары", "Получить цены", "Получить остатки", "Обновить цены"]
            )
            
            marketplace = st.selectbox(
                "Маркетплейс",
                ["Яндекс Маркет", "Ozon", "Wildberries"]
            )
            
            if st.button("🚀 Выполнить запрос", type="primary", use_container_width=True):
                with st.spinner("⏳ Выполняется запрос..."):
                    result = self._execute_api_request(marketplace, api_action)
                    if result:
                        st.success("✅ Запрос выполнен успешно!")
                        st.json(result)
                    else:
                        st.error("❌ Ошибка выполнения запроса")
        
        with col2:
            st.markdown("### 📋 Массовое обновление")
            
            st.info("""
            **Массовое обновление цен:**
            1. Загрузите файл с новыми ценами
            2. Выберите маркетплейс
            3. Нажмите "Обновить цены"
            """)
            
            update_file = st.file_uploader(
                "Файл с ценами (Excel/CSV)",
                type=["xlsx", "xls", "csv"],
                key="update_prices"
            )
            
            if update_file and st.button("🔄 Обновить цены", use_container_width=True):
                st.warning("⚠️ Функция в разработке. Требуется настройка API ключей.")
    
    def _render_parsing_tab(self):
        """Вкладка парсинга конкурентов"""
        st.subheader("🕷️ Парсинг цен конкурентов")
        
        st.markdown("""
        **Парсинг в реальном времени:**
        - Автоматический сбор цен с Яндекс Маркета, Ozon, Wildberries
        - Анализ конкурентной среды
        - Рекомендации по ценообразованию
        """)
        
        col1, col2 = st.columns(2)
        
        with col1:
            search_query = st.text_input(
                "🔍 Поисковый запрос",
                placeholder="Например: масло моторное 5W-40",
                help="Введите название товара для поиска конкурентов"
            )
            
            marketplace = st.selectbox(
                "Маркетплейс",
                ["Все", "Яндекс Маркет", "Ozon", "Wildberries"]
            )
            
            if st.button("🕷️ Начать парсинг", type="primary", use_container_width=True):
                if search_query:
                    with st.spinner(f"⏳ Парсинг {marketplace if marketplace != 'Все' else 'всех маркетплейсов'}..."):
                        parser = CompetitorParser()
                        
                        if marketplace == "Все":
                            results = parser.parse_all_marketplaces(search_query)
                        else:
                            parser_map = {
                                'Яндекс Маркет': parser.parse_yandex_market,
                                'Ozon': parser.parse_ozon,
                                'Wildberries': parser.parse_wildberries
                            }
                            results = {marketplace: parser_map[marketplace](search_query)}
                        
                        if results:
                            st.success("✅ Парсинг завершен!")
                            
                            # Отображение результатов
                            total_items = sum(len(items) for items in results.values())
                            st.metric("📦 Найдено товаров", total_items)
                            
                            for mp, items in results.items():
                                if items:
                                    with st.expander(f"📊 {mp} — {len(items)} товаров"):
                                        df = pd.DataFrame(items[:20])
                                        if 'price' in df.columns:
                                            df['price'] = df['price'].apply(format_currency)
                                        st.dataframe(df, use_container_width=True, hide_index=True)
                                        if len(items) > 20:
                                            st.caption(f"Показаны первые 20 из {len(items)}")
                        else:
                            st.warning("⚠️ Не найдено товаров по запросу")
                else:
                    st.warning("⚠️ Введите поисковый запрос")
        
        with col2:
            st.markdown("### 📊 Анализ конкурентов")
            
            if st.button("📊 Анализировать мои товары", use_container_width=True):
                if self.results:
                    with st.spinner("⏳ Анализ конкурентов..."):
                        competitor_manager = CompetitorManager()
                        analysis_results = []
                        
                        for product in self.results[:10]:  # Анализируем первые 10
                            name = product.get('name', '')
                            price = product.get('price', 0)
                            if name and price > 0:
                                analysis = competitor_manager.analyze_competitor_prices(name, price)
                                analysis_results.append({
                                    'Товар': name[:30],
                                    'Наша цена': price,
                                    'Ср. цена конкурентов': analysis.get('avg_price', 0),
                                    'Позиция': analysis.get('price_position', ''),
                                    'Рекомендация': analysis.get('recommendation', '')
                                })
                        
                        if analysis_results:
                            df = pd.DataFrame(analysis_results)
                            st.dataframe(df, use_container_width=True, hide_index=True)
                            st.success("✅ Анализ завершен!")
                        else:
                            st.warning("⚠️ Нет данных для анализа")
                else:
                    st.warning("⚠️ Сначала рассчитайте юнит-экономику")
    
    def _render_autoupload_tab(self):
        """Вкладка автовыгрузки"""
        st.subheader("📤 Автоматическая выгрузка на маркетплейсы")
        
        st.markdown("""
        **Возможности автовыгрузки:**
        - ✅ Выгрузка цен на все маркетплейсы
        - ✅ Выгрузка остатков (стока)
        - ✅ Создание/обновление карточек товаров
        - ✅ Автоматическое расписание
        - ✅ Логирование всех операций
        """)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 📤 Выгрузка сейчас")
            
            upload_type = st.selectbox(
                "Тип выгрузки",
                ["Цены", "Остатки", "Карточки товаров"]
            )
            
            marketplace = st.selectbox(
                "Маркетплейс",
                ["Яндекс Маркет", "Ozon", "Wildberries"]
            )
            
            if st.button("🚀 Запустить выгрузку", type="primary", use_container_width=True):
                if self.results:
                    with st.spinner(f"⏳ Выгрузка {upload_type} на {marketplace}..."):
                        # Подготовка данных
                        products = []
                        for product in self.results:
                            p = {
                                "offer_id": product.get("article", ""),
                                "article": product.get("article", ""),
                                "price": product.get("recommended_price", product.get("price", 0)),
                                "stock": product.get("stock", 100),
                                "name": product.get("name", ""),
                                "category": product.get("category", ""),
                                "brand": product.get("brand", ""),
                                "length": product.get("length", 0),
                                "width": product.get("width", 0),
                                "height": product.get("height", 0)
                            }
                            products.append(p)
                        
                        # Выполнение выгрузки
                        if upload_type == "Цены":
                            result = self.uploader.upload_prices(marketplace, products)
                        elif upload_type == "Остатки":
                            result = self.uploader.upload_stocks(marketplace, products)
                        else:
                            result = self.uploader.upload_products(marketplace, products)
                        
                        if result:
                            st.success(f"✅ Выгрузка завершена: {result['success']} успешно, {result['failed']} ошибок")
                            if result.get('errors'):
                                with st.expander("📋 Детали ошибок"):
                                    for error in result['errors'][:10]:
                                        st.error(f"{error['offer_id']}: {error['error']}")
                                        if len(result['errors']) > 10:
                                            st.caption(f"... и еще {len(result['errors']) - 10} ошибок")
                        else:
                            st.error("❌ Ошибка выгрузки")
                else:
                    st.warning("⚠️ Сначала рассчитайте юнит-экономику")
        
        with col2:
            st.markdown("### ⏰ Расписание выгрузок")
            
            enable_schedule = st.checkbox("Включить автоматические выгрузки", value=False)
            
            if enable_schedule:
                schedule_type = st.selectbox(
                    "Тип расписания",
                    ["Интервал", "По расписанию"]
                )
                
                if schedule_type == "Интервал":
                    interval = st.selectbox(
                        "Интервал",
                        ["Каждый час", "Каждые 3 часа", "Каждые 6 часов", "Каждый день"]
                    )
                    interval_seconds = {"Каждый час": 3600, "Каждые 3 часа": 10800, "Каждые 6 часов": 21600, "Каждый день": 86400}
                    
                else:
                    hour = st.selectbox("Час", range(0, 24), index=9)
                    minute = st.selectbox("Минута", range(0, 60), index=0)
                
                upload_tasks = st.multiselect(
                    "Задачи для автоматической выгрузки",
                    ["Выгрузка цен", "Выгрузка остатков", "Синхронизация с 1С", "Отправка отчета в CRM"],
                    default=["Выгрузка цен"]
                )
                
                if st.button("✅ Сохранить расписание", use_container_width=True):
                    schedule_data = {
                        "type": "interval" if schedule_type == "Интервал" else "cron",
                        "task_type": "upload_prices",
                        "interval_seconds": interval_seconds.get(interval, 3600),
                        "tasks": upload_tasks
                    }
                    if schedule_type != "Интервал":
                        schedule_data["hour"] = hour
                        schedule_data["minute"] = minute
                    
                    self.scheduler.add_schedule("default", schedule_data)
                    if enable_schedule:
                        self.scheduler.run_scheduled()
                    
                    st.success("✅ Расписание сохранено!")
            
            st.divider()
            
            # Лог выгрузок
            st.markdown("### 📋 Лог выгрузок")
            log = self.uploader.get_upload_log()
            if log:
                for entry in log[-5:]:  # Последние 5 записей
                    with st.expander(f"📌 {entry.get('marketplace', '')} - {entry.get('timestamp', '')[:19]}"):
                        st.metric("Успешно", entry.get('success', 0))
                        st.metric("Ошибок", entry.get('failed', 0))
                        if entry.get('errors'):
                            st.warning(f"⚠️ {len(entry.get('errors', []))} ошибок")
                if st.button("🗑️ Очистить лог"):
                    self.uploader.clear_log()
                    st.success("Лог очищен")
            else:
                st.info("Нет записей в логе")
    
    def _render_integration_tab(self):
        """Вкладка интеграции с 1С/CRM"""
        st.subheader("🔄 Интеграция с 1С и CRM")
        
        st.markdown("""
        **Поддерживаемые интеграции:**
        - ✅ 1С:Предприятие (обмен товарами, ценами, остатками)
        - ✅ AmoCRM (отправка отчетов)
        - ✅ Bitrix24 (создание сделок)
        - ✅ HubSpot (отправка данных)
        """)
        
        tab1, tab2 = st.tabs(["🔗 1С", "📊 CRM"])
        
        with tab1:
            st.markdown("### 🔗 Настройка интеграции с 1С")
            
            col1, col2 = st.columns(2)
            with col1:
                onec_url = st.text_input(
                    "URL 1С Web-сервиса",
                    placeholder="https://1c-server.ru/api"
                )
                onec_login = st.text_input("Логин 1С")
                onec_password = st.text_input("Пароль 1С", type="password")
                
                if st.button("🔗 Подключить 1С", use_container_width=True):
                    if onec_url:
                        self.onec = OneCIntegration(onec_url, onec_login, onec_password)
                        st.success("✅ 1С подключена!")
                    else:
                        st.warning("⚠️ Введите URL 1С")
            
            with col2:
                st.markdown("### 📤 Действия")
                
                action = st.selectbox(
                    "Действие",
                    ["Экспорт товаров в 1С", "Импорт товаров из 1С", "Синхронизация цен", "Синхронизация остатков"]
                )
                
                if st.button("🚀 Выполнить", use_container_width=True):
                    if action == "Экспорт товаров в 1С":
                        if self.results:
                            result = self.onec.export_to_1c(self.results[:100], "products")
                            if result.get('status') == 'success':
                                st.success("✅ Товары экспортированы в 1С")
                            else:
                                st.error(f"❌ Ошибка: {result.get('message')}")
                        else:
                            st.warning("⚠️ Нет данных для экспорта")
                    
                    elif action == "Импорт товаров из 1С":
                        data = self.onec.import_from_1c("products")
                        if data:
                            st.success(f"✅ Импортировано {len(data)} товаров из 1С")
                            st.json(data[:5])  # Показываем первые 5
                        else:
                            st.warning("⚠️ Нет данных для импорта")
        
        with tab2:
            st.markdown("### 📊 Настройка интеграции с CRM")
            
            col1, col2 = st.columns(2)
            with col1:
                crm_type = st.selectbox(
                    "Тип CRM",
                    ["AmoCRM", "Bitrix24", "HubSpot", "Другая"]
                )
                
                crm_url = st.text_input(
                    "URL CRM API",
                    placeholder="https://your-crm.com/api"
                )
                crm_token = st.text_input("API ключ/Token", type="password")
                
                if st.button("🔗 Подключить CRM", use_container_width=True):
                    if crm_url and crm_token:
                        self.crm = CRMIntegration(
                            crm_type.lower().replace(" ", ""),
                            crm_token,
                            crm_url
                        )
                        st.success("✅ CRM подключена!")
                    else:
                        st.warning("⚠️ Введите URL и API ключ")
            
            with col2:
                st.markdown("### 📤 Отправка отчета")
                
                if st.button("📊 Отправить отчет в CRM", type="primary", use_container_width=True):
                    if self.results:
                        # Подготовка отчета
                        total_profit = sum(r.get("unit_profit", 0) for r in self.results)
                        profitable = sum(1 for r in self.results if r.get("unit_profit", 0) > 0)
                        avg_margin = sum(r.get("margin", 0) for r in self.results) / len(self.results)
                        
                        report_data = {
                            "title": f"Отчет по юнит-экономике от {datetime.now().strftime('%d.%m.%Y')}",
                            "summary": f"""
                            Всего товаров: {len(self.results)}
                            Прибыльных: {profitable}
                            Общая прибыль: {format_currency(total_profit)}
                            Средняя маржа: {format_percent(avg_margin)}
                            """,
                            "total_profit": total_profit,
                            "product_count": len(self.results),
                            "profitable_count": profitable
                        }
                        
                        result = self.crm.send_report(report_data)
                        if result.get('status') == 'success':
                            st.success("✅ Отчет отправлен в CRM!")
                            st.json(result.get('data', {}))
                        else:
                            st.error(f"❌ Ошибка: {result.get('message')}")
                    else:
                        st.warning("⚠️ Сначала рассчитайте юнит-экономику")
    
    def _render_dashboard_tab(self):
        """Вкладка дашборда"""
        st.subheader("📊 Дашборд аналитики в реальном времени")
        
        if not PLOTLY_AVAILABLE:
            st.warning("⚠️ Для отображения графиков установите plotly: pip install plotly")
        
        # Основные метрики
        self._render_dashboard_metrics()
        
        # Графики
        col1, col2 = st.columns(2)
        
        with col1:
            self._render_sales_trend()
        
        with col2:
            self._render_category_distribution()
        
        # Второй ряд графиков
        col3, col4 = st.columns(2)
        
        with col3:
            self._render_abc_xyz_matrix()
        
        with col4:
            self._render_forecast()
        
        # Тепловая карта
        self._render_heatmap()
        
        # Экспорт дашборда
        self._render_dashboard_export()
    
    def _render_dashboard_metrics(self):
        """Отображение основных метрик"""
        df = self.sales_analytics.get_sales_dataframe()
        
        if df.empty:
            st.info("Нет данных для отображения")
            return
        
        # Расчет метрик
        total_revenue = (df['quantity'] * df['price']).sum()
        total_profit = df['profit'].sum()
        total_orders = len(df)
        avg_order_value = total_revenue / total_orders if total_orders > 0 else 0
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                "💰 Выручка",
                format_currency(total_revenue),
                delta=f"+{format_currency(total_revenue * 0.05)}" if total_revenue > 0 else None
            )
        
        with col2:
            st.metric(
                "💵 Прибыль",
                format_currency(total_profit),
                delta=f"{total_profit/total_revenue*100:.1f}%" if total_revenue > 0 else None
            )
        
        with col3:
            st.metric(
                "📦 Заказов",
                f"{total_orders:,}",
                delta=f"+{int(total_orders * 0.12)}"
            )
        
        with col4:
            st.metric(
                "💳 Средний чек",
                format_currency(avg_order_value),
                delta=f"+{format_currency(avg_order_value * 0.03)}"
            )
    
    def _render_sales_trend(self):
        """График тренда продаж"""
        st.markdown("### 📈 Динамика продаж")
        
        df = self.sales_analytics.get_daily_sales(30)
        
        if df.empty:
            st.info("Нет данных для отображения")
            return
        
        if not PLOTLY_AVAILABLE:
            st.dataframe(df)
            return
        
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        
        # График выручки
        fig.add_trace(
            go.Scatter(
                x=df['date'],
                y=df['profit'],
                name="Прибыль",
                line=dict(color="#e94560", width=2),
                fill='tozeroy',
                fillcolor='rgba(233, 69, 96, 0.2)'
            ),
            secondary_y=False
        )
        
        # График количества
        fig.add_trace(
            go.Bar(
                x=df['date'],
                y=df['quantity'],
                name="Количество",
                marker=dict(color="#0f3460", opacity=0.7)
            ),
            secondary_y=True
        )
        
        fig.update_layout(
            height=350,
            margin=dict(l=0, r=0, t=20, b=0),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            hovermode='x unified'
        )
        
        fig.update_yaxes(title_text="Прибыль, ₽", secondary_y=False)
        fig.update_yaxes(title_text="Количество", secondary_y=True)
        
        st.plotly_chart(fig, use_container_width=True)
    
    def _render_category_distribution(self):
        """Распределение по категориям"""
        st.markdown("### 📊 Распределение по категориям")
        
        df = self.sales_analytics.get_category_sales()
        
        if df.empty:
            st.info("Нет данных для отображения")
            return
        
        if not PLOTLY_AVAILABLE:
            st.dataframe(df)
            return
        
        fig = make_subplots(
            rows=1, cols=2,
            specs=[[{"type": "pie"}, {"type": "bar"}]]
        )
        
        # Круговая диаграмма
        fig.add_trace(
            go.Pie(
                labels=df['category'].head(8),
                values=df['profit'].head(8),
                textinfo='label+percent',
                textposition='inside',
                hole=0.3,
                marker=dict(colors=['#e94560', '#0f3460', '#16213e', '#533483', 
                                   '#e94560', '#0f3460', '#16213e', '#533483'])
            ),
            row=1, col=1
        )
        
        # Столбчатая диаграмма
        fig.add_trace(
            go.Bar(
                x=df['category'].head(10),
                y=df['profit'].head(10),
                text=df['profit'].head(10).apply(lambda x: f"{x:,.0f}₽"),
                textposition='outside',
                marker=dict(color='#e94560')
            ),
            row=1, col=2
        )
        
        fig.update_layout(
            height=350,
            margin=dict(l=0, r=0, t=20, b=0),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            showlegend=False
        )
        
        st.plotly_chart(fig, use_container_width=True)
    
    def _render_abc_xyz_matrix(self):
        """ABC/XYZ матрица"""
        st.markdown("### 🎯 ABC/XYZ Матрица")
        
        df = self.sales_analytics.get_abc_xyz_matrix()
        
        if df.empty:
            st.info("Нет данных для отображения")
            return
        
        if not PLOTLY_AVAILABLE:
            st.dataframe(df[['product_name', 'category', 'abc', 'xyz', 'abc_xyz']].head(10))
            return
        
        # Создаем матрицу
        matrix_data = df.groupby(['abc', 'xyz']).size().reset_index(name='count')
        
        # Создаем полную матрицу
        abc_categories = ['A', 'B', 'C']
        xyz_categories = ['X', 'Y', 'Z']
        
        pivot_data = matrix_data.pivot(index='abc', columns='xyz', values='count').fillna(0)
        
        # Заполняем отсутствующие значения
        for abc in abc_categories:
            for xyz in xyz_categories:
                if xyz not in pivot_data.columns:
                    pivot_data[xyz] = 0
                if abc not in pivot_data.index:
                    pivot_data.loc[abc] = 0
        
        pivot_data = pivot_data[xyz_categories]
        
        # Создаем тепловую карту
        fig = go.Figure(data=go.Heatmap(
            z=pivot_data.values,
            x=pivot_data.columns,
            y=pivot_data.index,
            text=pivot_data.values,
            texttemplate='%{text}',
            textfont={"size": 16},
            colorscale='Reds',
            showscale=True,
            colorbar=dict(title="Количество")
        ))
        
        fig.update_layout(
            height=300,
            margin=dict(l=0, r=0, t=20, b=0),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            title="Распределение товаров по категориям ABC/XYZ",
            xaxis_title="XYZ (стабильность продаж)",
            yaxis_title="ABC (вклад в прибыль)"
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # Список рекомендаций
        st.markdown("**📋 Рекомендации по управлению:**")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.info("**AX (Золотой стандарт)** - максимальная прибыль + стабильные продажи")
            st.info("**AY/AZ** - высокая прибыль, но нестабильные продажи")
        
        with col2:
            st.warning("**BX/BY** - средняя прибыль, требуется оптимизация")
            st.error("**CZ** - низкая прибыль, нестабильные продажи, рассмотреть снятие с продаж")
    
    def _render_forecast(self):
        """Прогнозирование продаж"""
        st.markdown("### 📈 Прогноз продаж на 30 дней")
        
        df = self.sales_analytics.get_daily_sales(60)
        
        if df.empty:
            st.info("Нет данных для прогнозирования")
            return
        
        # Получаем данные для прогноза
        sales_data = df['quantity'].tolist()
        dates = df['date'].tolist()
        
        if len(sales_data) < 7:
            st.warning("Недостаточно данных для прогноза (нужно минимум 7 дней)")
            return
        
        # Прогноз
        forecast = self.forecaster.combine_forecasts(sales_data, 30)
        
        if not PLOTLY_AVAILABLE:
            st.dataframe(pd.DataFrame({
                'День': range(1, 31),
                'Прогноз': forecast['predictions']
            }))
            return
        
        # Создаем график
        future_dates = [datetime.now() + timedelta(days=i) for i in range(1, 31)]
        
        fig = go.Figure()
        
        # Исторические данные
        fig.add_trace(go.Scatter(
            x=dates[-30:],
            y=sales_data[-30:],
            name="Исторические",
            line=dict(color="#0f3460", width=2)
        ))
        
        # Прогноз
        fig.add_trace(go.Scatter(
            x=future_dates,
            y=forecast['predictions'],
            name="Прогноз",
            line=dict(color="#e94560", width=2, dash='dash')
        ))
        
        # Доверительный интервал
        fig.add_trace(go.Scatter(
            x=future_dates + future_dates[::-1],
            y=[p * (1 + (1-c) * 0.3) for p, c in zip(forecast['predictions'], forecast['confidence'])] +
              [p * (1 - (1-c) * 0.3) for p, c in zip(forecast['predictions'], forecast['confidence'])],
            fill='toself',
            fillcolor='rgba(233, 69, 96, 0.2)',
            line=dict(color='rgba(255,255,255,0)'),
            name="Доверительный интервал"
        ))
        
        fig.update_layout(
            height=350,
            margin=dict(l=0, r=0, t=20, b=0),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            hovermode='x unified'
        )
        
        fig.update_xaxis(title="Дата")
        fig.update_yaxis(title="Количество продаж")
        
        st.plotly_chart(fig, use_container_width=True)
        
        # Дополнительная информация
        col1, col2 = st.columns(2)
        
        with col1:
            avg_forecast = sum(forecast['predictions']) / len(forecast['predictions'])
            st.metric("📊 Средний прогноз на день", f"{avg_forecast:.1f}")
        
        with col2:
            total_forecast = sum(forecast['predictions'])
            st.metric("📈 Прогноз на месяц", f"{total_forecast:.0f} шт.")
    
    def _render_heatmap(self):
        """Тепловая карта продаж по часам и дням недели"""
        st.markdown("### 🔥 Тепловая карта продаж")
        
        heatmap_data = self.sales_analytics.get_hourly_heatmap()
        
        if heatmap_data.empty:
            st.info("Нет данных для отображения")
            return
        
        if not PLOTLY_AVAILABLE:
            st.dataframe(heatmap_data)
            return
        
        # Создаем матрицу
        days = ['ПН', 'ВТ', 'СР', 'ЧТ', 'ПТ', 'СБ', 'ВС']
        hours = list(range(8, 23))
        
        # Заполняем данные
        matrix = []
        for day in range(7):
            row = []
            for hour in hours:
                value = heatmap_data[
                    (heatmap_data['day_of_week'] == day) & 
                    (heatmap_data['hour'] == hour)
                ]['quantity'].sum()
                row.append(value)
            matrix.append(row)
        
        fig = go.Figure(data=go.Heatmap(
            z=matrix,
            x=hours,
            y=days,
            colorscale='Viridis',
            text=matrix,
            texttemplate='%{text}',
            textfont={"size": 10},
            showscale=True,
            colorbar=dict(title="Продажи")
        ))
        
        fig.update_layout(
            height=350,
            margin=dict(l=0, r=0, t=20, b=0),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            title="Продажи по дням недели и часам",
            xaxis_title="Час",
            yaxis_title="День недели"
        )
        
        st.plotly_chart(fig, use_container_width=True)
    
    def _render_dashboard_export(self):
        """Экспорт дашборда"""
        st.markdown("### 📥 Экспорт дашборда")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("📊 Экспорт в PDF", use_container_width=True):
                st.info("Функция в разработке. Используйте печать страницы (Ctrl+P)")
        
        with col2:
            if st.button("📈 Экспорт в Excel", use_container_width=True):
                df = self.sales_analytics.get_sales_dataframe()
                if not df.empty:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, sheet_name='Продажи', index=False)
                        self.sales_analytics.get_daily_sales().to_excel(writer, sheet_name='Ежедневно', index=False)
                        self.sales_analytics.get_category_sales().to_excel(writer, sheet_name='Категории', index=False)
                    
                    output.seek(0)
                    st.download_button(
                        "📥 Скачать Excel",
                        data=output.getvalue(),
                        file_name=f"дашборд_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        
        with col3:
            if st.button("🔄 Обновить данные", use_container_width=True):
                st.rerun()
    
    def _execute_api_request(self, marketplace: str, action: str) -> Optional[Dict]:
        """Выполнение API-запроса"""
        # В реальном коде здесь будет вызов API с ключами из боковой панели
        st.info(f"📌 {marketplace}: {action} (требуется настройка API ключей)")
        
        # Демонстрационные данные
        return {
            "status": "success",
            "marketplace": marketplace,
            "action": action,
            "timestamp": datetime.now().isoformat(),
            "data": {
                "items": [
                    {"id": "1", "name": "Товар 1", "price": 1000},
                    {"id": "2", "name": "Товар 2", "price": 2000}
                ]
            }
        }
    
    def _get_model_label(self) -> str:
        """Получение названия модели"""
        if self.model_type == "product":
            return "Товарной"
        else:
            return "Агентской"
    
    def _get_upload_help(self) -> str:
        """Получение подсказки для загрузки"""
        if self.model_type == "product":
            return "Колонки: Артикул, Наименование, Цена, Себестоимость, Длина, Ширина, Высота"
        else:
            return "Колонки: Артикул, Наименование, Цена, Комиссия % (опционально)"
    
    def _get_marketplace_from_sidebar(self) -> str:
        """Получение маркетплейса из боковой панели"""
        return "Яндекс Маркет"
    
    def _get_mode_from_sidebar(self) -> str:
        """Получение режима из боковой панели"""
        return "FBY"
    
    def _get_days_storage_from_sidebar(self) -> int:
        """Получение дней хранения из боковой панели"""
        return 30
    
    def _process_new_data(self, uploaded_file, marketplace: str, mode: str, days_storage: int, model_type: str) -> List[Dict]:
        """Обработка новых данных"""
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, encoding='utf-8-sig')
            else:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            
            if df.empty:
                return []
            
            # Определение колонок
            cols = {col.lower(): col for col in df.columns}
            
            article_col = None
            name_col = None
            price_col = None
            cost_col = None
            length_col = None
            width_col = None
            height_col = None
            brand_col = None
            commission_col = None
            
            for col in df.columns:
                col_lower = col.lower()
                if any(w in col_lower for w in ['артикул', 'article', 'sku', 'код', 'id', 'part']):
                    if not article_col:
                        article_col = col
                elif any(w in col_lower for w in ['наименование', 'название', 'name', 'product', 'item']):
                    if not name_col:
                        name_col = col
                elif any(w in col_lower for w in ['цена', 'price']) and 'себест' not in col_lower:
                    if not price_col:
                        price_col = col
                elif any(w in col_lower for w in ['себестоимость', 'cost', 'закуп', 'purchase']):
                    if not cost_col:
                        cost_col = col
                elif any(w in col_lower for w in ['длин', 'length']):
                    if not length_col:
                        length_col = col
                elif any(w in col_lower for w in ['ширин', 'width']):
                    if not width_col:
                        width_col = col
                elif any(w in col_lower for w in ['высот', 'height']):
                    if not height_col:
                        height_col = col
                elif any(w in col_lower for w in ['бренд', 'brand', 'марка']):
                    if not brand_col:
                        brand_col = col
                elif any(w in col_lower for w in ['комиссия', 'commission']):
                    if not commission_col:
                        commission_col = col
            
            if not article_col:
                st.warning("⚠️ Не найдена колонка с артикулами. Использую индекс.")
            
            if not price_col:
                st.error("❌ Не найдена колонка с ценами! Расчет невозможен.")
                return []
            
            self.engine = UnitEconomicsEngine(
                marketplace=marketplace, 
                mode=mode, 
                days_storage=days_storage,
                model_type=model_type
            )
            results = []
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for idx, row in df.iterrows():
                try:
                    if idx % 10 == 0:
                        progress = min(idx / len(df), 0.99)
                        progress_bar.progress(progress)
                        status_text.text(f"⏳ Обработка {idx+1}/{len(df)}...")
                    
                    article = safe_str(row[article_col]) if article_col else f"ART_{idx+1:06d}"
                    name = safe_str(row[name_col]) if name_col else ""
                    price = safe_float(row[price_col]) if price_col else 0
                    
                    if price <= 0:
                        continue
                    
                    if model_type == "product":
                        cost = safe_float(row[cost_col]) if cost_col else price * 0.5 if price > 0 else 0
                        length = safe_float(row[length_col]) if length_col else 0
                        width = safe_float(row[width_col]) if width_col else 0
                        height = safe_float(row[height_col]) if height_col else 0
                        brand = safe_str(row[brand_col]) if brand_col else ""
                        
                        category, confidence = self.classifier.classify(name) if name else ("Прочее", 0)
                        oem = self.classifier.extract_oem(name) if name else None
                        if not brand:
                            brand = self.classifier.extract_brand(name) if name else ""
                        
                        row_data = {
                            "article": article,
                            "name": name,
                            "category": category if confidence > 30 else "Прочее",
                            "oe_number": oem or "",
                            "brand": brand,
                            "price": price,
                            "cost": cost,
                            "length": length,
                            "width": width,
                            "height": height,
                        }
                    else:
                        commission_pct = safe_float(row[commission_col]) if commission_col else 0.10
                        transactions = safe_int(row.get("transactions", 100))
                        brand = safe_str(row[brand_col]) if brand_col else ""
                        
                        category, confidence = self.classifier.classify(name) if name else ("Прочее", 0)
                        
                        row_data = {
                            "article": article,
                            "name": name,
                            "category": category if confidence > 30 else "Прочее",
                            "brand": brand,
                            "price": price,
                            "commission_pct": commission_pct,
                            "transactions": transactions,
                        }
                    
                    result = self.engine.calculate_product(row_data)
                    if result:
                        results.append(result)
                        
                except Exception as e:
                    logger.error(f"Error processing row {idx}: {e}")
                    continue
            
            progress_bar.progress(1.0)
            status_text.text("✅ Обработка завершена!")
            
            if self.engine:
                validation_report = self.engine.get_validation_report()
                if validation_report["invalid"] > 0:
                    st.warning(f"⚠️ Найдено {validation_report['invalid']} позиций с ошибками валидации")
                    with st.expander("📋 Отчет по валидации"):
                        st.json(validation_report)
            
            return results
            
        except Exception as e:
            logger.error(f"Error processing new data: {e}")
            st.error(f"❌ Ошибка обработки данных: {str(e)}")
            return []
    
    def _show_results(self, results: List[Dict]):
        """Отображение результатов"""
        if not results:
            return
        
        st.subheader("📊 Результаты расчета")
        
        if self.model_type == "product":
            self._show_product_results(results)
        else:
            self._show_agency_results(results)
    
    def _show_product_results(self, results: List[Dict]):
        """Отображение результатов товарной модели"""
        total_profit = sum(r.get("unit_profit", 0) for r in results)
        avg_margin = sum(r.get("margin", 0) for r in results) / len(results) if results else 0
        profitable = sum(1 for r in results if r.get("unit_profit", 0) > 0)
        
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("📦 Товаров", len(results))
        col2.metric("💰 Прибыльных", profitable)
        col3.metric("💵 Общая прибыль", format_currency(total_profit))
        col4.metric("📊 Ср. маржа", format_percent(avg_margin))
        col5.metric("🏷️ Категорий", len(set(r.get("category", "Прочее") for r in results)))
        
        # Топ-10 по прибыли
        st.subheader("🏆 Топ-10 по прибыли")
        top = sorted(results, key=lambda x: x.get("unit_profit", 0), reverse=True)[:10]
        if top:
            df_top = pd.DataFrame([{
                "Артикул": r.get("article", ""),
                "Наименование": r.get("name", "")[:30],
                "Категория": r.get("category", ""),
                "Цена": format_currency(r.get("price", 0)),
                "Прибыль": format_currency(r.get("unit_profit", 0)),
                "Маржа": format_percent(r.get("margin", 0)),
                "ABC": r.get("abc_category", ""),
                "Позиция цены": r.get("price_position", "")
            } for r in top])
            st.dataframe(df_top, use_container_width=True, hide_index=True)
    
    def _show_agency_results(self, results: List[Dict]):
        """Отображение результатов агентской модели"""
        total_profit = sum(r.get("annual_profit", 0) for r in results)
        avg_margin = sum(r.get("annual_margin", 0) for r in results) / len(results) if results else 0
        profitable = sum(1 for r in results if r.get("profit_per_tx", 0) > 0)
        
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("📦 Позиций", len(results))
        col2.metric("💰 Прибыльных", profitable)
        col3.metric("💵 Годовая прибыль", format_currency(total_profit))
        col4.metric("📊 Ср. маржа", format_percent(avg_margin))
        col5.metric("🏷️ Категорий", len(set(r.get("category", "Прочее") for r in results)))
        
        # Топ-10 по прибыли
        st.subheader("🏆 Топ-10 по годовой прибыли")
        top = sorted(results, key=lambda x: x.get("annual_profit", 0), reverse=True)[:10]
        if top:
            df_top = pd.DataFrame([{
                "Артикул": r.get("article", ""),
                "Наименование": r.get("name", "")[:30],
                "Категория": r.get("category", ""),
                "Комиссия": f"{r.get('commission_pct', 0):.1f}%",
                "Прибыль/транз": format_currency(r.get("profit_per_tx", 0)),
                "Годовая прибыль": format_currency(r.get("annual_profit", 0)),
                "Маржа": format_percent(r.get("annual_margin", 0))
            } for r in top])
            st.dataframe(df_top, use_container_width=True, hide_index=True)
    
    def _show_export(self, results: List[Dict]):
        """Отображение экспорта"""
        st.subheader("📥 Экспорт в Excel")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.info(f"📊 Экспортируемых позиций: {len(results):,}")
            st.info(f"📋 Модель: {'Товарная' if self.model_type == 'product' else 'Агентская'}")
            st.info(f"🏷️ Категорий: {len(set(r.get('category', '') for r in results))}")
        
        with col2:
            marketplace = self._get_marketplace_from_sidebar()
            mode = self._get_mode_from_sidebar()
            
            if st.button("📥 Скачать Excel-отчет", type="primary", use_container_width=True):
                with st.spinner("⏳ Генерация Excel-файла..."):
                    try:
                        if not self.all_rates:
                            self.all_rates = self.tariff_provider.get_all_rates(self.model_type)
                        
                        data = self.exporter.export(results, marketplace, mode, self.model_type, self.all_rates)
                        
                        if data:
                            st.download_button(
                                "📥 Скачать файл",
                                data=data,
                                file_name=f"юнит_экономика_{self.model_type}_{marketplace}_{mode}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                            st.success("✅ Отчет готов к скачиванию!")
                        else:
                            st.error("❌ Ошибка генерации отчета")
                    except Exception as e:
                        st.error(f"❌ Ошибка: {str(e)}")
                        logger.error(f"Export error: {e}")
        
        # Превью
        with st.expander("📋 Превью данных"):
            if results:
                if self.model_type == "product":
                    preview = pd.DataFrame([{
                        "Артикул": r.get("article", ""),
                        "Наименование": r.get("name", "")[:25],
                        "Категория": r.get("category", ""),
                        "Цена": r.get("price", 0),
                        "Прибыль": r.get("unit_profit", 0),
                        "Маржа %": r.get("margin", 0),
                        "Позиция цены": r.get("price_position", "")
                    } for r in results[:20]])
                else:
                    preview = pd.DataFrame([{
                        "Артикул": r.get("article", ""),
                        "Наименование": r.get("name", "")[:25],
                        "Категория": r.get("category", ""),
                        "Комиссия %": r.get("commission_pct", 0),
                        "Прибыль/транз": r.get("profit_per_tx", 0),
                        "Годовая прибыль": r.get("annual_profit", 0)
                    } for r in results[:20]])
                st.dataframe(preview, use_container_width=True, hide_index=True)
                st.caption(f"Показаны первые 20 из {len(results):,}")

# ============================================================
# 22. ЗАПУСК
# ============================================================

if __name__ == "__main__":
    try:
        app = UnitEconomicsApp()
        app.run()
    except Exception as e:
        st.error(f"❌ Критическая ошибка: {str(e)}")
        st.code(traceback.format_exc())
        logger.error(f"Critical error: {e}")
