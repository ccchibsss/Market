"""
================================================================================
🚀 ULTIMATE UNIT ECONOMICS ENGINE v54.0 - ПОЛНАЯ ВЕРСИЯ (БЕЗ СОКРАЩЕНИЙ)
================================================================================
📌 ВЕРСИЯ: 54.0.0
📌 ОБЩИЙ ОБЪЕМ: 5500+ СТРОК
📌 НОВЫЕ ФУНКЦИИ:
    ✅ ИСПРАВЛЕН РАСЧЕТ ЮНИТ-ЭКОНОМИКИ
    ✅ Улучшенная обработка ошибок
    ✅ Полный код без сокращений
    ✅ ИИ-редактирование данных
    ✅ Множественный парсинг
    ✅ Расширенная аналитика
    ✅ Визуализация графиков
    ✅ Экспорт в Excel с форматированием
    ✅ База OE номеров
    ✅ Конвертация размеров
================================================================================
"""

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
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional, Tuple, Union
from dataclasses import dataclass, field
from collections import Counter, defaultdict
from functools import lru_cache
from threading import Thread, Lock, Event
from queue import Queue
import traceback
import os
import pickle
import random
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from contextlib import contextmanager
from enum import Enum

# Подавление предупреждений
warnings.filterwarnings('ignore')
os.environ['PYTHONWARNINGS'] = 'ignore'

# --------------------------------------------
# ВЕРСИЯ И КОНФИГУРАЦИЯ
# --------------------------------------------
APP_VERSION = "54.0.0"
APP_NAME = "🚀 Юнит-экономика с ИИ-редактированием"

# --------------------------------------------
# НАСТРОЙКА ЛОГИРОВАНИЯ
# --------------------------------------------
class Logger:
    """Улучшенный логгер с поддержкой многопоточности"""
    _instance = None
    _lock = Lock()
    
    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
                    cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        if self._initialized:
            return
        self._initialized = True
        
        self.logger = logging.getLogger('UnitEconomy')
        self.logger.setLevel(logging.DEBUG)
        
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        fh = logging.FileHandler('unit_economy.log', encoding='utf-8')
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(formatter)
        self.logger.addHandler(fh)
        
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)
        ch.setFormatter(formatter)
        self.logger.addHandler(ch)
    
    def get(self):
        return self.logger

logger = Logger().get()

# --------------------------------------------
# ПРОВЕРКА НАЛИЧИЯ БИБЛИОТЕК
# --------------------------------------------
LIBRARIES = {
    'openpyxl': False,
    'plotly': False,
    'sklearn': False,
    'gspread': False,
    'openai': False,
}

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.dataframe import dataframe_to_rows
    LIBRARIES['openpyxl'] = True
except ImportError as e:
    logger.warning(f"OpenPyXL не установлен: {e}")

try:
    import plotly.graph_objects as go
    import plotly.express as px
    from plotly.subplots import make_subplots
    LIBRARIES['plotly'] = True
except ImportError:
    logger.warning("Plotly не установлен")

try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.naive_bayes import MultinomialNB
    from sklearn.pipeline import Pipeline
    from sklearn.model_selection import train_test_split
    from sklearn.metrics import accuracy_score, classification_report
    import joblib
    LIBRARIES['sklearn'] = True
except ImportError:
    logger.warning("Scikit-learn не установлен")

try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    LIBRARIES['gspread'] = True
except ImportError:
    logger.warning("gspread не установлен")

try:
    import openai
    LIBRARIES['openai'] = True
except ImportError:
    logger.warning("openai не установлен")

# --------------------------------------------
# КОНФИГУРАЦИЯ
# --------------------------------------------
class Config:
    """Класс конфигурации с поддержкой ENV переменных"""
    
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        if self._initialized:
            return
        self._initialized = True
        
        self.version = APP_VERSION
        self.app_name = APP_NAME
        self.currency = os.getenv('CURRENCY', '₽')
        self.language = os.getenv('LANGUAGE', 'ru')
        
        self.marketplaces = ["Яндекс Маркет", "Ozon", "Wildberries", "AliExpress", "Мегамаркет"]
        self.operation_modes = ["FBY", "FBS", "FBO", "DBS"]
        self.dimension_units = ["мм", "см"]
        self.default_dimension_unit = "мм"
        
        self.category_keywords = {
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
        }
        
        self.oem_patterns = [
            r'[0-9]{6,12}',
            r'[A-Z0-9]{6,12}',
            r'[A-Z]{2}[0-9]{6,10}',
            r'[A-Z]{2}[0-9]{4}[A-Z]{2}',
            r'[0-9]{4}[A-Z]{2}[0-9]{4}'
        ]
        
        self.barcode_patterns = [
            r'[0-9]{8}',
            r'[0-9]{12}',
            r'[0-9]{13}',
            r'[0-9]{14}'
        ]
        
        self.validation = {
            "min_price": float(os.getenv('MIN_PRICE', 10)),
            "max_price": float(os.getenv('MAX_PRICE', 1000000)),
            "min_cost": float(os.getenv('MIN_COST', 1)),
            "max_cost": float(os.getenv('MAX_COST', 500000)),
            "min_dimension": float(os.getenv('MIN_DIMENSION', 0.1)),
            "max_dimension": float(os.getenv('MAX_DIMENSION', 1000)),
            "min_volume": float(os.getenv('MIN_VOLUME', 0.001)),
            "max_volume": float(os.getenv('MAX_VOLUME', 10000)),
            "min_weight": float(os.getenv('MIN_WEIGHT', 0.001)),
            "max_weight": float(os.getenv('MAX_WEIGHT', 1000))
        }
        
        self.api = {
            "cache_ttl": int(os.getenv('CACHE_TTL', 300)),
            "max_retries": int(os.getenv('MAX_RETRIES', 3)),
            "timeout": int(os.getenv('API_TIMEOUT', 30)),
            "rate_limit": int(os.getenv('RATE_LIMIT', 10))
        }
        
        self.proxy = {
            "enabled": os.getenv('PROXY_ENABLED', 'false').lower() == 'true',
            "http": os.getenv('HTTP_PROXY', ''),
            "https": os.getenv('HTTPS_PROXY', '')
        }
    
    def get_text(self, key: str, lang: str = None) -> str:
        if lang is None:
            lang = self.language
        
        texts = {
            'ru': {
                'upload_title': '📁 Загрузка данных',
                'calculate': '🚀 Рассчитать',
                'export': '📥 Экспорт',
                'profit': 'Прибыль',
                'margin': 'Маржа',
                'price': 'Цена',
                'cost': 'Себестоимость',
                'ai_edit': '🤖 ИИ-редактирование',
                'fix_data': 'Исправить данные через ИИ',
                'ai_prompt': 'Опишите, что нужно исправить в данных'
            },
            'en': {
                'upload_title': '📁 Upload Data',
                'calculate': '🚀 Calculate',
                'export': '📥 Export',
                'profit': 'Profit',
                'margin': 'Margin',
                'price': 'Price',
                'cost': 'Cost',
                'ai_edit': '🤖 AI Editing',
                'fix_data': 'Fix data with AI',
                'ai_prompt': 'Describe what needs to be fixed in the data'
            }
        }
        
        return texts.get(lang, texts['ru']).get(key, key)

CONFIG = Config()

# --------------------------------------------
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# --------------------------------------------
@contextmanager
def timer(name: str):
    start = time.time()
    try:
        yield
    finally:
        elapsed = time.time() - start
        logger.info(f"⏱ {name}: {elapsed:.2f}с")

def safe_float(val: Any, default: float = 0.0) -> float:
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
    try:
        if val is None:
            return default
        if isinstance(val, (int, float)) and (math.isnan(val) or math.isinf(val)):
            return default
        return str(val).strip() if str(val).strip() else default
    except (ValueError, TypeError):
        return default

def safe_int(val: Any, default: int = 0) -> int:
    try:
        return int(safe_float(val, default))
    except (ValueError, TypeError):
        return default

def calculate_volume(length: float, width: float, height: float) -> float:
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

def convert_dimension(value: float, from_unit: str, to_unit: str) -> float:
    if value == 0:
        return 0.0
    
    if from_unit == to_unit:
        return value
    
    if from_unit == "мм" and to_unit == "см":
        return value / 10.0
    elif from_unit == "см" and to_unit == "мм":
        return value * 10.0
    
    return value

def format_currency(value: float) -> str:
    try:
        if value is None or math.isnan(value) or math.isinf(value):
            return "0 ₽"
        currency = CONFIG.currency
        return f"{value:,.0f} {currency}" if abs(value) >= 1 else f"{value:.2f} {currency}"
    except (ValueError, TypeError):
        return "0 ₽"

def format_percent(value: float) -> str:
    try:
        if value is None or math.isnan(value) or math.isinf(value):
            return "0%"
        return f"{value:.1f}%" if abs(value) >= 0.1 else f"{value:.2f}%"
    except (ValueError, TypeError):
        return "0%"

def generate_cache_key(*args) -> str:
    key = "|".join(str(arg) for arg in args)
    return hashlib.md5(key.encode()).hexdigest()

def is_valid_barcode(barcode: str) -> bool:
    if not barcode:
        return False
    barcode = re.sub(r'[^\d]', '', barcode)
    if len(barcode) not in [8, 12, 13, 14]:
        return False
    return True

def format_barcode(barcode: str) -> str:
    if not barcode:
        return ""
    barcode = re.sub(r'[^\d]', '', barcode)
    if len(barcode) == 13:
        return f"{barcode[:3]} {barcode[3:7]} {barcode[7:11]} {barcode[11:]}"
    elif len(barcode) == 12:
        return f"{barcode[:2]} {barcode[2:6]} {barcode[6:10]} {barcode[10:]}"
    elif len(barcode) == 8:
        return f"{barcode[:2]} {barcode[2:5]} {barcode[5:]}"
    return barcode

def validate_article(article: str) -> bool:
    if not article or not article.strip():
        return False
    return bool(re.match(r'^[A-Za-z0-9\-_]+$', article.strip()))

def normalize_text(text: str) -> str:
    if not text:
        return ""
    text = text.lower()
    text = re.sub(r'[^\w\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def extract_numbers(text: str) -> List[float]:
    if not text:
        return []
    return [float(x) for x in re.findall(r'\d+\.?\d*', text)]

def calculate_price_recommendation(price: float, competitor_avg: float, margin: float) -> Tuple[float, str]:
    if margin < 15:
        return price * 1.15, "Повысить (низкая маржа)"
    elif margin > 35:
        return price * 0.95, "Снизить (высокая маржа)"
    elif competitor_avg > 0 and price > competitor_avg * 1.2:
        return competitor_avg * 0.95, "Снизить (выше конкурентов)"
    elif competitor_avg > 0 and price < competitor_avg * 0.8:
        return competitor_avg * 1.05, "Повысить (ниже конкурентов)"
    return price, "Оставить (оптимально)"

# --------------------------------------------
# 📊 КЭШИРОВАНИЕ
# --------------------------------------------
class CacheManager:
    def __init__(self, cache_dir: str = "cache"):
        self.cache_dir = cache_dir
        self.memory_cache = {}
        self.lock = Lock()
        self.stats = {
            'hits': 0,
            'misses': 0,
            'size': 0
        }
        
        if not os.path.exists(cache_dir):
            os.makedirs(cache_dir)
        
        self._clean_old_cache()
    
    def _clean_old_cache(self, max_age_days: int = 7):
        try:
            now = time.time()
            for filename in os.listdir(self.cache_dir):
                filepath = os.path.join(self.cache_dir, filename)
                if os.path.isfile(filepath):
                    if now - os.path.getmtime(filepath) > max_age_days * 86400:
                        try:
                            os.remove(filepath)
                            logger.info(f"Удален старый кэш: {filename}")
                        except:
                            pass
        except Exception as e:
            logger.warning(f"Ошибка очистки кэша: {e}")
    
    def get(self, key: str) -> Optional[Any]:
        with self.lock:
            if key in self.memory_cache:
                data, timestamp = self.memory_cache[key]
                if (datetime.now() - timestamp).total_seconds() < CONFIG.api["cache_ttl"]:
                    self.stats['hits'] += 1
                    return data
            
            cache_file = os.path.join(self.cache_dir, f"{hashlib.md5(key.encode()).hexdigest()}.pkl")
            if os.path.exists(cache_file):
                try:
                    with open(cache_file, 'rb') as f:
                        data, timestamp = pickle.load(f)
                        if (datetime.now() - timestamp).total_seconds() < CONFIG.api["cache_ttl"]:
                            self.memory_cache[key] = (data, timestamp)
                            self.stats['hits'] += 1
                            return data
                except Exception as e:
                    logger.warning(f"Ошибка чтения кэша: {e}")
            
            self.stats['misses'] += 1
            return None
    
    def set(self, key: str, value: Any):
        with self.lock:
            timestamp = datetime.now()
            self.memory_cache[key] = (value, timestamp)
            self.stats['size'] = len(self.memory_cache)
            
            try:
                cache_file = os.path.join(self.cache_dir, f"{hashlib.md5(key.encode()).hexdigest()}.pkl")
                with open(cache_file, 'wb') as f:
                    pickle.dump((value, timestamp), f)
            except Exception as e:
                logger.warning(f"Ошибка записи кэша: {e}")
    
    def clear(self):
        with self.lock:
            self.memory_cache.clear()
            self.stats['size'] = 0
            for file in os.listdir(self.cache_dir):
                try:
                    os.remove(os.path.join(self.cache_dir, file))
                except:
                    pass
            logger.info("Кэш очищен")
    
    def get_stats(self) -> Dict:
        return self.stats.copy()

# --------------------------------------------
# 🤖 ML-КАТЕГОРИЗАЦИЯ
# --------------------------------------------
class AutoClassifier:
    def __init__(self, model_path: str = "category_model.pkl"):
        self.model_path = model_path
        self.model = None
        self.vectorizer = None
        self.categories = []
        self.accuracy = 0.0
        self.load_model()
    
    def load_model(self):
        if os.path.exists(self.model_path) and LIBRARIES['sklearn']:
            try:
                self.model = joblib.load(self.model_path)
                self.categories = self.model.classes_ if hasattr(self.model, 'classes_') else []
                logger.info(f"ML-модель загружена, категорий: {len(self.categories)}")
                return
            except Exception as e:
                logger.warning(f"Ошибка загрузки модели: {e}")
        
        self._train_model()
    
    def _train_model(self):
        if not LIBRARIES['sklearn']:
            return
        
        try:
            X = []
            y = []
            
            for category, keywords in CONFIG.category_keywords.items():
                for keyword in keywords:
                    X.append(keyword)
                    y.append(category)
                
                for keyword in keywords:
                    X.append(keyword + " " + category.lower())
                    y.append(category)
                    X.append(category.lower() + " " + keyword)
                    y.append(category)
            
            if X:
                X_train, X_test, y_train, y_test = train_test_split(
                    X, y, test_size=0.2, random_state=42
                )
                
                self.model = Pipeline([
                    ('tfidf', TfidfVectorizer(max_features=2000, ngram_range=(1, 2))),
                    ('clf', MultinomialNB(alpha=0.1))
                ])
                
                self.model.fit(X_train, y_train)
                self.categories = self.model.classes_
                
                y_pred = self.model.predict(X_test)
                self.accuracy = accuracy_score(y_test, y_pred)
                
                joblib.dump(self.model, self.model_path)
                logger.info(f"ML-модель обучена на {len(X)} примерах, точность: {self.accuracy:.2%}")
        except Exception as e:
            logger.error(f"Ошибка обучения модели: {e}")
            self.model = None
    
    def predict(self, name: str) -> Tuple[str, float]:
        if not self.model or not name or not LIBRARIES['sklearn']:
            return "Прочее", 0.0
        
        try:
            pred = self.model.predict([name])[0]
            probs = self.model.predict_proba([name])[0]
            confidence = max(probs) * 100
            
            if confidence < 30:
                return "Прочее", confidence
            
            return pred, confidence
        except Exception as e:
            logger.error(f"Ошибка предсказания: {e}")
            return "Прочее", 0.0
    
    def predict_batch(self, names: List[str]) -> List[Tuple[str, float]]:
        if not self.model or not names or not LIBRARIES['sklearn']:
            return [("Прочее", 0.0) for _ in names]
        
        try:
            predictions = self.model.predict(names)
            probabilities = self.model.predict_proba(names)
            
            results = []
            for pred, probs in zip(predictions, probabilities):
                confidence = max(probs) * 100
                if confidence < 30:
                    results.append(("Прочее", confidence))
                else:
                    results.append((pred, confidence))
            return results
        except Exception as e:
            logger.error(f"Ошибка пакетного предсказания: {e}")
            return [("Прочее", 0.0) for _ in names]

# --------------------------------------------
# 📊 КЭШИРОВАНИЕ API
# --------------------------------------------
class APICache:
    def __init__(self):
        self.cache = CacheManager("api_cache")
    
    def get(self, key: str) -> Optional[Any]:
        return self.cache.get(key)
    
    def set(self, key: str, value: Any):
        self.cache.set(key, value)
    
    def clear(self):
        self.cache.clear()
    
    def get_stats(self) -> Dict:
        return self.cache.get_stats()

# --------------------------------------------
# 🤖 ИИ-РЕДАКТИРОВАНИЕ ДАННЫХ
# --------------------------------------------
class AIFileEditor:
    def __init__(self, api_key: str = None):
        self.api_key = api_key or os.getenv('DEEPSEEK_API_KEY', '')
        self.api_url = "https://api.deepseek.com/v1/chat/completions"
        self.cache = CacheManager("ai_edit_cache")
        
    def edit_data(self, df: pd.DataFrame, prompt: str, instructions: str = "") -> Tuple[pd.DataFrame, str]:
        if not self.api_key:
            return df, "❌ API ключ не установлен"
        
        if not LIBRARIES['openai']:
            return df, "❌ Библиотека openai не установлена"
        
        try:
            sample_data = df.head(20).to_dict(orient='records')
            columns_info = df.dtypes.to_dict()
            
            system_prompt = """Ты - эксперт по обработке данных и юнит-экономике. 
Твоя задача - исправить и улучшить данные в файле согласно запросу пользователя.

Правила:
1. Исправляй только те данные, которые явно указаны в запросе
2. Сохраняй структуру данных
3. Если нужно пересчитать формулы - делай это корректно
4. Возвращай результат в формате JSON с исправленными данными
5. Если данные не требуют исправления, верни их как есть"""

            user_prompt = f"""
Запрос пользователя: {prompt}

Дополнительные инструкции: {instructions}

Данные (первые 20 строк):
{json.dumps(sample_data, ensure_ascii=False, indent=2, default=str)}

Типы данных колонок:
{json.dumps({k: str(v) for k, v in columns_info.items()}, ensure_ascii=False, indent=2)}

Пожалуйста, исправь данные согласно запросу и верни их в формате JSON.
Если нужно пересчитать какие-то значения - сделай это.
Верни только JSON с исправленными данными.
"""

            import openai
            client = openai.OpenAI(api_key=self.api_key)
            
            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.3,
                max_tokens=4000,
                response_format={"type": "json_object"}
            )
            
            content = response.choices[0].message.content
            
            json_match = re.search(r'\{.*\}', content, re.DOTALL)
            if not json_match:
                return df, "❌ Не удалось извлечь JSON из ответа ИИ"
            
            result_data = json.loads(json_match.group())
            
            if 'data' in result_data:
                fixed_data = result_data['data']
            else:
                fixed_data = result_data
            
            if isinstance(fixed_data, list) and len(fixed_data) > 0:
                new_df = pd.DataFrame(fixed_data)
                
                if len(new_df) < len(df):
                    remaining = df.iloc[len(new_df):].copy()
                    for col in remaining.columns:
                        if col not in new_df.columns:
                            new_df[col] = None
                    new_df = pd.concat([new_df, remaining], ignore_index=True)
                
                for col in df.columns:
                    if col not in new_df.columns:
                        new_df[col] = None
                
                new_df = new_df[df.columns]
                
                message = result_data.get('message', '✅ Данные успешно исправлены')
                return new_df, f"✅ {message}"
            
            return df, "❌ Не удалось обработать ответ ИИ"
            
        except Exception as e:
            logger.error(f"AI edit error: {e}")
            return df, f"❌ Ошибка: {str(e)}"
    
    def edit_formulas(self, df: pd.DataFrame, formula_prompt: str) -> Tuple[pd.DataFrame, str]:
        if not self.api_key:
            return df, "❌ API ключ не установлен"
        
        if not LIBRARIES['openai']:
            return df, "❌ Библиотека openai не установлена"
        
        try:
            sample_data = df.head(10).to_dict(orient='records')
            columns_info = df.dtypes.to_dict()
            
            numeric_cols = [col for col, dtype in df.dtypes.items() if dtype in ['int64', 'float64']]
            
            system_prompt = """Ты - эксперт по юнит-экономике и Excel-формулам.
Твоя задача - создать или исправить формулы для расчета показателей.

Правила:
1. Создавай формулы для расчета: маржинальности, прибыли, рентабельности, окупаемости
2. Используй только существующие колонки
3. Добавляй новые колонки с формулами
4. Возвращай результат в формате JSON с исправленными данными
5. Объясни, какие формулы были добавлены или исправлены"""

            user_prompt = f"""
Запрос: {formula_prompt}

Данные (первые 10 строк):
{json.dumps(sample_data, ensure_ascii=False, indent=2, default=str)}

Числовые колонки: {numeric_cols}

Типы данных колонок:
{json.dumps({k: str(v) for k, v in columns_info.items()}, ensure_ascii=False, indent=2)}

Пожалуйста, создай или исправь формулы согласно запросу.
Верни JSON с:
1. 'data' - исправленные данные
2. 'message' - пояснение о проделанной работе
3. 'formulas' - список добавленных формул
"""

            import openai
            client = openai.OpenAI(api_key=self.api_key)
            
            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.2,
                max_tokens=4000,
                response_format={"type": "json_object"}
            )
            
            content = response.choices[0].message.content
            
            json_match = re.search(r'\{.*\}', content, re.DOTALL)
            if not json_match:
                return df, "❌ Не удалось извлечь JSON из ответа ИИ"
            
            result_data = json.loads(json_match.group())
            
            if 'data' in result_data:
                fixed_data = result_data['data']
            else:
                fixed_data = result_data
            
            if isinstance(fixed_data, list) and len(fixed_data) > 0:
                new_df = pd.DataFrame(fixed_data)
                
                if len(new_df) < len(df):
                    remaining = df.iloc[len(new_df):].copy()
                    for col in remaining.columns:
                        if col not in new_df.columns:
                            new_df[col] = None
                    new_df = pd.concat([new_df, remaining], ignore_index=True)
                
                for col in df.columns:
                    if col not in new_df.columns:
                        new_df[col] = None
                
                new_df = new_df[df.columns]
                
                formulas = result_data.get('formulas', [])
                formula_msg = "\n".join([f"- {f}" for f in formulas]) if formulas else ""
                
                message = result_data.get('message', '✅ Формулы успешно применены')
                return new_df, f"✅ {message}\n\n📊 Формулы:\n{formula_msg}"
            
            return df, "❌ Не удалось обработать ответ ИИ"
            
        except Exception as e:
            logger.error(f"AI formula edit error: {e}")
            return df, f"❌ Ошибка: {str(e)}"
    
    def analyze_and_fix_errors(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, str]:
        if not self.api_key:
            return df, "❌ API ключ не установлен"
        
        if not LIBRARIES['openai']:
            return df, "❌ Библиотека openai не установлена"
        
        try:
            sample_data = df.head(20).to_dict(orient='records')
            columns_info = df.dtypes.to_dict()
            
            errors = []
            for col in df.columns:
                if df[col].dtype in ['int64', 'float64']:
                    if df[col].min() < 0 and 'цена' not in col.lower() and 'cost' not in col.lower():
                        errors.append(f"Отрицательные значения в {col}")
                    if df[col].max() > 1000000:
                        errors.append(f"Аномально большие значения в {col}")
                    if df[col].isna().sum() > len(df) * 0.5:
                        errors.append(f"Много пропусков в {col}")
            
            system_prompt = """Ты - эксперт по очистке и исправлению данных.
Твоя задача - найти и исправить ошибки в данных.

Правила:
1. Исправляй все найденные ошибки
2. Заполняй пропуски разумными значениями
3. Исправляй выбросы
4. Нормализуй формат данных
5. Возвращай результат в формате JSON"""

            user_prompt = f"""
Найденные ошибки: {json.dumps(errors, ensure_ascii=False)}

Данные (первые 20 строк):
{json.dumps(sample_data, ensure_ascii=False, indent=2, default=str)}

Типы данных колонок:
{json.dumps({k: str(v) for k, v in columns_info.items()}, ensure_ascii=False, indent=2)}

Пожалуйста, исправь все ошибки и верни JSON с исправленными данными.
"""

            import openai
            client = openai.OpenAI(api_key=self.api_key)
            
            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.1,
                max_tokens=4000,
                response_format={"type": "json_object"}
            )
            
            content = response.choices[0].message.content
            
            json_match = re.search(r'\{.*\}', content, re.DOTALL)
            if not json_match:
                return df, "❌ Не удалось извлечь JSON из ответа ИИ"
            
            result_data = json.loads(json_match.group())
            
            if 'data' in result_data:
                fixed_data = result_data['data']
            else:
                fixed_data = result_data
            
            if isinstance(fixed_data, list) and len(fixed_data) > 0:
                new_df = pd.DataFrame(fixed_data)
                
                if len(new_df) < len(df):
                    remaining = df.iloc[len(new_df):].copy()
                    for col in remaining.columns:
                        if col not in new_df.columns:
                            new_df[col] = None
                    new_df = pd.concat([new_df, remaining], ignore_index=True)
                
                for col in df.columns:
                    if col not in new_df.columns:
                        new_df[col] = None
                
                new_df = new_df[df.columns]
                
                message = result_data.get('message', '✅ Ошибки исправлены')
                return new_df, f"✅ {message}"
            
            return df, "❌ Не удалось обработать ответ ИИ"
            
        except Exception as e:
            logger.error(f"AI error analysis error: {e}")
            return df, f"❌ Ошибка: {str(e)}"

# --------------------------------------------
# 🔌 API-КЛИЕНТЫ ДЛЯ МАРКЕТПЛЕЙСОВ
# --------------------------------------------
class BaseMarketplaceAPI:
    def __init__(self, api_key: str = None, client_id: str = None):
        self.api_key = api_key
        self.client_id = client_id
        self.cache = APICache()
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (compatible; UnitEconomy/1.0)'
        })
        
        if CONFIG.proxy["enabled"]:
            proxies = {}
            if CONFIG.proxy["http"]:
                proxies['http'] = CONFIG.proxy["http"]
            if CONFIG.proxy["https"]:
                proxies['https'] = CONFIG.proxy["https"]
            if proxies:
                self.session.proxies.update(proxies)
    
    def _request(self, method: str, url: str, **kwargs) -> Optional[Dict]:
        for attempt in range(CONFIG.api["max_retries"]):
            try:
                response = self.session.request(
                    method, url,
                    timeout=CONFIG.api["timeout"],
                    **kwargs
                )
                if response.status_code == 200:
                    return response.json()
                elif response.status_code == 429:
                    wait_time = 2 ** attempt
                    logger.warning(f"Rate limit, ждем {wait_time}с")
                    time.sleep(wait_time)
                else:
                    logger.warning(f"API error {response.status_code}: {response.text[:200]}")
                    break
            except requests.exceptions.Timeout:
                logger.warning(f"Таймаут запроса (попытка {attempt+1})")
                time.sleep(1)
            except Exception as e:
                logger.error(f"Request error (attempt {attempt+1}): {e}")
                time.sleep(1)
        
        return None


class YandexMarketAPI(BaseMarketplaceAPI):
    BASE_URL = "https://api.partner.market.yandex.ru/v2"
    
    def __init__(self, api_key: str = None, business_id: str = None):
        super().__init__(api_key)
        self.business_id = business_id
        self.session.headers.update({
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json'
        })
    
    def get_products(self, page: int = 1, page_size: int = 100) -> Optional[Dict]:
        cache_key = generate_cache_key('yandex_products', page, page_size)
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        url = f"{self.BASE_URL}/businesses/{self.business_id}/offer-mappings"
        params = {"page": page, "page_size": page_size}
        
        with timer("Yandex get_products"):
            result = self._request('GET', url, params=params)
        
        if result:
            self.cache.set(cache_key, result)
        return result
    
    def get_offer_prices(self, offer_ids: List[str]) -> Optional[Dict]:
        cache_key = generate_cache_key('yandex_prices', *sorted(offer_ids))
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        url = f"{self.BASE_URL}/businesses/{self.business_id}/offer-prices"
        data = {"offerIds": offer_ids}
        
        with timer("Yandex get_offer_prices"):
            result = self._request('POST', url, json=data)
        
        if result:
            self.cache.set(cache_key, result)
        return result
    
    def update_price(self, offer_id: str, price: float) -> Optional[Dict]:
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
    
    def update_stock(self, offer_id: str, stock: int) -> Optional[Dict]:
        url = f"{self.BASE_URL}/businesses/{self.business_id}/stocks"
        data = {
            "skus": [{
                "sku": offer_id,
                "stock": stock
            }]
        }
        return self._request('PUT', url, json=data)
    
    def create_product(self, product_data: Dict) -> Optional[Dict]:
        url = f"{self.BASE_URL}/businesses/{self.business_id}/offer-mappings"
        data = {
            "offerMappings": [{
                "offer": {
                    "offerId": product_data.get('offer_id', ''),
                    "name": product_data.get('name', ''),
                    "price": product_data.get('price', 0),
                    "category": product_data.get('category', ''),
                    "barcode": product_data.get('barcode', '')
                }
            }]
        }
        return self._request('POST', url, json=data)


class OzonAPI(BaseMarketplaceAPI):
    BASE_URL = "https://api-seller.ozon.ru/v2"
    
    def __init__(self, api_key: str = None, client_id: str = None):
        super().__init__(api_key, client_id)
        self.session.headers.update({
            'Api-Key': api_key,
            'Client-Id': client_id,
            'Content-Type': 'application/json'
        })
    
    def get_products(self, page: int = 1, page_size: int = 100) -> Optional[Dict]:
        cache_key = generate_cache_key('ozon_products', page, page_size)
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        url = f"{self.BASE_URL}/product/list"
        data = {"page": page, "page_size": page_size}
        
        with timer("Ozon get_products"):
            result = self._request('POST', url, json=data)
        
        if result:
            self.cache.set(cache_key, result)
        return result
    
    def get_prices(self, product_ids: List[str]) -> Optional[Dict]:
        cache_key = generate_cache_key('ozon_prices', *sorted(product_ids))
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        url = f"{self.BASE_URL}/product/prices"
        data = {"product_ids": product_ids}
        
        with timer("Ozon get_prices"):
            result = self._request('POST', url, json=data)
        
        if result:
            self.cache.set(cache_key, result)
        return result
    
    def update_price(self, product_id: str, price: float) -> Optional[Dict]:
        url = f"{self.BASE_URL}/product/price"
        data = {
            "product_id": product_id,
            "price": price
        }
        return self._request('POST', url, json=data)
    
    def update_stock(self, product_id: str, stock: int) -> Optional[Dict]:
        url = f"{self.BASE_URL}/product/stocks"
        data = {
            "stocks": [{
                "product_id": product_id,
                "stock": stock
            }]
        }
        return self._request('POST', url, json=data)
    
    def create_product(self, product_data: Dict) -> Optional[Dict]:
        url = f"{self.BASE_URL}/product"
        data = {
            "name": product_data.get('name', ''),
            "price": product_data.get('price', 0),
            "category": product_data.get('category', ''),
            "offer_id": product_data.get('offer_id', ''),
            "barcode": product_data.get('barcode', '')
        }
        return self._request('POST', url, json=data)


class WildberriesAPI(BaseMarketplaceAPI):
    BASE_URL = "https://suppliers-api.wildberries.ru"
    
    def __init__(self, api_key: str = None):
        super().__init__(api_key)
        self.session.headers.update({
            'Authorization': api_key,
            'Content-Type': 'application/json'
        })
    
    def get_products(self, page: int = 1, limit: int = 100) -> Optional[Dict]:
        cache_key = generate_cache_key('wb_products', page, limit)
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        url = f"{self.BASE_URL}/content/v2/get/cards/list"
        data = {"limit": limit, "offset": (page - 1) * limit}
        
        with timer("WB get_products"):
            result = self._request('POST', url, json=data)
        
        if result:
            self.cache.set(cache_key, result)
        return result
    
    def get_prices(self, nm_ids: List[int]) -> Optional[Dict]:
        cache_key = generate_cache_key('wb_prices', *sorted(nm_ids))
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
        url = f"{self.BASE_URL}/public/v1/prices"
        data = {"nm_ids": nm_ids}
        
        with timer("WB get_prices"):
            result = self._request('POST', url, json=data)
        
        if result:
            self.cache.set(cache_key, result)
        return result
    
    def update_price(self, nm_id: int, price: float) -> Optional[Dict]:
        url = f"{self.BASE_URL}/public/v1/prices"
        data = [{"nm_id": nm_id, "price": price}]
        return self._request('POST', url, json=data)
    
    def update_stock(self, nm_id: int, stock: int) -> Optional[Dict]:
        url = f"{self.BASE_URL}/public/v1/stocks"
        data = {"nm_id": nm_id, "stock": stock}
        return self._request('POST', url, json=data)
    
    def create_product(self, product_data: Dict) -> Optional[Dict]:
        url = f"{self.BASE_URL}/content/v2/create/card"
        data = {
            "nm_id": product_data.get('offer_id', ''),
            "name": product_data.get('name', ''),
            "price": product_data.get('price', 0),
            "category": product_data.get('category', ''),
            "barcode": product_data.get('barcode', '')
        }
        return self._request('POST', url, json=data)

# --------------------------------------------
# 🗄️ БАЗА ДАННЫХ OE НОМЕРОВ
# --------------------------------------------
class OEMDatabase:
    def __init__(self, db_path: str = "oem_database.json"):
        self.db_path = db_path
        self.data = {}
        self.cache = {}
        self._load_database()
    
    def _load_database(self):
        if os.path.exists(self.db_path):
            try:
                with open(self.db_path, 'r', encoding='utf-8') as f:
                    self.data = json.load(f)
                logger.info(f"Загружено {len(self.data)} OE номеров из базы")
            except Exception as e:
                logger.error(f"Ошибка загрузки базы OE: {e}")
                self._create_demo_database()
        else:
            self._create_demo_database()
            self._save_database()
    
    def _create_demo_database(self):
        self.data = {}
        
        demo_data = {
            "0986AF0059": {
                "category": "Фильтры",
                "subcategory": "Масляные фильтры",
                "brand": "BOSCH",
                "compatibility": ["BMW 3", "BMW 5", "Audi A4", "Audi A6", "VW Passat"],
                "weight": 0.35,
                "dimensions": {"length": 100, "width": 80, "height": 50},
                "cross_reference": ["MANN W842/2", "MAHLE OC 205"],
                "barcode": "1234567890123",
                "description": "Масляный фильтр BOSCH для европейских автомобилей"
            },
            "W842/2": {
                "category": "Фильтры",
                "subcategory": "Масляные фильтры",
                "brand": "MANN",
                "compatibility": ["BMW 3", "BMW 5", "Audi A4", "Audi A6", "VW Passat"],
                "weight": 0.32,
                "dimensions": {"length": 95, "width": 75, "height": 48},
                "cross_reference": ["BOSCH 0986AF0059", "MAHLE OC 205"],
                "barcode": "1234567890124",
                "description": "Масляный фильтр MANN для европейских автомобилей"
            },
            "OC205": {
                "category": "Фильтры",
                "subcategory": "Масляные фильтры",
                "brand": "MAHLE",
                "compatibility": ["BMW 3", "BMW 5", "Audi A4", "Audi A6", "VW Passat"],
                "weight": 0.33,
                "dimensions": {"length": 98, "width": 78, "height": 49},
                "cross_reference": ["BOSCH 0986AF0059", "MANN W842/2"],
                "barcode": "1234567890125",
                "description": "Масляный фильтр MAHLE для европейских автомобилей"
            },
            "AB123456789": {
                "category": "Тормозная система",
                "subcategory": "Тормозные колодки",
                "brand": "BREMBO",
                "compatibility": ["VW Golf", "VW Passat", "Skoda Octavia", "Audi A3"],
                "weight": 1.2,
                "dimensions": {"length": 150, "width": 120, "height": 30},
                "cross_reference": ["BREMBO P85012", "ATE 13.0460-1234.2"],
                "barcode": "1234567890126",
                "description": "Тормозные колодки BREMBO для VAG группы"
            },
            "5524": {
                "category": "Свечи зажигания",
                "subcategory": "Свечи зажигания",
                "brand": "NGK",
                "compatibility": ["BMW 3", "BMW 5", "Audi A4", "VW Golf", "Toyota Camry"],
                "weight": 0.05,
                "dimensions": {"length": 50, "width": 20, "height": 20},
                "cross_reference": ["BOSCH FR7DC", "DENSO K16PR-U11"],
                "barcode": "1234567890127",
                "description": "Свеча зажигания NGK для бензиновых двигателей"
            }
        }
        
        self.data.update(demo_data)
        
        categories = ["Двигатель", "Трансмиссия", "Подвеска", "Электрооборудование", 
                     "Система охлаждения", "Масла и жидкости"]
        brands = ["BOSCH", "DENSO", "NGK", "BREMBO", "AISIN", "HITACHI", "VALEO", 
                 "MANN", "MAHLE", "SKF", "FAG", "TIMKEN", "CONTINENTAL"]
        
        for i in range(70):
            oe = f"OE{random.randint(10000, 99999)}{chr(65+random.randint(0, 25))}"
            category = random.choice(categories)
            brand = random.choice(brands)
            self.data[oe] = {
                "category": category,
                "subcategory": "Запчасть",
                "brand": brand,
                "compatibility": random.sample(["BMW 3", "Audi A4", "VW Golf", "Toyota Camry", "Honda Civic"], 
                                             random.randint(1, 3)),
                "weight": round(random.uniform(0.1, 10), 2),
                "dimensions": {
                    "length": random.randint(50, 500),
                    "width": random.randint(20, 400),
                    "height": random.randint(10, 200)
                },
                "barcode": f"{random.randint(1000000000000, 9999999999999)}",
                "description": f"{category} {brand} для европейских автомобилей"
            }
        
        logger.info(f"Создана демо-база OE: {len(self.data)} записей")
    
    def _save_database(self):
        try:
            with open(self.db_path, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
            logger.info(f"База OE сохранена: {len(self.data)} записей")
        except Exception as e:
            logger.error(f"Ошибка сохранения базы OE: {e}")
    
    def get_by_oe(self, oe_number: str) -> Optional[Dict]:
        if not oe_number:
            return None
        
        oe_number = oe_number.strip().upper()
        
        if oe_number in self.cache:
            return self.cache[oe_number].copy()
        
        if oe_number in self.data:
            self.cache[oe_number] = self.data[oe_number].copy()
            return self.cache[oe_number].copy()
        
        for key, value in self.data.items():
            if oe_number in key or key in oe_number:
                self.cache[oe_number] = value.copy()
                return self.cache[oe_number].copy()
        
        return None
    
    def get_category(self, oe_number: str) -> str:
        data = self.get_by_oe(oe_number)
        return data.get("category", "Прочее") if data else "Прочее"
    
    def get_brand(self, oe_number: str) -> Optional[str]:
        data = self.get_by_oe(oe_number)
        return data.get("brand") if data else None
    
    def get_dimensions(self, oe_number: str) -> Dict:
        data = self.get_by_oe(oe_number)
        return data.get("dimensions", {}) if data else {}
    
    def get_weight(self, oe_number: str) -> float:
        data = self.get_by_oe(oe_number)
        return data.get("weight", 0) if data else 0
    
    def get_compatibility(self, oe_number: str) -> List[str]:
        data = self.get_by_oe(oe_number)
        return data.get("compatibility", []) if data else []
    
    def get_barcode(self, oe_number: str) -> str:
        data = self.get_by_oe(oe_number)
        return data.get("barcode", "") if data else ""
    
    def get_description(self, oe_number: str) -> str:
        data = self.get_by_oe(oe_number)
        return data.get("description", "") if data else ""
    
    def add_or_update(self, oe_number: str, data: Dict):
        oe_number = oe_number.strip().upper()
        self.data[oe_number] = data
        self.cache[oe_number] = data.copy()
        self._save_database()
    
    def search_by_brand(self, brand: str) -> List[Dict]:
        results = []
        for oe, data in self.data.items():
            if data.get("brand", "").upper() == brand.upper():
                result = data.copy()
                result["oe_number"] = oe
                results.append(result)
        return results
    
    def search_by_category(self, category: str) -> List[Dict]:
        results = []
        for oe, data in self.data.items():
            if data.get("category", "").lower() == category.lower():
                result = data.copy()
                result["oe_number"] = oe
                results.append(result)
        return results
    
    def search_by_compatibility(self, query: str) -> List[Dict]:
        results = []
        query_lower = query.lower()
        for oe, data in self.data.items():
            compatibility = data.get("compatibility", [])
            if any(query_lower in c.lower() for c in compatibility):
                result = data.copy()
                result["oe_number"] = oe
                results.append(result)
        return results
    
    def get_statistics(self) -> Dict:
        stats = {
            "total": len(self.data),
            "categories": {},
            "brands": {},
            "with_barcode": 0,
            "with_dimensions": 0,
            "with_description": 0
        }
        
        for data in self.data.values():
            category = data.get("category", "Прочее")
            stats["categories"][category] = stats["categories"].get(category, 0) + 1
            
            brand = data.get("brand", "Неизвестно")
            stats["brands"][brand] = stats["brands"].get(brand, 0) + 1
            
            if data.get("barcode"):
                stats["with_barcode"] += 1
            if data.get("dimensions"):
                stats["with_dimensions"] += 1
            if data.get("description"):
                stats["with_description"] += 1
        
        return stats

# --------------------------------------------
# 🕷️ ПАРСЕРЫ ЦЕН КОНКУРЕНТОВ
# --------------------------------------------
class CompetitorParser:
    def __init__(self):
        self.cache = APICache()
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
        })
        self.retry_count = 3
        self.timeout = 15
        self.progress_callback = None
        
        if CONFIG.proxy["enabled"]:
            proxies = {}
            if CONFIG.proxy["http"]:
                proxies['http'] = CONFIG.proxy["http"]
            if CONFIG.proxy["https"]:
                proxies['https'] = CONFIG.proxy["https"]
            if proxies:
                self.session.proxies.update(proxies)
    
    def set_progress_callback(self, callback):
        self.progress_callback = callback
    
    def _update_progress(self, current: int, total: int, message: str = ""):
        if self.progress_callback:
            self.progress_callback(current, total, message)
    
    @st.cache_data(ttl=300)
    def parse_yandex_market(_self, query: str, max_pages: int = 2) -> List[Dict]:
        results = []
        
        for page in range(1, max_pages + 1):
            for attempt in range(_self.retry_count):
                try:
                    url = "https://market.yandex.ru/search"
                    params = {
                        'text': query,
                        'page': page,
                        'rs': 'eJwzrLDM1QAAGm0B_Q'
                    }
                    
                    time.sleep(random.uniform(1, 3))
                    
                    response = _self.session.get(url, params=params, timeout=_self.timeout)
                    
                    if response.status_code != 200:
                        if response.status_code == 429:
                            time.sleep(5)
                            continue
                        break
                    
                    html = response.text
                    
                    json_match = re.search(r'window\.__initialState\s*=\s*({.*?});', html, re.DOTALL)
                    if json_match:
                        try:
                            data = json.loads(json_match.group(1))
                            products = _self._extract_yandex_products_from_json(data)
                            for product in products[:10]:
                                if product.get('price', 0) > 0:
                                    results.append({
                                        'marketplace': 'Яндекс Маркет',
                                        'offer_id': product.get('id', ''),
                                        'name': product.get('name', ''),
                                        'price': product.get('price', 0),
                                        'url': product.get('url', ''),
                                        'parsed_at': datetime.now().isoformat()
                                    })
                        except:
                            pass
                    
                    if not results or len(results) < 5:
                        product_pattern = r'<article[^>]*data-zone-name="[^"]*product[^"]*"[^>]*>.*?<h3[^>]*>.*?<a[^>]*href="([^"]+)"[^>]*>(.*?)</a>.*?</h3>.*?<span[^>]*class="[^"]*price[^"]*"[^>]*>([0-9\s]+)</span>'
                        matches = re.findall(product_pattern, html, re.DOTALL)
                        
                        for url, name, price in matches[:10]:
                            price_clean = re.sub(r'[^\d]', '', price)
                            if price_clean:
                                results.append({
                                    'marketplace': 'Яндекс Маркет',
                                    'offer_id': url.split('/')[-1] if '/' in url else '',
                                    'name': re.sub(r'<[^>]+>', '', name).strip(),
                                    'price': float(price_clean),
                                    'url': 'https://market.yandex.ru' + url if url.startswith('/') else url,
                                    'parsed_at': datetime.now().isoformat()
                                })
                    
                    if not results:
                        results = _self._generate_demo_results("Яндекс Маркет", query, 5)
                    
                    break
                    
                except Exception as e:
                    logger.error(f"Yandex parse error (attempt {attempt+1}): {e}")
                    time.sleep(2)
            
            if len(results) >= 20:
                break
        
        return results
    
    def _extract_yandex_products_from_json(self, data: Dict) -> List[Dict]:
        products = []
        
        try:
            search_results = data.get('search', {}).get('results', [])
            if not search_results:
                search_results = data.get('searchResults', {}).get('items', [])
            
            for item in search_results:
                if isinstance(item, dict):
                    product = {
                        'id': item.get('id', ''),
                        'name': item.get('name', ''),
                        'price': item.get('price', {}).get('value', 0),
                        'url': item.get('url', '')
                    }
                    if product.get('price', 0) > 0:
                        products.append(product)
        except:
            pass
        
        return products
    
    @st.cache_data(ttl=300)
    def parse_ozon(_self, query: str, max_pages: int = 2) -> List[Dict]:
        results = []
        
        for page in range(1, max_pages + 1):
            for attempt in range(_self.retry_count):
                try:
                    url = f"https://www.ozon.ru/search/"
                    params = {
                        'text': query,
                        'page': page,
                        'sorting': 'relevance'
                    }
                    
                    time.sleep(random.uniform(1.5, 3))
                    
                    response = _self.session.get(url, params=params, timeout=_self.timeout)
                    
                    if response.status_code != 200:
                        if response.status_code == 429:
                            time.sleep(5)
                            continue
                        break
                    
                    html = response.text
                    
                    json_patterns = [
                        r'<script[^>]*id="__NEXT_DATA__"[^>]*>(.*?)</script>',
                        r'<script[^>]*type="application/json"[^>]*>(.*?)</script>',
                        r'window\.__INITIAL_STATE__\s*=\s*({.*?});'
                    ]
                    
                    for pattern in json_patterns:
                        match = re.search(pattern, html, re.DOTALL)
                        if match:
                            try:
                                data = json.loads(match.group(1))
                                products = _self._extract_ozon_products_from_json(data)
                                for product in products[:10]:
                                    if product.get('price', 0) > 0:
                                        results.append({
                                            'marketplace': 'Ozon',
                                            'product_id': product.get('id', ''),
                                            'name': product.get('name', ''),
                                            'price': product.get('price', 0),
                                            'url': product.get('url', ''),
                                            'parsed_at': datetime.now().isoformat()
                                        })
                                break
                            except:
                                continue
                    
                    if not results:
                        product_pattern = r'<div[^>]*data-widget="[^"]*searchResultsV2[^"]*"[^>]*>.*?<a[^>]*href="([^"]+)"[^>]*>.*?<span[^>]*class="[^"]*tsBody500Medium[^"]*"[^>]*>(.*?)</span>.*?<span[^>]*class="[^"]*tsBodyL[^"]*"[^>]*>([0-9\s]+)</span>'
                        matches = re.findall(product_pattern, html, re.DOTALL)
                        
                        for url, name, price in matches[:10]:
                            price_clean = re.sub(r'[^\d]', '', price)
                            if price_clean:
                                results.append({
                                    'marketplace': 'Ozon',
                                    'product_id': url.split('/')[-1] if '/' in url else '',
                                    'name': re.sub(r'<[^>]+>', '', name).strip(),
                                    'price': float(price_clean),
                                    'url': 'https://www.ozon.ru' + url if url.startswith('/') else url,
                                    'parsed_at': datetime.now().isoformat()
                                })
                    
                    if not results:
                        results = _self._generate_demo_results("Ozon", query, 5)
                    
                    break
                    
                except Exception as e:
                    logger.error(f"Ozon parse error (attempt {attempt+1}): {e}")
                    time.sleep(2)
            
            if len(results) >= 20:
                break
        
        return results
    
    def _extract_ozon_products_from_json(self, data: Dict) -> List[Dict]:
        products = []
        
        try:
            items = data.get('props', {}).get('pageProps', {}).get('catalog', {}).get('items', [])
            if not items:
                items = data.get('catalog', {}).get('items', [])
            if not items:
                items = data.get('items', [])
            
            for item in items:
                if isinstance(item, dict):
                    product = {
                        'id': item.get('id', item.get('productId', '')),
                        'name': item.get('name', item.get('title', '')),
                        'price': item.get('price', {}).get('price', 0) or item.get('price', 0),
                        'url': item.get('link', item.get('url', ''))
                    }
                    if product.get('price', 0) > 0:
                        products.append(product)
        except:
            pass
        
        return products
    
    @st.cache_data(ttl=300)
    def parse_wildberries(_self, query: str, max_pages: int = 2) -> List[Dict]:
        results = []
        
        for page in range(1, max_pages + 1):
            for attempt in range(_self.retry_count):
                try:
                    url = "https://search.wb.ru/exactmatch/ru/common/v4/search"
                    params = {
                        'query': query,
                        'page': page,
                        'limit': 50,
                        'currency': 'rub',
                        'appType': 1,
                        'lang': 'ru',
                        'dest': -1257786
                    }
                    
                    time.sleep(random.uniform(1, 2))
                    
                    headers = {
                        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                        'Accept': 'application/json',
                        'Origin': 'https://www.wildberries.ru',
                        'Referer': 'https://www.wildberries.ru/'
                    }
                    
                    response = _self.session.get(url, params=params, headers=headers, timeout=_self.timeout)
                    
                    if response.status_code != 200:
                        if response.status_code == 429:
                            time.sleep(5)
                            continue
                        break
                    
                    data = response.json()
                    products = data.get('data', {}).get('products', [])
                    
                    for product in products[:20]:
                        price = product.get('priceU', 0)
                        if price > 0:
                            results.append({
                                'marketplace': 'Wildberries',
                                'nm_id': product.get('id', ''),
                                'name': product.get('name', ''),
                                'price': price / 100,
                                'url': f"https://www.wildberries.ru/catalog/{product.get('id', '')}/detail.aspx",
                                'parsed_at': datetime.now().isoformat()
                            })
                    
                    if not results:
                        results = _self._generate_demo_results("Wildberries", query, 5)
                    
                    break
                    
                except Exception as e:
                    logger.error(f"WB parse error (attempt {attempt+1}): {e}")
                    time.sleep(2)
            
            if len(results) >= 20:
                break
        
        return results
    
    def _generate_demo_results(self, marketplace: str, query: str, count: int = 5) -> List[Dict]:
        results = []
        
        base_prices = [1500, 2300, 890, 3450, 1200, 5600, 780, 2150, 4300, 990]
        
        for i in range(min(count, len(base_prices))):
            price = base_prices[i] * random.uniform(0.8, 1.2)
            results.append({
                'marketplace': marketplace,
                'name': f"{query} (артикул {random.randint(1000, 9999)})",
                'price': round(price, 2),
                'parsed_at': datetime.now().isoformat(),
                'is_demo': True
            })
        
        return results
    
    def parse_all_marketplaces(self, query: str) -> Dict[str, List[Dict]]:
        results = {}
        
        try:
            results['Яндекс Маркет'] = self.parse_yandex_market(query)
            time.sleep(2)
        except Exception as e:
            logger.error(f"Yandex parse error: {e}")
            results['Яндекс Маркет'] = []
        
        try:
            results['Ozon'] = self.parse_ozon(query)
            time.sleep(2)
        except Exception as e:
            logger.error(f"Ozon parse error: {e}")
            results['Ozon'] = []
        
        try:
            results['Wildberries'] = self.parse_wildberries(query)
        except Exception as e:
            logger.error(f"WB parse error: {e}")
            results['Wildberries'] = []
        
        total_items = sum(len(items) for items in results.values())
        if total_items == 0:
            logger.info(f"Все маркетплейсы вернули 0 результатов для '{query}', генерируем демо")
            for marketplace in ['Яндекс Маркет', 'Ozon', 'Wildberries']:
                results[marketplace] = self._generate_demo_results(marketplace, query, 4)
        
        return results
    
    def parse_multiple_articles(self, articles: List[str], marketplace: str = "Все", max_pages: int = 1) -> Dict[str, Dict[str, List[Dict]]]:
        results = {}
        total = len(articles)
        
        for idx, article in enumerate(articles):
            if not article or not article.strip():
                continue
            
            article = article.strip()
            self._update_progress(idx + 1, total, f"Парсинг артикула: {article}")
            
            results[article] = {}
            
            try:
                if marketplace == "Все" or marketplace == "Яндекс Маркет":
                    results[article]['Яндекс Маркет'] = self.parse_yandex_market(article, max_pages)
                    time.sleep(1)
                
                if marketplace == "Все" or marketplace == "Ozon":
                    results[article]['Ozon'] = self.parse_ozon(article, max_pages)
                    time.sleep(1)
                
                if marketplace == "Все" or marketplace == "Wildberries":
                    results[article]['Wildberries'] = self.parse_wildberries(article, max_pages)
                    time.sleep(1)
                
                total_found = 0
                for mp, items in results[article].items():
                    total_found += len(items)
                
                if total_found == 0:
                    for mp in results[article].keys():
                        results[article][mp] = self._generate_demo_results(mp, article, 3)
                
            except Exception as e:
                logger.error(f"Error parsing article {article}: {e}")
                results[article] = {
                    'error': str(e),
                    'timestamp': datetime.now().isoformat()
                }
        
        self._update_progress(total, total, "✅ Парсинг завершен!")
        return results
    
    def parse_articles_from_file(self, file_bytes: bytes, article_column: str = "Артикул", 
                                  brand_column: str = "Бренд", marketplace: str = "Все", 
                                  max_pages: int = 1) -> Dict:
        try:
            if file_bytes[:4] == b'PK\x03\x04':
                df = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
            else:
                df = pd.read_csv(io.BytesIO(file_bytes), encoding='utf-8-sig')
            
            if df.empty:
                return {"error": "Файл пуст"}
            
            if article_column not in df.columns:
                for col in df.columns:
                    col_lower = col.lower()
                    if any(w in col_lower for w in ['артикул', 'article', 'sku', 'код', 'id']):
                        article_column = col
                        break
                else:
                    return {"error": f"Колонка '{article_column}' не найдена. Доступные: {list(df.columns)}"}
            
            if brand_column not in df.columns:
                for col in df.columns:
                    col_lower = col.lower()
                    if any(w in col_lower for w in ['бренд', 'brand', 'марка', 'производитель']):
                        brand_column = col
                        break
                else:
                    brand_column = None
            
            articles_data = []
            for idx, row in df.iterrows():
                article = safe_str(row[article_column])
                if article and article.strip():
                    brand = safe_str(row[brand_column]) if brand_column else ""
                    articles_data.append({
                        'article': article.strip(),
                        'brand': brand
                    })
            
            if not articles_data:
                return {"error": "Нет артикулов для парсинга"}
            
            results = self.parse_multiple_articles(
                [a['article'] for a in articles_data], 
                marketplace, 
                max_pages
            )
            
            brand_map = {a['article']: a['brand'] for a in articles_data}
            for article, brand in brand_map.items():
                if article in results:
                    results[article]['_brand'] = brand
            
            return results
            
        except Exception as e:
            logger.error(f"Error parsing file: {e}")
            return {"error": str(e)}
    
    def format_multiple_results(self, results: Dict) -> pd.DataFrame:
        rows = []
        
        for article, data in results.items():
            if isinstance(data, dict) and 'error' in data:
                rows.append({
                    'Артикул': article,
                    'Бренд': data.get('_brand', ''),
                    'Маркетплейс': 'Ошибка',
                    'Наименование': '',
                    'Цена': 0,
                    'Статус': f"Ошибка: {data.get('error', '')}",
                    'URL': '',
                    'ID': ''
                })
                continue
            
            brand = data.get('_brand', '')
            
            for marketplace, items in data.items():
                if marketplace == '_brand':
                    continue
                
                if not items:
                    rows.append({
                        'Артикул': article,
                        'Бренд': brand,
                        'Маркетплейс': marketplace,
                        'Наименование': 'Не найдено',
                        'Цена': 0,
                        'Статус': 'Не найдено',
                        'URL': '',
                        'ID': ''
                    })
                else:
                    for item in items[:5]:
                        rows.append({
                            'Артикул': article,
                            'Бренд': brand,
                            'Маркетплейс': marketplace,
                            'Наименование': item.get('name', ''),
                            'Цена': item.get('price', 0),
                            'Статус': 'Демо' if item.get('is_demo', False) else 'Найден',
                            'URL': item.get('url', ''),
                            'ID': item.get('offer_id', item.get('product_id', item.get('nm_id', '')))
                        })
        
        df = pd.DataFrame(rows)
        if not df.empty:
            df = df.sort_values(['Артикул', 'Маркетплейс']).reset_index(drop=True)
        
        return df

# --------------------------------------------
# 🏪 УПРАВЛЕНИЕ КОНКУРЕНТАМИ
# --------------------------------------------
class CompetitorManager:
    def __init__(self):
        self.parser = CompetitorParser()
        self.competitor_data = {}
        self.last_update = {}
        self.cache = CacheManager("competitor_cache")
    
    def get_competitor_prices(self, query: str, marketplace: str = None) -> Dict:
        cache_key = generate_cache_key('competitor_prices', query, marketplace or 'all')
        
        cached = self.cache.get(cache_key)
        if cached:
            return cached
        
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
        
        self.cache.set(cache_key, results)
        return results
    
    def analyze_competitor_prices(self, product_name: str, our_price: float) -> Dict:
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
            'competitors': all_prices[:10]
        }

# --------------------------------------------
# 🧠 AI-ПОЛУЧЕНИЕ ТАРИФОВ
# --------------------------------------------
class AITariffProvider:
    def __init__(self, api_key: str = None, cache_ttl: int = 3600):
        self.api_key = api_key or os.getenv('DEEPSEEK_API_KEY', '')
        self.api_url = "https://api.deepseek.com/v1/chat/completions"
        self.cache = CacheManager("tariff_cache")
        self.cache_ttl = cache_ttl
        self.last_update = {}
        
    def get_rates(self, marketplace: str, mode: str = "FBY") -> Dict:
        cache_key = generate_cache_key(marketplace, mode)
        
        cached = self.cache.get(cache_key)
        if cached:
            logger.info(f"Использую кэшированные тарифы для {marketplace}/{mode}")
            return cached.copy()
        
        rates = self._get_base_rates(marketplace, mode)
        
        if self.api_key and LIBRARIES['openai']:
            try:
                ai_rates = self._get_ai_rates(marketplace, mode)
                if ai_rates and isinstance(ai_rates, dict):
                    for key, value in ai_rates.items():
                        if key in rates and isinstance(value, (int, float)):
                            rates[key] = value
                    logger.info(f"Получены AI-тарифы для {marketplace}/{mode}")
            except Exception as e:
                logger.error(f"AI tariff error: {e}")
        
        self.cache.set(cache_key, rates.copy())
        self.last_update[cache_key] = datetime.now()
        
        return rates
    
    def _get_base_rates(self, marketplace: str, mode: str) -> Dict:
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
            "insurance": 0.005
        }
        
        marketplace_rates = {
            "Яндекс Маркет": {"commission": 0.11},
            "Ozon": {"commission": 0.10},
            "Wildberries": {"commission": 0.12},
            "AliExpress": {"commission": 0.08},
            "Мегамаркет": {"commission": 0.09}
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
    
    def _get_ai_rates(self, marketplace: str, mode: str) -> Optional[Dict]:
        if not self.api_key or not LIBRARIES['openai']:
            return None
        
        try:
            import openai
            client = openai.OpenAI(api_key=self.api_key)
            
            prompt = f"""
            Предоставь актуальные тарифы для маркетплейса {marketplace} 
            для продажи автозапчастей на начало 2026 года.
            Режим работы: {mode}
            
            Верни ТОЛЬКО JSON с тарифами.
            """
            
            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                max_tokens=1000,
                response_format={"type": "json_object"}
            )
            
            content = response.choices[0].message.content
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
    
    def get_all_rates(self) -> Dict:
        all_rates = {}
        for marketplace in CONFIG.marketplaces:
            all_rates[marketplace] = {}
            for mode in CONFIG.operation_modes:
                all_rates[marketplace][mode] = self.get_rates(marketplace, mode)
        return all_rates
    
    def clear_cache(self):
        self.cache.clear()
        self.last_update.clear()
        logger.info("Кэш тарифов очищен")

# --------------------------------------------
# 🏷️ КЛАССИФИКАТОР КАТЕГОРИЙ (с ML)
# --------------------------------------------
class CategoryClassifier:
    def __init__(self):
        self.keywords = CONFIG.category_keywords
        self.categories = list(self.keywords.keys())
        self.cache = {}
        self.oem_patterns = CONFIG.oem_patterns
        self.barcode_patterns = CONFIG.barcode_patterns
        self.ml_classifier = AutoClassifier() if LIBRARIES['sklearn'] else None
        
    @lru_cache(maxsize=10000)
    def classify(self, name: str) -> Tuple[str, float]:
        if not name:
            return "Прочее", 0.0
        
        if self.ml_classifier:
            ml_category, ml_confidence = self.ml_classifier.predict(name)
            if ml_confidence > 50:
                return ml_category, ml_confidence
        
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
    
    def classify_batch(self, names: List[str]) -> List[Tuple[str, float]]:
        if not names:
            return []
        
        results = []
        for name in names:
            results.append(self.classify(name))
        return results
    
    def extract_oem(self, name: str) -> Optional[str]:
        if not name:
            return None
        
        for pattern in self.oem_patterns:
            match = re.search(pattern, name.upper())
            if match:
                return match.group()
        return None
    
    def extract_barcode(self, name: str) -> Optional[str]:
        if not name:
            return None
        
        for pattern in self.barcode_patterns:
            match = re.search(pattern, name)
            if match:
                return match.group()
        return None
    
    def extract_brand(self, name: str) -> Optional[str]:
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

# --------------------------------------------
# 📊 КЛАСС ВАЛИДАЦИИ ДАННЫХ
# --------------------------------------------
class DataValidator:
    def __init__(self):
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
        errors = []
        warnings = []
        
        price = safe_float(row.get("price", 0))
        if price <= 0:
            errors.append("Цена должна быть больше 0")
        elif price < CONFIG.validation["min_price"]:
            warnings.append(f"Цена очень низкая: {price} ₽")
        elif price > CONFIG.validation["max_price"]:
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
            elif value > CONFIG.validation["max_dimension"]:
                warnings.append(f"Подозрительно большая {label.lower()}: {value} мм")
        
        weight = safe_float(row.get("weight", 0))
        if weight < 0:
            errors.append("Вес не может быть отрицательным")
        elif weight > CONFIG.validation["max_weight"]:
            warnings.append(f"Подозрительно большой вес: {weight} кг")
        
        name = safe_str(row.get("name", ""))
        if not name:
            warnings.append("Отсутствует наименование товара")
        
        article = safe_str(row.get("article", ""))
        if not article:
            warnings.append("Отсутствует артикул товара")
        elif not validate_article(article):
            warnings.append(f"Некорректный артикул: {article}")
        
        barcode = safe_str(row.get("barcode", ""))
        if barcode and not is_valid_barcode(barcode):
            warnings.append(f"Неверный формат штрихкода: {barcode}")
        
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
        self.errors = []
        self.warnings = []
        self.stats = {
            "total": 0,
            "valid": 0,
            "invalid": 0,
            "skipped": 0,
            "by_error_type": defaultdict(int)
        }

# --------------------------------------------
# 💎 ДВИЖОК РАСЧЕТА ЮНИТ-ЭКОНОМИКИ
# --------------------------------------------
class UnitEconomicsEngine:
    def __init__(self, 
                 marketplace: str = "Яндекс Маркет", 
                 mode: str = "FBY", 
                 days_storage: int = 30,
                 dimension_unit: str = "мм"):
        
        self.marketplace = marketplace
        self.mode = mode
        self.days_storage = days_storage
        self.dimension_unit = dimension_unit
        self.tariff_provider = AITariffProvider()
        self.rates = self.tariff_provider.get_rates(marketplace, mode)
        self.fixed_costs = 50000.0
        self.avg_orders = 100
        self.retention_rate = 0.7
        self.discount_rate = 0.1
        self.lead_time = 7
        self.z_score = 1.65
        self.validator = DataValidator()
        self.competitor_manager = CompetitorManager()
        self.oem_db = OEMDatabase()
        self.classifier = CategoryClassifier()
        self.cache = CacheManager("engine_cache")
    
    def _convert_to_mm(self, value: float) -> float:
        if value == 0:
            return 0.0
        
        if self.dimension_unit == "мм":
            return value
        elif self.dimension_unit == "см":
            return value * 10.0
        
        return value
    
    def calculate_product(self, row: Dict) -> Optional[Dict]:
        try:
            cache_key = generate_cache_key('product', row.get('article', ''), self.marketplace, self.mode)
            cached = self.cache.get(cache_key)
            if cached:
                return cached.copy()
            
            is_valid, validation_issues = self.validator.validate_product(row)
            result = self._calculate_product_model(row, is_valid, validation_issues)
            
            if result:
                self.cache.set(cache_key, result.copy())
            
            return result
                
        except Exception as e:
            logger.error(f"Error calculating product {row.get('article', 'unknown')}: {e}")
            return None
    
    def _calculate_product_model(self, row: Dict, is_valid: bool, validation_issues: List[str]) -> Dict:
        article = safe_str(row.get("article", ""))
        name = safe_str(row.get("name", ""))
        price = safe_float(row.get("price", 0))
        
        length_orig = safe_float(row.get("length", 0))
        width_orig = safe_float(row.get("width", 0))
        height_orig = safe_float(row.get("height", 0))
        
        length = self._convert_to_mm(length_orig)
        width = self._convert_to_mm(width_orig)
        height = self._convert_to_mm(height_orig)
        
        cost = safe_float(row.get("cost", price * 0.5 if price > 0 else 0))
        weight = safe_float(row.get("weight", 0))
        oe_number = safe_str(row.get("oe_number", ""))
        brand = safe_str(row.get("brand", ""))
        category = safe_str(row.get("category", ""))
        compatibility = safe_str(row.get("compatibility", ""))
        barcode = safe_str(row.get("barcode", ""))
        
        if price <= 0:
            return None
        
        if not category or category == "" or category == "Прочее":
            category, confidence = self.classifier.classify(name)
        
        competitor_analysis = self.competitor_manager.analyze_competitor_prices(name, price)
        
        volume = calculate_volume(length, width, height)
        if volume == 0:
            volume = 1.0
        
        if weight == 0 and volume > 0:
            weight = volume * 0.8
        
        if weight == 0:
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
        
        recommended_price, price_reason = calculate_price_recommendation(
            price, 
            competitor_analysis.get('avg_price', 0),
            margin
        )
        
        if recommended_price > price * 1.05:
            price_action = "Повысить"
        elif recommended_price < price * 0.95:
            price_action = "Снизить"
        else:
            price_action = "Оставить"
        
        return {
            "article": article,
            "name": name,
            "brand": brand,
            "oe_number": oe_number,
            "category": category,
            "compatibility": compatibility,
            "barcode": barcode,
            "price": round(price, 2),
            "cost": round(cost, 2),
            "length_orig": round(length_orig, 1),
            "width_orig": round(width_orig, 1),
            "height_orig": round(height_orig, 1),
            "length_mm": round(length, 1),
            "width_mm": round(width, 1),
            "height_mm": round(height, 1),
            "dimension_unit": self.dimension_unit,
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
            "validation_valid": is_valid,
            "validation_issues": "; ".join(validation_issues[:3]) if validation_issues else "",
        }
    
    def _calculate_logistics(self, volume: float, weight: float, price: float) -> float:
        try:
            if self.mode == "FBY":
                logistics = self.rates.get("fba_base", 70.0) + weight * self.rates.get("fba_per_kg", 12.0)
            elif self.mode == "FBS":
                logistics = self.rates.get("fbs_logistics", 115.0)
                if volume <= 1.0:
                    logistics = self.rates.get("logistics_base", 70.0)
                elif volume <= 160.0:
                    logistics = self.rates.get("logistics_base", 70.0) + (volume - 1.0) * self.rates.get("logistics_per_liter", 22.0)
                else:
                    logistics = self.rates.get("logistics_base", 70.0) + 159.0 * 22.0 + (volume - 160.0) * 2.0
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
        if profit > 1000:
            return "A"
        elif profit > 100:
            return "B"
        return "C"
    
    def _xyz_category(self, sales: float) -> str:
        if sales > 100:
            return "X"
        elif sales > 50:
            return "Y"
        return "Z"
    
    def get_validation_report(self) -> Dict:
        return self.validator.get_report()
    
    def calculate_batch(self, rows: List[Dict], parallel: bool = True) -> List[Dict]:
        if parallel and len(rows) > 10:
            with ThreadPoolExecutor(max_workers=4) as executor:
                futures = [executor.submit(self.calculate_product, row) for row in rows]
                return [f.result() for f in futures if f.result() is not None]
        else:
            return [self.calculate_product(row) for row in rows if self.calculate_product(row) is not None]

# --------------------------------------------
# 📤 АВТОМАТИЧЕСКАЯ ВЫГРУЗКА НА МАРКЕТПЛЕЙСЫ
# --------------------------------------------
class MarketplaceUploader:
    def __init__(self):
        self.api_clients = {}
        self.upload_queue = Queue()
        self.is_running = False
        self.upload_log = []
        self.lock = Lock()
        self.stop_event = Event()
        
    def register_client(self, marketplace: str, client: BaseMarketplaceAPI):
        self.api_clients[marketplace] = client
        
    def upload_prices(self, marketplace: str, products: List[Dict]) -> Dict:
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
    
    def upload_stocks(self, marketplace: str, products: List[Dict]) -> Dict:
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
                card_data = {
                    "offer_id": product.get('article', ''),
                    "name": product.get('name', ''),
                    "price": product.get('price', 0),
                    "category": product.get('category', ''),
                    "brand": product.get('brand', ''),
                    "barcode": product.get('barcode', ''),
                    "description": product.get('description', ''),
                    "images": product.get('images', []),
                    "dimensions": {
                        "length": product.get('length_mm', 0),
                        "width": product.get('width_mm', 0),
                        "height": product.get('height_mm', 0)
                    }
                }
                
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
        with self.lock:
            return self.upload_log.copy()
    
    def clear_log(self):
        with self.lock:
            self.upload_log = []
    
    def start_worker(self):
        if self.is_running:
            return
        
        self.is_running = True
        self.stop_event.clear()
        thread = Thread(target=self._worker_loop)
        thread.daemon = True
        thread.start()
    
    def stop_worker(self):
        self.is_running = False
        self.stop_event.set()
    
    def _worker_loop(self):
        while not self.stop_event.is_set():
            try:
                if not self.upload_queue.empty():
                    task = self.upload_queue.get(timeout=1)
                    marketplace = task.get('marketplace')
                    task_type = task.get('type')
                    products = task.get('products', [])
                    
                    if task_type == 'prices':
                        self.upload_prices(marketplace, products)
                    elif task_type == 'stocks':
                        self.upload_stocks(marketplace, products)
                    elif task_type == 'products':
                        self.upload_products(marketplace, products)
                    
                    self.upload_queue.task_done()
                else:
                    time.sleep(1)
            except Exception as e:
                logger.error(f"Worker error: {e}")
                time.sleep(5)

# --------------------------------------------
# 📤 ЭКСПОРТ В EXCEL
# --------------------------------------------
class ExcelExportEngine:
    def __init__(self):
        self.classifier = CategoryClassifier()
    
    def export(self, results: List[Dict], marketplace: str, mode: str, all_rates: Dict = None) -> bytes:
        output = io.BytesIO()
        
        if not results:
            return output.getvalue()
        
        if not LIBRARIES['openpyxl']:
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
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment
            from openpyxl.utils import get_column_letter
            
            wb = Workbook()
            ws = wb.active
            ws.title = "Юнит-экономика"
            
            header_font = Font(bold=True, color="FFFFFF", size=11)
            header_fill = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
            header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
            headers = self._get_product_headers()
            
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
            
            for row_idx, row_data in enumerate(results, 2):
                for col_idx, header in enumerate(headers, 1):
                    key = self._get_key_mapping().get(header, header.lower())
                    value = row_data.get(key, "")
                    cell = ws.cell(row=row_idx, column=col_idx, value=value)
                    
                    if isinstance(value, (int, float)):
                        if any(w in header for w in ["Цена", "Прибыль", "Доход", "стоимость", "LTV", "CAC"]):
                            cell.number_format = '#,##0.00 ₽'
                        elif any(w in header for w in ["%", "марж", "рр", "ROI", "CM"]):
                            cell.number_format = '0.00%'
                        elif any(w in header for w in ["дн", "запас", "EOQ"]):
                            cell.number_format = '#,##0'
                        
                        if key == "unit_profit" and isinstance(value, (int, float)):
                            if value > 0:
                                cell.font = Font(color="006400")
                            elif value < 0:
                                cell.font = Font(color="8B0000")
            
            for col in range(1, len(headers) + 1):
                col_letter = get_column_letter(col)
                ws.column_dimensions[col_letter].width = 15
            
            if all_rates:
                self._create_help_sheets(wb, all_rates)
            
            self._create_summary_sheet(wb, results, marketplace, mode)
            
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
        return [
            "Артикул", "Наименование", "Бренд", "OE номер", "Категория", "Применимость", "Штрихкод",
            "Цена", "Себестоимость", 
            "Длина (исх.)", "Ширина (исх.)", "Высота (исх.)", "Ед. изм.",
            "Длина (мм)", "Ширина (мм)", "Высота (мм)",
            "Объем (л)", "Вес (кг)",
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
    
    def _get_key_mapping(self) -> Dict[str, str]:
        return {
            "Артикул": "article",
            "Наименование": "name",
            "Бренд": "brand",
            "OE номер": "oe_number",
            "Категория": "category",
            "Применимость": "compatibility",
            "Штрихкод": "barcode",
            "Цена": "price",
            "Себестоимость": "cost",
            "Длина (исх.)": "length_orig",
            "Ширина (исх.)": "width_orig",
            "Высота (исх.)": "height_orig",
            "Ед. изм.": "dimension_unit",
            "Длина (мм)": "length_mm",
            "Ширина (мм)": "width_mm",
            "Высота (мм)": "height_mm",
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
    
    def _create_help_sheets(self, wb, all_rates: Dict):
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
        
        headers = ["Режим", "Комиссия %", "Эквайринг %", "Логистика база", "Хранение за литр", 
                   "Логистика за литр", "Бесплатно дней", "FBA база", "FBA за кг"]
        
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        for marketplace, modes in all_rates.items():
            sheet_name = f"Тарифы {marketplace}"[:31]
            ws = wb.create_sheet(sheet_name)
            
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
            
            row_idx = 2
            for mode, rates in modes.items():
                ws.cell(row=row_idx, column=1, value=mode)
                ws.cell(row=row_idx, column=2, value=rates.get("commission", 0))
                ws.cell(row=row_idx, column=3, value=rates.get("acquiring", 0))
                ws.cell(row=row_idx, column=4, value=rates.get("logistics_base", 0))
                ws.cell(row=row_idx, column=5, value=rates.get("storage_per_liter", 0))
                ws.cell(row=row_idx, column=6, value=rates.get("logistics_per_liter", 0))
                ws.cell(row=row_idx, column=7, value=rates.get("storage_free_days", 0))
                ws.cell(row=row_idx, column=8, value=rates.get("fba_base", 0))
                ws.cell(row=row_idx, column=9, value=rates.get("fba_per_kg", 0))
                
                for col in [2, 3]:
                    cell = ws.cell(row=row_idx, column=col)
                    cell.number_format = '0.00%'
                
                row_idx += 1
            
            column_widths = [12, 14, 14, 16, 16, 16, 14, 14, 14]
            for col, width in enumerate(column_widths, 1):
                ws.column_dimensions[get_column_letter(col)].width = width
    
    def _create_summary_sheet(self, wb, results: List[Dict], marketplace: str, mode: str):
        from openpyxl.styles import Font, PatternFill, Alignment
        
        ws = wb.create_sheet("Сводка")
        
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
        
        headers = ["Показатель", "Значение"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
        
        if results:
            total_profit = sum(r.get("unit_profit", 0) for r in results)
            avg_margin = sum(r.get("margin", 0) for r in results) / len(results) if results else 0
            profitable = sum(1 for r in results if r.get("unit_profit", 0) > 0)
            
            summary_data = [
                ["Всего товаров", len(results)],
                ["Общая прибыль", format_currency(total_profit)],
                ["Средняя маржа", format_percent(avg_margin)],
                ["Прибыльных товаров", profitable],
                ["Убыточных товаров", len(results) - profitable],
                ["", ""],
                ["Параметры:", ""],
                ["Маркетплейс", marketplace],
                ["Режим", mode],
                ["Единицы измерения", results[0].get('dimension_unit', 'мм') if results else 'мм'],
                ["Дата", datetime.now().strftime("%d.%m.%Y %H:%M")],
                ["Версия", APP_VERSION]
            ]
            
            for row_idx, (label, value) in enumerate(summary_data, 2):
                ws.cell(row=row_idx, column=1, value=label)
                ws.cell(row=row_idx, column=2, value=value)
        
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 20
    
    def load_existing(self, file_bytes: bytes) -> List[Dict]:
        if not LIBRARIES['openpyxl'] or not file_bytes:
            return []
        
        try:
            from openpyxl import load_workbook
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
            key_mapping = self._get_key_mapping()
            
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
                              "competitor_avg_price", "competitor_min_price", "competitor_max_price"]:
                        value = safe_float(value)
                    elif key in ["margin", "contribution_margin_pct", "roi", "ltv_cac_ratio"]:
                        value = safe_float(value)
                    elif key in ["payback_days", "break_even_units", "eoq", "competitor_count"]:
                        value = safe_int(value)
                    else:
                        value = safe_str(value)
                    
                    row_data[key] = value
                
                if row_data.get("article"):
                    if "mode" not in row_data:
                        row_data["mode"] = "FBY"
                    if "marketplace" not in row_data:
                        row_data["marketplace"] = "Не указан"
                    if "dimension_unit" not in row_data:
                        row_data["dimension_unit"] = "мм"
                    results.append(row_data)
            
            return results
            
        except Exception as e:
            logger.error(f"Error loading existing data: {e}")
            return []

# --------------------------------------------
# 🎨 ОСНОВНОЕ ПРИЛОЖЕНИЕ
# --------------------------------------------
class UnitEconomicsApp:
    def __init__(self):
        self.classifier = CategoryClassifier()
        self.exporter = ExcelExportEngine()
        self.tariff_provider = AITariffProvider()
        self.results = []
        self.all_rates = {}
        self.engine = None
        self.api_clients = {}
        self.uploader = MarketplaceUploader()
        self.oem_db = OEMDatabase()
        self.ai_editor = AIFileEditor()
        self.dimension_unit = "мм"
        self.existing_data = []
        
        if "theme" not in st.session_state:
            st.session_state.theme = "light"
        if "dimension_unit" not in st.session_state:
            st.session_state.dimension_unit = "мм"
        if "results" not in st.session_state:
            st.session_state.results = []
        if "existing_data" not in st.session_state:
            st.session_state.existing_data = []
        if "language" not in st.session_state:
            st.session_state.language = "ru"
        if "uploaded_data" not in st.session_state:
            st.session_state.uploaded_data = None
        if "ai_edited_data" not in st.session_state:
            st.session_state.ai_edited_data = None
    
    def run(self):
        st.set_page_config(
            page_title=CONFIG.app_name,
            page_icon="🚀",
            layout="wide",
            initial_sidebar_state="expanded"
        )
        
        self._render_header()
        self._render_sidebar()
        self._render_main()
    
    def _render_header(self):
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
                    padding: 1.5rem; border-radius: 15px; color: white; text-align: center; margin-bottom: 1.5rem;
                    border: 2px solid #e94560; box-shadow: 0 4px 15px rgba(0,0,0,0.3);">
            <h1 style="font-size: 2.8rem; margin: 0; text-shadow: 2px 2px 4px rgba(0,0,0,0.3);">
                🚀 {CONFIG.app_name}
            </h1>
            <p style="font-size: 1.2rem; opacity: 0.95; margin-top: 0.3rem;">
                📊 <strong>Товарная модель (B2C)</strong> | ИИ-редактирование | OE база | Конвертация размеров | Множественный парсинг
            </p>
            <div style="display: flex; justify-content: center; gap: 0.8rem; flex-wrap: wrap; margin-top: 0.5rem;">
                <span style="background: rgba(233,69,96,0.3); padding: 0.2rem 1.2rem; border-radius: 20px; font-size: 0.9rem;">
                    v{APP_VERSION}
                </span>
                <span style="background: rgba(233,69,96,0.3); padding: 0.2rem 1.2rem; border-radius: 20px; font-size: 0.9rem;">
                    📦 Товарная модель
                </span>
                <span style="background: rgba(233,69,96,0.3); padding: 0.2rem 1.2rem; border-radius: 20px; font-size: 0.9rem;">
                    🤖 ИИ-редактирование
                </span>
                <span style="background: rgba(233,69,96,0.3); padding: 0.2rem 1.2rem; border-radius: 20px; font-size: 0.9rem;">
                    📏 Конвертация мм/см
                </span>
                <span style="background: rgba(233,69,96,0.3); padding: 0.2rem 1.2rem; border-radius: 20px; font-size: 0.9rem;">
                    📋 Множественный парсинг
                </span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    def _render_sidebar(self):
        with st.sidebar:
            st.markdown("## ⚙️ Настройки")
            
            language = st.selectbox(
                "🌐 Язык / Language",
                ["Русский", "English"],
                index=0 if st.session_state.language == "ru" else 1,
                key="language_select"
            )
            st.session_state.language = "ru" if language == "Русский" else "en"
            
            st.divider()
            
            theme = st.toggle("🌙 Темная тема", value=st.session_state.theme == "dark", key="theme_toggle")
            st.session_state.theme = "dark" if theme else "light"
            
            st.divider()
            
            st.markdown("### 📏 Единицы измерения")
            
            dimension_unit = st.radio(
                "Выберите единицы для размеров",
                options=["мм", "см"],
                index=0 if st.session_state.dimension_unit == "мм" else 1,
                help="Выберите в каких единицах указаны размеры в вашем файле",
                key="dimension_unit_radio"
            )
            st.session_state.dimension_unit = dimension_unit
            self.dimension_unit = dimension_unit
            
            st.divider()
            
            st.markdown("### 🔑 API ключи")
            
            ym_api_key = st.text_input(
                "Яндекс Маркет API ключ",
                type="password",
                placeholder="Ваш API ключ",
                key="ym_api_key"
            )
            
            ozon_api_key = st.text_input(
                "Ozon API ключ",
                type="password",
                placeholder="Ваш API ключ",
                key="ozon_api_key"
            )
            
            ozon_client_id = st.text_input(
                "Ozon Client ID",
                type="password",
                placeholder="Ваш Client ID",
                key="ozon_client_id"
            )
            
            wb_api_key = st.text_input(
                "Wildberries API ключ",
                type="password",
                placeholder="Ваш API ключ",
                key="wb_api_key"
            )
            
            ds_api_key = st.text_input(
                "🔑 DeepSeek API ключ",
                type="password",
                placeholder="sk-...",
                help="Для AI-тарифов и ИИ-редактирования",
                key="ds_api_key"
            )
            if ds_api_key:
                self.tariff_provider.api_key = ds_api_key
                self.ai_editor.api_key = ds_api_key
                st.success("✅ DeepSeek ключ установлен")
            
            st.divider()
            
            st.markdown("### 📦 Параметры")
            
            marketplace = st.selectbox(
                "🏪 Маркетплейс",
                CONFIG.marketplaces,
                index=0,
                key="marketplace_select"
            )
            
            mode = st.selectbox(
                "📦 Режим работы",
                CONFIG.operation_modes,
                index=0,
                key="mode_select"
            )
            
            days_storage = st.number_input(
                "📦 Хранение (дней)",
                min_value=1,
                max_value=730,
                value=30,
                key="days_storage"
            )
            
            st.divider()
            
            col1, col2 = st.columns(2)
            with col1:
                if st.button("🔄 Обновить тарифы", use_container_width=True, key="refresh_rates"):
                    with st.spinner("⏳ Получение тарифов..."):
                        self.tariff_provider.clear_cache()
                        self.all_rates = self.tariff_provider.get_all_rates()
                        st.success("✅ Тарифы обновлены!")
            
            with col2:
                if st.button("🗑️ Очистить кэш", use_container_width=True, key="clear_cache"):
                    CacheManager().clear()
                    st.success("✅ Кэш очищен!")
            
            st.divider()
            
            if self.results:
                st.markdown("### 📊 Статистика")
                st.metric("📦 Позиций", len(self.results))
                profitable = sum(1 for r in self.results if r.get("unit_profit", 0) > 0)
                total_profit = sum(r.get("unit_profit", 0) for r in self.results)
                st.metric("💰 Прибыльных", f"{profitable}/{len(self.results)}")
                st.metric("💵 Общая прибыль", format_currency(total_profit))
            
            st.divider()
            st.markdown("### ℹ️ Система")
            st.caption(f"Версия: {APP_VERSION}")
            st.caption(f"Python: {sys.version[:10]}")
            st.caption(f"Библиотеки: {sum(1 for v in LIBRARIES.values() if v)}/{len(LIBRARIES)}")
    
    def _render_main(self):
        tabs = st.tabs([
            "📁 Загрузка данных", 
            "🤖 ИИ-редактирование",
            "📊 Расчет", 
            "🔍 OE база",
            "📋 Парсинг",
            "📤 Экспорт"
        ])
        
        with tabs[0]:
            self._render_upload_tab()
        
        with tabs[1]:
            self._render_ai_edit_tab()
        
        with tabs[2]:
            self._render_calculate_tab()
        
        with tabs[3]:
            self._render_oem_database_tab()
        
        with tabs[4]:
            self._render_parsing_tab()
        
        with tabs[5]:
            self._render_export_tab()
    
    def _render_upload_tab(self):
        st.markdown("## 📁 Загрузка данных")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 📤 Загрузите файл с товарами")
            
            uploaded_file = st.file_uploader(
                "Выберите Excel или CSV файл",
                type=['xlsx', 'xls', 'csv'],
                key="file_uploader"
            )
            
            if uploaded_file is not None:
                try:
                    if uploaded_file.name.endswith('.csv'):
                        df = pd.read_csv(uploaded_file, encoding='utf-8-sig')
                    else:
                        df = pd.read_excel(uploaded_file, engine='openpyxl')
                    
                    st.session_state.uploaded_data = df
                    st.success(f"✅ Загружено {len(df)} строк")
                    
                    st.markdown("### 📊 Предпросмотр данных")
                    st.dataframe(df.head(10), use_container_width=True)
                    
                except Exception as e:
                    st.error(f"❌ Ошибка загрузки файла: {str(e)}")
        
        with col2:
            st.markdown("### 📋 Инструкция по загрузке")
            st.info("""
            **Поддерживаемые форматы файлов:**
            - Excel (.xlsx, .xls)
            - CSV (.csv)
            
            **Необходимые колонки:**
            - `Артикул` - идентификатор товара
            - `Наименование` - название товара
            - `Цена` - цена продажи
            - `Длина`, `Ширина`, `Высота` - габариты в мм или см
            - `Вес` - вес в кг (опционально)
            
            **Дополнительные колонки:**
            - `OE номер` - оригинальный номер запчасти
            - `Бренд` - производитель
            - `Штрихкод` - штрихкод товара
            """)
            
            st.markdown("### 📥 Скачать шаблон")
            if st.button("📥 Скачать шаблон Excel", use_container_width=True):
                template_df = pd.DataFrame({
                    "Артикул": ["ABC-001", "ABC-002"],
                    "Наименование": ["Деталь A", "Деталь B"],
                    "Цена": [1000, 2500],
                    "Длина": [100, 150],
                    "Ширина": [80, 120],
                    "Высота": [50, 70],
                    "Вес": [0.5, 1.2],
                    "OE номер": ["123456", "789012"],
                    "Бренд": ["BOSCH", "DENSO"],
                    "Штрихкод": ["1234567890123", "4567890123456"]
                })
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    template_df.to_excel(writer, sheet_name='Товары', index=False)
                output.seek(0)
                st.download_button(
                    label="Скачать шаблон",
                    data=output.getvalue(),
                    file_name="шаблон_юнит_экономика.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
    def _render_ai_edit_tab(self):
        st.markdown("## 🤖 ИИ-редактирование данных")
        
        if st.session_state.uploaded_data is None:
            st.warning("⚠️ Сначала загрузите файл в разделе '📁 Загрузка данных'")
            return
        
        st.info("""
        🔮 **ИИ может:**
        - Исправлять ошибки в данных (цены, размеры, артикулы)
        - Заполнять пропуски
        - Пересчитывать формулы и показатели
        - Нормализовать формат данных
        - Добавлять новые колонки с расчетами
        - Исправлять несоответствия в категориях
        """)
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            edit_mode = st.radio(
                "Выберите режим редактирования",
                ["📝 Исправить данные", "🧮 Пересчитать формулы", "🔍 Найти и исправить ошибки"],
                horizontal=True,
                key="ai_edit_mode"
            )
            
            ai_prompt = st.text_area(
                "📝 Опишите, что нужно сделать с данными",
                placeholder="Пример: исправь цены, если они меньше 100 рублей, увеличь до 150; пересчитай маржинальность; добавь колонку 'Прибыль' = Цена - Себестоимость; исправь артикулы в формате ABC-123; заполни пропуски в колонке 'Бренд'...",
                height=120,
                key="ai_prompt"
            )
            
            col_btn1, col_btn2 = st.columns(2)
            
            with col_btn1:
                if st.button("🚀 Выполнить ИИ-редактирование", use_container_width=True, key="ai_edit_btn"):
                    if edit_mode == "📝 Исправить данные" and not ai_prompt.strip():
                        st.warning("⚠️ Опишите, что нужно сделать с данными")
                    else:
                        with st.spinner("🤖 ИИ обрабатывает данные..."):
                            df = st.session_state.uploaded_data
                            
                            if edit_mode == "📝 Исправить данные":
                                result_df, message = self.ai_editor.edit_data(df, ai_prompt)
                            elif edit_mode == "🧮 Пересчитать формулы":
                                result_df, message = self.ai_editor.edit_formulas(df, ai_prompt)
                            else:
                                result_df, message = self.ai_editor.analyze_and_fix_errors(df)
                            
                            st.session_state.ai_edited_data = result_df
                            st.session_state.uploaded_data = result_df
                            
                            if message.startswith("✅"):
                                st.success(message)
                            else:
                                st.warning(message)
            
            with col_btn2:
                if st.button("🔄 Сбросить изменения", use_container_width=True, key="ai_reset"):
                    st.session_state.ai_edited_data = None
                    st.session_state.uploaded_data = None
                    st.success("✅ Изменения сброшены")
        
        with col2:
            st.markdown("### 📊 Текущие данные")
            if st.session_state.uploaded_data is not None:
                df = st.session_state.uploaded_data
                st.metric("📦 Строк", len(df))
                st.metric("📋 Колонок", len(df.columns))
                
                st.markdown("#### 📋 Колонки:")
                for col in df.columns[:10]:
                    st.caption(f"• {col} ({df[col].dtype})")
                if len(df.columns) > 10:
                    st.caption(f"... и еще {len(df.columns) - 10} колонок")
        
        st.divider()
        
        if st.session_state.ai_edited_data is not None:
            st.markdown("### 📊 Результат ИИ-редактирования")
            
            df = st.session_state.ai_edited_data
            st.dataframe(df.head(20), use_container_width=True)
            
            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                csv = df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📥 Скачать CSV",
                    data=csv,
                    file_name="исправленные_данные.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            with col_dl2:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Данные', index=False)
                output.seek(0)
                st.download_button(
                    label="📥 Скачать Excel",
                    data=output.getvalue(),
                    file_name="исправленные_данные.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
    
    def _render_calculate_tab(self):
        st.markdown("## 📊 Расчет юнит-экономики")
        
        if st.session_state.uploaded_data is None:
            st.warning("⚠️ Сначала загрузите файл с данными")
            return
        
        col1, col2 = st.columns(2)
        
        with col1:
            marketplace = st.selectbox(
                "🏪 Маркетплейс",
                CONFIG.marketplaces,
                index=0,
                key="calc_marketplace"
            )
            
            mode = st.selectbox(
                "📦 Режим работы",
                CONFIG.operation_modes,
                index=0,
                key="calc_mode"
            )
            
            days_storage = st.number_input(
                "📦 Срок хранения (дней)",
                min_value=1,
                max_value=730,
                value=30,
                key="calc_days"
            )
            
            dimension_unit = st.radio(
                "📏 Единицы измерения размеров",
                ["мм", "см"],
                index=0 if st.session_state.dimension_unit == "мм" else 1,
                key="calc_dimension"
            )
            st.session_state.dimension_unit = dimension_unit
        
        with col2:
            st.markdown("### 📊 Параметры расчета")
            if st.session_state.uploaded_data is not None:
                df = st.session_state.uploaded_data
                st.caption(f"📦 Маркетплейс: {marketplace}")
                st.caption(f"📦 Режим: {mode}")
                st.caption(f"📅 Хранение: {days_storage} дней")
                st.caption(f"📏 Единицы: {dimension_unit}")
                st.caption(f"📋 Строк в файле: {len(df)}")
        
        if st.button("🚀 Рассчитать юнит-экономику", use_container_width=True, key="calc_btn"):
            with st.spinner("⏳ Выполняется расчет..."):
                df = st.session_state.uploaded_data.copy()
                rows = df.to_dict(orient='records')
                
                self.engine = UnitEconomicsEngine(
                    marketplace=marketplace,
                    mode=mode,
                    days_storage=days_storage,
                    dimension_unit=dimension_unit
                )
                
                self.results = self.engine.calculate_batch(rows, parallel=True)
                st.session_state.results = self.results
                
                st.success(f"✅ Рассчитано {len(self.results)} товаров")
                self._display_results()
        
        if st.session_state.results:
            self._display_results()
    
    def _display_results(self):
        if not self.results:
            self.results = st.session_state.get('results', [])
        
        if not self.results:
            st.warning("⚠️ Нет результатов для отображения")
            return
        
        df_results = pd.DataFrame(self.results)
        
        st.markdown("### 📊 Результаты расчета")
        st.dataframe(df_results, use_container_width=True, height=400)
        
        st.markdown("### 📈 Сводная статистика")
        col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
        
        total_profit = sum(r.get("unit_profit", 0) for r in self.results)
        avg_margin = sum(r.get("margin", 0) for r in self.results) / len(self.results) if self.results else 0
        profitable = sum(1 for r in self.results if r.get("unit_profit", 0) > 0)
        avg_roi = sum(r.get("roi", 0) for r in self.results) / len(self.results) if self.results else 0
        
        with col_stat1:
            st.metric("💰 Общая прибыль", format_currency(total_profit))
        with col_stat2:
            st.metric("📈 Средняя маржа", format_percent(avg_margin))
        with col_stat3:
            st.metric("✅ Прибыльных товаров", f"{profitable}/{len(self.results)}")
        with col_stat4:
            st.metric("📊 Средний ROI", format_percent(avg_roi))
        
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            csv = df_results.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 Скачать CSV",
                data=csv,
                file_name=f"юнит_экономика_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        with col_dl2:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_results.to_excel(writer, sheet_name='Юнит-экономика', index=False)
            output.seek(0)
            st.download_button(
                label="📥 Скачать Excel",
                data=output.getvalue(),
                file_name=f"юнит_экономика_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        if LIBRARIES['plotly']:
            try:
                st.markdown("### 📊 Визуализация")
                
                col_ch1, col_ch2 = st.columns(2)
                
                with col_ch1:
                    fig1 = px.histogram(
                        df_results, x="margin",
                        title="📈 Распределение маржинальности",
                        labels={"margin": "Маржа, %", "count": "Количество"},
                        nbins=20,
                        color_discrete_sequence=["#1a1a2e"]
                    )
                    st.plotly_chart(fig1, use_container_width=True)
                
                with col_ch2:
                    top_products = df_results.nlargest(10, "unit_profit")
                    if not top_products.empty:
                        fig2 = px.bar(
                            top_products,
                            x="name",
                            y="unit_profit",
                            title="💰 Топ-10 по прибыли",
                            labels={"name": "Товар", "unit_profit": "Прибыль, ₽"},
                            color="margin",
                            color_continuous_scale="Viridis"
                        )
                        fig2.update_layout(xaxis_tickangle=-45)
                        st.plotly_chart(fig2, use_container_width=True)
                
                category_data = df_results.groupby("category").agg({
                    "unit_profit": "sum",
                    "name": "count"
                }).reset_index().rename(columns={"name": "count"})
                
                if not category_data.empty:
                    fig3 = px.pie(
                        category_data,
                        values="unit_profit",
                        names="category",
                        title="💰 Прибыль по категориям"
                    )
                    st.plotly_chart(fig3, use_container_width=True)
                
                fig4 = px.scatter(
                    df_results,
                    x="price",
                    y="unit_profit",
                    color="margin",
                    size="volume",
                    title="💎 Прибыль vs Цена",
                    labels={"price": "Цена, ₽", "unit_profit": "Прибыль, ₽"},
                    color_continuous_scale="Viridis",
                    hover_data=["name", "category"]
                )
                st.plotly_chart(fig4, use_container_width=True)
                
            except Exception as e:
                logger.error(f"Chart error: {e}")
                st.warning("⚠️ Ошибка построения графиков")
    
    def _render_oem_database_tab(self):
        st.markdown("## 🔍 База OE номеров")
        
        col1, col2 = st.columns([1, 2])
        
        with col1:
            st.markdown("### 🔎 Поиск по OE номеру")
            oe_query = st.text_input("Введите OE номер", placeholder="Например: 0986AF0059", key="oe_search")
            
            if oe_query:
                result = self.oem_db.get_by_oe(oe_query)
                if result:
                    st.success(f"✅ Найдено: {result.get('brand', '')} - {result.get('category', '')}")
                    st.json(result)
                else:
                    st.warning("❌ OE номер не найден в базе")
            
            st.divider()
            
            st.markdown("### 📊 Статистика базы")
            stats = self.oem_db.get_statistics()
            st.metric("📦 Всего записей", stats["total"])
            st.metric("🏷️ Категорий", len(stats["categories"]))
            st.metric("🏢 Брендов", len(stats["brands"]))
            
            st.markdown("#### 🏷️ Топ категории:")
            for cat, count in sorted(stats["categories"].items(), key=lambda x: x[1], reverse=True)[:5]:
                st.caption(f"• {cat}: {count}")
        
        with col2:
            st.markdown("### 📋 Просмотр базы")
            
            search_by = st.selectbox(
                "Поиск по",
                ["Все", "Бренд", "Категория", "Применимость"],
                key="oe_search_by"
            )
            
            if search_by == "Бренд":
                brand_query = st.text_input("Введите бренд", placeholder="BOSCH", key="brand_search")
                if brand_query:
                    results = self.oem_db.search_by_brand(brand_query)
                    if results:
                        st.dataframe(pd.DataFrame(results), use_container_width=True)
                    else:
                        st.info("Ничего не найдено")
            elif search_by == "Категория":
                cat_query = st.text_input("Введите категорию", placeholder="Фильтры", key="cat_search")
                if cat_query:
                    results = self.oem_db.search_by_category(cat_query)
                    if results:
                        st.dataframe(pd.DataFrame(results), use_container_width=True)
                    else:
                        st.info("Ничего не найдено")
            elif search_by == "Применимость":
                compat_query = st.text_input("Введите модель", placeholder="BMW 3", key="compat_search")
                if compat_query:
                    results = self.oem_db.search_by_compatibility(compat_query)
                    if results:
                        st.dataframe(pd.DataFrame(results), use_container_width=True)
                    else:
                        st.info("Ничего не найдено")
            else:
                st.dataframe(
                    pd.DataFrame([
                        {"OE номер": oe, **data} 
                        for oe, data in list(self.oem_db.data.items())[:50]
                    ]),
                    use_container_width=True
                )
    
    def _render_parsing_tab(self):
        st.markdown("## 📋 Множественный парсинг цен")
        
        st.info("""
        📌 **Парсинг артикулов с маркетплейсов**
        
        Введите список артикулов или загрузите файл для массового парсинга цен конкурентов.
        Система проверит наличие товаров на выбранных маркетплейсах.
        """)
        
        tab1, tab2 = st.tabs(["📝 Список артикулов", "📁 Загрузка файла"])
        
        with tab1:
            articles_text = st.text_area(
                "Введите артикулы (по одному на строку)",
                placeholder="ABC-001\nABC-002\nABC-003",
                height=150,
                key="articles_text"
            )
            
            marketplace = st.selectbox(
                "Выберите маркетплейс для парсинга",
                ["Все", "Яндекс Маркет", "Ozon", "Wildberries"],
                key="parse_marketplace"
            )
            
            max_pages = st.slider(
                "Количество страниц для парсинга",
                min_value=1,
                max_value=5,
                value=1,
                key="parse_pages"
            )
            
            if st.button("🚀 Парсить артикулы", use_container_width=True, key="parse_btn"):
                articles = [a.strip() for a in articles_text.split('\n') if a.strip()]
                
                if not articles:
                    st.warning("⚠️ Введите хотя бы один артикул")
                else:
                    with st.spinner(f"⏳ Парсинг {len(articles)} артикулов..."):
                        parser = CompetitorParser()
                        
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        def update_progress(current, total, message):
                            progress_bar.progress(current / total)
                            status_text.text(message)
                        
                        parser.set_progress_callback(update_progress)
                        
                        results = parser.parse_multiple_articles(
                            articles,
                            marketplace,
                            max_pages
                        )
                        
                        st.session_state.parse_results = results
                        
                        df = parser.format_multiple_results(results)
                        st.dataframe(df, use_container_width=True)
                        
                        st.success(f"✅ Парсинг завершен! Найдено {len(df)} результатов")
                        
                        csv = df.to_csv(index=False, encoding='utf-8-sig')
                        st.download_button(
                            label="📥 Скачать результаты",
                            data=csv,
                            file_name="результаты_парсинга.csv",
                            mime="text/csv"
                        )
        
        with tab2:
            st.markdown("### 📤 Загрузите файл с артикулами")
            
            parse_file = st.file_uploader(
                "Выберите Excel или CSV файл",
                type=['xlsx', 'xls', 'csv'],
                key="parse_file_uploader"
            )
            
            if parse_file is not None:
                try:
                    if parse_file.name.endswith('.csv'):
                        df_parse = pd.read_csv(parse_file, encoding='utf-8-sig')
                    else:
                        df_parse = pd.read_excel(parse_file, engine='openpyxl')
                    
                    st.dataframe(df_parse.head(10), use_container_width=True)
                    
                    article_col = st.selectbox(
                        "Выберите колонку с артикулами",
                        df_parse.columns,
                        key="parse_article_col"
                    )
                    
                    brand_col = st.selectbox(
                        "Выберите колонку с брендами (опционально)",
                        [None] + list(df_parse.columns),
                        key="parse_brand_col"
                    )
                    
                    if st.button("🚀 Парсить из файла", use_container_width=True, key="parse_file_btn"):
                        with st.spinner("⏳ Парсинг артикулов из файла..."):
                            articles = df_parse[article_col].astype(str).tolist()
                            articles = [a.strip() for a in articles if a.strip()]
                            
                            parser = CompetitorParser()
                            results = parser.parse_multiple_articles(articles, marketplace, max_pages)
                            
                            if brand_col:
                                brand_map = dict(zip(df_parse[article_col].astype(str), df_parse[brand_col].astype(str)))
                                for article, brand in brand_map.items():
                                    if article in results:
                                        results[article]['_brand'] = brand
                            
                            df_results = parser.format_multiple_results(results)
                            st.dataframe(df_results, use_container_width=True)
                            st.success(f"✅ Парсинг завершен! Найдено {len(df_results)} результатов")
                            
                            csv = df_results.to_csv(index=False, encoding='utf-8-sig')
                            st.download_button(
                                label="📥 Скачать результаты",
                                data=csv,
                                file_name="результаты_парсинга_файл.csv",
                                mime="text/csv"
                            )
                            
                except Exception as e:
                    st.error(f"❌ Ошибка загрузки файла: {str(e)}")
    
    def _render_export_tab(self):
        st.markdown("## 📤 Экспорт результатов")
        
        if not self.results:
            st.warning("⚠️ Сначала выполните расчет в разделе '📊 Расчет'")
            return
        
        st.success(f"✅ Готово к экспорту: {len(self.results)} товаров")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 📊 Excel экспорт")
            
            include_summary = st.checkbox("Включить сводку", value=True, key="export_summary")
            include_rates = st.checkbox("Включить тарифы", value=True, key="export_rates")
            
            if st.button("📥 Экспорт в Excel", use_container_width=True, key="export_excel_btn"):
                with st.spinner("⏳ Формирование Excel файла..."):
                    try:
                        marketplace = self.engine.marketplace if self.engine else "Яндекс Маркет"
                        mode = self.engine.mode if self.engine else "FBY"
                        
                        all_rates = self.tariff_provider.get_all_rates() if include_rates else None
                        
                        excel_data = self.exporter.export(
                            self.results,
                            marketplace,
                            mode,
                            all_rates
                        )
                        
                        st.download_button(
                            label="📥 Скачать Excel файл",
                            data=excel_data,
                            file_name=f"юнит_экономика_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        
                        st.success("✅ Excel файл готов к скачиванию!")
                        
                    except Exception as e:
                        st.error(f"❌ Ошибка экспорта: {str(e)}")
        
        with col2:
            st.markdown("### 📊 Другие форматы")
            
            if st.button("📥 CSV (для Excel/Google Sheets)", use_container_width=True, key="export_csv_btn"):
                df = pd.DataFrame(self.results)
                csv = df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📥 Скачать CSV",
                    data=csv,
                    file_name=f"юнит_экономика_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            
            if st.button("📥 JSON (для API)", use_container_width=True, key="export_json_btn"):
                json_data = json.dumps(self.results, ensure_ascii=False, indent=2, default=str)
                st.download_button(
                    label="📥 Скачать JSON",
                    data=json_data,
                    file_name=f"юнит_экономика_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                    mime="application/json",
                    use_container_width=True
                )
            
            st.divider()
            
            st.markdown("### 📤 Отправка в Telegram")
            if "telegram" in st.session_state and st.session_state.telegram:
                if st.button("📤 Отправить отчет в Telegram", use_container_width=True):
                    with st.spinner("⏳ Отправка..."):
                        try:
                            if st.session_state.telegram.send_report(self.results):
                                st.success("✅ Отчет отправлен в Telegram!")
                            else:
                                st.error("❌ Ошибка отправки в Telegram")
                        except Exception as e:
                            st.error(f"❌ Ошибка: {str(e)}")
            else:
                st.caption("🔑 Для отправки в Telegram настройте токен в боковой панели")

# --------------------------------------------
# ЗАПУСК
# --------------------------------------------
if __name__ == "__main__":
    try:
        app = UnitEconomicsApp()
        app.run()
    except Exception as e:
        st.error(f"❌ Критическая ошибка: {str(e)}")
        st.code(traceback.format_exc())
        logger.error(f"Critical error: {e}")
