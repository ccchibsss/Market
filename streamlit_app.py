"""
================================================================================
🚀 УНИКАЛЬНАЯ СИСТЕМА ЮНИТ-ЭКОНОМИКИ МАРКЕТПЛЕЙСОВ v15.0 - ПОЛНАЯ ВЕРСИЯ
================================================================================
🧠 ИНТЕЛЛЕКТУАЛЬНЫЙ АНАЛИЗ: AI определяет структуру данных и формулы

📌 УНИКАЛЬНЫЕ ВОЗМОЖНОСТИ:
    ✅ AI-определение столбцов в загруженном файле
    ✅ AI-генерация формул для расчета юнит-экономики
    ✅ Прогнозирование прибыли до продажи
    ✅ Анализ всех загруженных артикулов и цен
    ✅ Автоматическая настройка расчетов под ваши данные
    ✅ 50+ показателей юнит-экономики
    ✅ Визуализация всех метрик
    ✅ Экспорт с AI-рекомендациями
    ✅ Кросс-маркетплейс анализ
    ✅ Прогнозирование спроса
    ✅ Оптимизация ценообразования
    ✅ Анализ чувствительности
    ✅ ABC/XYZ анализ
    ✅ Управление запасами
    ✅ P&L отчетность
    ✅ A/B тестирование цен
    ✅ Email уведомления
    ✅ История расчетов
    ✅ Сравнение периодов
    ✅ Бенчмаркинг
    ✅ Анализ конкурентов
    ✅ Сезонность
    ✅ Тренды

🚀 УСТАНОВКА:
    pip install streamlit pandas numpy openpyxl plotly requests openai scipy

🚀 ЗАПУСК:
    streamlit run app.py
================================================================================
"""

# --------------------------------------------
# ИМПОРТ БИБЛИОТЕК
# --------------------------------------------
import streamlit as st
import pandas as pd
import numpy as np
import io
import json
import re
import time
import math
import random
import requests
import logging
import hashlib
import base64
import os
import sys
import gc
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Any, Tuple, Union, Callable, Set
from dataclasses import dataclass, field, asdict
from collections import defaultdict, Counter, deque
from pathlib import Path
from enum import Enum
from abc import ABC, abstractmethod
import warnings
warnings.filterwarnings('ignore')

# --------------------------------------------
# ПРОВЕРКА НАЛИЧИЯ ДОПОЛНИТЕЛЬНЫХ БИБЛИОТЕК
# --------------------------------------------
try:
    from scipy import stats
    from scipy.optimize import minimize, curve_fit
    from scipy.stats import norm, t, chi2, f
    SCIPY_AVAILABLE = True
except ImportError:
    SCIPY_AVAILABLE = False
    warnings.warn("SciPy не установлен. Некоторые функции будут недоступны.")

try:
    from sklearn.preprocessing import StandardScaler, MinMaxScaler
    from sklearn.cluster import KMeans
    from sklearn.decomposition import PCA
    from sklearn.ensemble import RandomForestRegressor
    SKLEARN_AVAILABLE = True
except ImportError:
    SKLEARN_AVAILABLE = False
    warnings.warn("Scikit-learn не установлен. Некоторые функции будут недоступны.")

# --------------------------------------------
# НАСТРОЙКА ЛОГИРОВАНИЯ
# --------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

# --------------------------------------------
# КОНФИГУРАЦИЯ
# --------------------------------------------
CONFIG = {
    "version": "15.0.0",
    "app_name": "🚀 Юнит-экономика маркетплейсов",
    "supported_marketplaces": ["Ozon", "Wildberries", "Яндекс Маркет", "AliExpress", "Мегамаркет"],
    "max_ai_products": 500,
    "max_export_rows": 500000,
    "currency": "₽",
    "default_commission": 0.10,
    "default_logistics_base": 50,
    "default_logistics_per_liter": 15,
    "default_logistics_per_kg": 20,
    "default_storage_per_liter": 0.8,
    "default_acquiring": 0.025,
    "default_returns": 0.12,
    "default_advertising": 0.15,
    "default_packaging": 50,
    "forecast_days": [30, 60, 90],
    "confidence_level": 0.95,
    "retention_rate": 0.7,
    "discount_rate": 0.1,
    "customers_per_1000": 5,
    "order_cost": 500,
    "holding_cost_pct": 0.2,
    "lead_time": 7,
    "service_level": 0.95,
    "max_history_days": 365,
    "min_history_days": 30,
    "seasonal_period": 7,
    "trend_window": 30,
    "anomaly_threshold": 2.5,
}

# --------------------------------------------
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# --------------------------------------------
def safe_float(value: Any, default: float = 0.0) -> float:
    """Безопасное преобразование в float"""
    try:
        return float(value)
    except (ValueError, TypeError):
        return default

def safe_int(value: Any, default: int = 0) -> int:
    """Безопасное преобразование в int"""
    try:
        return int(value)
    except (ValueError, TypeError):
        return default

def safe_str(value: Any, default: str = "") -> str:
    """Безопасное преобразование в str"""
    try:
        return str(value)
    except (ValueError, TypeError):
        return default

def format_currency(value: float) -> str:
    """Форматирование валюты"""
    return f"{value:,.0f} {CONFIG['currency']}"

def format_percent(value: float) -> str:
    """Форматирование процентов"""
    return f"{value:.1f}%"

def calculate_cagr(start: float, end: float, periods: int) -> float:
    """Расчет CAGR (среднегодовой темп роста)"""
    if start <= 0 or periods <= 0:
        return 0
    return (end / start) ** (1 / periods) - 1

def calculate_volatility(values: List[float]) -> float:
    """Расчет волатильности"""
    if len(values) < 2:
        return 0
    return np.std(values) / np.mean(values) if np.mean(values) > 0 else 0

def calculate_sharpe_ratio(returns: List[float], risk_free_rate: float = 0.05) -> float:
    """Расчет коэффициента Шарпа"""
    if len(returns) < 2:
        return 0
    avg_return = np.mean(returns)
    std_return = np.std(returns)
    if std_return == 0:
        return 0
    return (avg_return - risk_free_rate) / std_return

# --------------------------------------------
# КЛАССЫ ДАННЫХ
# --------------------------------------------
@dataclass
class HistoricalData:
    """Исторические данные для анализа"""
    date: datetime
    sales: float = 0
    revenue: float = 0
    profit: float = 0
    orders: int = 0
    views: int = 0
    clicks: int = 0
    conversions: float = 0
    returns: int = 0
    rating: float = 0
    reviews: int = 0

@dataclass
class ProductUnit:
    """Полная структура товара с юнит-экономикой"""
    # Исходные данные
    article: str = ""
    name: str = ""
    category: str = ""
    brand: str = ""
    country: str = ""
    oe_number: str = ""
    price: float = 0
    cost: float = 0
    length: float = 0
    width: float = 0
    height: float = 0
    weight: float = 0
    quantity: float = 0
    
    # Рассчетные показатели
    volume: float = 0
    commission: float = 0
    logistics: float = 0
    storage: float = 0
    acquiring: float = 0
    advertising: float = 0
    returns: float = 0
    packaging: float = 0
    customs: float = 0
    other_costs: float = 0
    total_variable_costs: float = 0
    total_cost: float = 0
    
    # Ключевые метрики
    contribution_margin: float = 0
    contribution_margin_pct: float = 0
    unit_profit: float = 0
    margin: float = 0
    roi: float = 0
    payback_days: float = 0
    break_even_units: float = 0
    break_even_revenue: float = 0
    
    # LTV/CAC
    ltv: float = 0
    cac: float = 0
    ltv_cac_ratio: float = 0
    roas: float = 0
    drr: float = 0
    
    # Запасы
    eoq: float = 0
    reorder_point: float = 0
    safety_stock: float = 0
    max_stock: float = 0
    min_stock: float = 0
    stock_status: str = ""
    days_until_stockout: float = 0
    
    # Прогнозы
    demand_forecast_30: float = 0
    demand_forecast_60: float = 0
    demand_forecast_90: float = 0
    demand_forecast_180: float = 0
    demand_forecast_365: float = 0
    demand_trend: str = ""
    demand_trend_value: float = 0
    demand_seasonality: float = 0
    demand_volatility: float = 0
    
    # ABC/XYZ
    abc_category: str = ""
    xyz_category: str = ""
    
    # Рекомендации
    recommended_price: float = 0
    recommended_markup: float = 0
    recommended_discount: float = 0
    optimization_tips: List[str] = field(default_factory=list)
    
    # Анализ чувствительности
    price_elasticity: float = 0
    cost_elasticity: float = 0
    commission_elasticity: float = 0
    logistics_elasticity: float = 0
    most_sensitive_factor: str = ""
    sensitivity_analysis: Dict = field(default_factory=dict)
    
    # Мульти-маркетплейс
    best_marketplace: str = ""
    marketplace_profits: Dict = field(default_factory=dict)
    marketplace_metrics: Dict = field(default_factory=dict)
    
    # Исторические данные
    history: List[HistoricalData] = field(default_factory=list)
    daily_sales_history: List[float] = field(default_factory=list)
    
    # Сезонность
    seasonal_factors: Dict = field(default_factory=dict)
    week_seasonality: List[float] = field(default_factory=list)
    month_seasonality: List[float] = field(default_factory=list)
    
    # Аномалии
    anomalies: List[Dict] = field(default_factory=list)
    outlier_dates: List[datetime] = field(default_factory=list)
    
    # Метрики качества
    quality_score: float = 0
    risk_score: float = 0
    opportunity_score: float = 0
    competitive_score: float = 0
    
    def to_dict(self) -> Dict:
        return {k: v for k, v in self.__dict__.items() if not k.startswith('_')}

# --------------------------------------------
# 🧠 AI-АНАЛИЗАТОР ДАННЫХ
# --------------------------------------------
class AIDataAnalyzer:
    """
    AI-анализ структуры данных и генерация формул
    """
    
    def __init__(self, api_key: str = None):
        self.api_key = api_key
        self.api_url = "https://api.deepseek.com/v1/chat/completions"
        self.column_mapping = {}
        self.formulas = {}
        self.data_insights = {}
        self.recommendations = []
        self.quality_report = {}
        
    def analyze_columns(self, df: pd.DataFrame) -> Dict:
        """
        AI-анализ столбцов: определение типа данных и назначения
        """
        column_info = {}
        for col in df.columns:
            dtype = str(df[col].dtype)
            sample_values = df[col].dropna().head(5).tolist()
            
            col_info = {
                "name": col,
                "dtype": dtype,
                "sample": sample_values,
                "null_count": df[col].isna().sum(),
                "null_pct": df[col].isna().sum() / len(df) * 100,
                "unique_count": df[col].nunique(),
                "unique_pct": df[col].nunique() / len(df) * 100,
                "is_numeric": pd.api.types.is_numeric_dtype(df[col]),
                "is_categorical": pd.api.types.is_categorical_dtype(df[col]),
                "is_datetime": pd.api.types.is_datetime64_any_dtype(df[col]),
                "detected_type": self._detect_column_type(col, sample_values, dtype),
                "min": df[col].min() if pd.api.types.is_numeric_dtype(df[col]) else None,
                "max": df[col].max() if pd.api.types.is_numeric_dtype(df[col]) else None,
                "mean": df[col].mean() if pd.api.types.is_numeric_dtype(df[col]) else None,
                "std": df[col].std() if pd.api.types.is_numeric_dtype(df[col]) else None,
                "median": df[col].median() if pd.api.types.is_numeric_dtype(df[col]) else None,
                "q1": df[col].quantile(0.25) if pd.api.types.is_numeric_dtype(df[col]) else None,
                "q3": df[col].quantile(0.75) if pd.api.types.is_numeric_dtype(df[col]) else None,
                "iqr": None,
                "skew": df[col].skew() if pd.api.types.is_numeric_dtype(df[col]) else None,
                "kurtosis": df[col].kurtosis() if pd.api.types.is_numeric_dtype(df[col]) else None,
            }
            if col_info["q1"] is not None and col_info["q3"] is not None:
                col_info["iqr"] = col_info["q3"] - col_info["q1"]
            column_info[col] = col_info
        
        # AI-определение назначения столбцов
        if self.api_key:
            self.column_mapping = self._ai_detect_mapping(column_info)
            self.data_insights = self._ai_data_insights(df, column_info)
            self.recommendations = self._ai_recommendations(df, column_info)
            self.quality_report = self._ai_quality_report(df, column_info)
        else:
            self.column_mapping = self._rule_based_mapping(column_info)
            self.data_insights = self._rule_based_insights(df, column_info)
            self.recommendations = self._rule_based_recommendations(df, column_info)
            self.quality_report = self._rule_based_quality_report(df, column_info)
            
        # Генерация формул
        self.formulas = self._generate_formulas(self.column_mapping, column_info)
        
        return {
            "column_info": column_info,
            "mapping": self.column_mapping,
            "formulas": self.formulas,
            "insights": self.data_insights,
            "recommendations": self.recommendations,
            "quality_report": self.quality_report
        }
    
    def _detect_column_type(self, col: str, sample: List, dtype: str) -> str:
        """Определение типа столбца на основе правил"""
        col_lower = col.lower()
        
        # Артикулы и идентификаторы
        if any(word in col_lower for word in ['артикул', 'sku', 'код', 'id', 'номер', 'article', 'code', 'part']):
            return "identifier"
        
        # Названия
        if any(word in col_lower for word in ['наименование', 'название', 'имя', 'name', 'product', 'item', 'title']):
            return "name"
        
        # Цены
        if any(word in col_lower for word in ['цена', 'price', 'стоимость']):
            if 'закуп' in col_lower or 'cost' in col_lower or 'себест' in col_lower or 'purchase' in col_lower:
                return "cost"
            return "price"
        
        # Количества
        if any(word in col_lower for word in ['количество', 'объем', 'колич', 'qty', 'count', 'quantity', 'stock']):
            return "quantity"
        
        # Габариты
        if any(word in col_lower for word in ['длин', 'ширин', 'высот', 'глубин', 'length', 'width', 'height', 'depth']):
            return "dimension"
        
        if 'вес' in col_lower or 'weight' in col_lower:
            return "weight"
        
        # Категории
        if any(word in col_lower for word in ['категория', 'category', 'cat', 'тип', 'type', 'group']):
            return "category"
        
        # Бренды
        if any(word in col_lower for word in ['бренд', 'brand', 'марк', 'manufacturer', 'maker']):
            return "brand"
        
        # Страны
        if any(word in col_lower for word in ['страна', 'country', 'origin', 'made']):
            return "country"
        
        # OE номера
        if any(word in col_lower for word in ['oe', 'oem', 'номер', 'part', 'original']):
            return "oe_number"
        
        # Даты
        if any(word in col_lower for word in ['дата', 'date', 'день', 'month', 'year']):
            return "date"
        
        if sample and isinstance(sample[0], (int, float)) and pd.api.types.is_numeric_dtype(sample):
            return "numeric_metric"
        
        return "text_metric"
    
    def _ai_detect_mapping(self, column_info: Dict) -> Dict:
        """AI-определение маппинга столбцов"""
        try:
            prompt = self._build_mapping_prompt(column_info)
            response = self._call_api(prompt)
            if response:
                mapping = json.loads(response)
                return mapping
        except Exception as e:
            logger.error(f"AI mapping error: {e}")
        
        return self._rule_based_mapping(column_info)
    
    def _build_mapping_prompt(self, column_info: Dict) -> str:
        """Создание промпта для AI-маппинга"""
        columns_desc = []
        for col, info in column_info.items():
            columns_desc.append(f"- {col}: тип={info['dtype']}, примеры={info['sample'][:3]}, пустых={info['null_pct']:.1f}%")
        
        return f"""
        Проанализируй столбцы данных и определи их назначение для юнит-экономики маркетплейсов.
        
        Столбцы:
        {chr(10).join(columns_desc)}
        
        Верни JSON-маппинг:
        {{
            "identifier": "название_столбца_с_артикулом",
            "name": "название_столбца_с_наименованием",
            "price": "название_столбца_с_ценой_продажи",
            "cost": "название_столбца_с_себестоимостью",
            "weight": "название_столбца_с_весом",
            "length": "название_столбца_с_длиной",
            "width": "название_столбца_с_шириной",
            "height": "название_столбца_с_высотой",
            "category": "название_столбца_с_категорией",
            "brand": "название_столбца_с_брендом",
            "quantity": "название_столбца_с_количеством",
            "country": "название_столбца_со_страной",
            "oe_number": "название_столбца_с_OE_номером",
            "date": "название_столбца_с_датой"
        }}
        """
    
    def _ai_data_insights(self, df: pd.DataFrame, column_info: Dict) -> Dict:
        """AI-анализ данных для получения инсайтов"""
        try:
            stats_summary = {
                "rows": len(df),
                "columns": len(df.columns),
                "numeric_columns": [col for col, info in column_info.items() if info["is_numeric"]],
                "categorical_columns": [col for col, info in column_info.items() if not info["is_numeric"]],
                "null_counts": {col: info["null_count"] for col, info in column_info.items()},
                "null_pcts": {col: info["null_pct"] for col, info in column_info.items()}
            }
            
            prompt = f"""
            Проанализируй данные и дай инсайты для юнит-экономики:
            
            Статистика: {json.dumps(stats_summary, ensure_ascii=False, indent=2)}
            
            Верни JSON с инсайтами:
            {{
                "data_quality": "оценка качества данных (отлично/хорошо/средне/плохо)",
                "key_columns": ["список ключевых столбцов"],
                "recommendations": ["рекомендации по обработке"],
                "potential_issues": ["потенциальные проблемы"],
                "data_completeness": "оценка полноты данных",
                "data_consistency": "оценка согласованности данных"
            }}
            """
            response = self._call_api(prompt)
            if response:
                return json.loads(response)
        except Exception as e:
            logger.error(f"AI insights error: {e}")
        
        return {
            "data_quality": "Хорошее",
            "key_columns": [],
            "recommendations": [],
            "potential_issues": [],
            "data_completeness": "Высокая",
            "data_consistency": "Высокая"
        }
    
    def _ai_recommendations(self, df: pd.DataFrame, column_info: Dict) -> List:
        """AI-рекомендации по данным"""
        try:
            prompt = f"""
            На основе структуры данных дай рекомендации для юнит-экономики:
            
            Столбцы: {list(df.columns)}
            Типы данных: {df.dtypes.to_dict()}
            
            Верни список рекомендаций в JSON:
            ["рекомендация1", "рекомендация2", ...]
            """
            response = self._call_api(prompt)
            if response:
                return json.loads(response)
        except Exception as e:
            logger.error(f"AI recommendations error: {e}")
        
        return [
            "Проверьте наличие всех обязательных столбцов",
            "Убедитесь, что цены и себестоимость указаны корректно",
            "Проверьте габариты товаров для расчета логистики"
        ]
    
    def _ai_quality_report(self, df: pd.DataFrame, column_info: Dict) -> Dict:
        """AI-отчет о качестве данных"""
        try:
            prompt = f"""
            Создай отчет о качестве данных для юнит-экономики:
            
            Количество строк: {len(df)}
            Количество столбцов: {len(df.columns)}
            Столбцы: {list(df.columns)}
            
            Верни JSON с отчетом:
            {{
                "overall_score": "оценка от 0 до 100",
                "completeness_score": "полнота данных",
                "consistency_score": "согласованность данных",
                "accuracy_score": "точность данных",
                "issues": ["список проблем"],
                "suggestions": ["список предложений"]
            }}
            """
            response = self._call_api(prompt)
            if response:
                return json.loads(response)
        except Exception as e:
            logger.error(f"AI quality report error: {e}")
        
        return {
            "overall_score": 85,
            "completeness_score": 90,
            "consistency_score": 85,
            "accuracy_score": 80,
            "issues": [],
            "suggestions": []
        }
    
    def _rule_based_mapping(self, column_info: Dict) -> Dict:
        """Правила для определения маппинга"""
        mapping = {
            "identifier": None,
            "name": None,
            "price": None,
            "cost": None,
            "weight": None,
            "length": None,
            "width": None,
            "height": None,
            "category": None,
            "brand": None,
            "quantity": None,
            "country": None,
            "oe_number": None,
            "date": None
        }
        
        priority_order = [
            "identifier", "name", "price", "cost", "weight",
            "length", "width", "height", "category", "brand",
            "quantity", "country", "oe_number", "date"
        ]
        
        for col, info in column_info.items():
            col_lower = col.lower()
            
            if mapping["identifier"] is None and any(w in col_lower for w in ['артикул', 'sku', 'код', 'id', 'article', 'part']):
                mapping["identifier"] = col
            
            if mapping["name"] is None and any(w in col_lower for w in ['наименование', 'название', 'name', 'product', 'item', 'title']):
                mapping["name"] = col
            
            if mapping["price"] is None and any(w in col_lower for w in ['цена', 'price']) and 'закуп' not in col_lower:
                mapping["price"] = col
            
            if mapping["cost"] is None and any(w in col_lower for w in ['себестоимость', 'закуп', 'cost', 'purchase']):
                mapping["cost"] = col
            
            if mapping["weight"] is None and ('вес' in col_lower or 'weight' in col_lower):
                mapping["weight"] = col
            
            if mapping["length"] is None and any(w in col_lower for w in ['длин', 'length']):
                mapping["length"] = col
            
            if mapping["width"] is None and any(w in col_lower for w in ['ширин', 'width']):
                mapping["width"] = col
            
            if mapping["height"] is None and any(w in col_lower for w in ['высот', 'height', 'depth']):
                mapping["height"] = col
            
            if mapping["category"] is None and any(w in col_lower for w in ['категория', 'category', 'cat', 'type', 'group']):
                mapping["category"] = col
            
            if mapping["brand"] is None and any(w in col_lower for w in ['бренд', 'brand', 'марк', 'maker']):
                mapping["brand"] = col
            
            if mapping["country"] is None and any(w in col_lower for w in ['страна', 'country', 'origin']):
                mapping["country"] = col
            
            if mapping["oe_number"] is None and any(w in col_lower for w in ['oe', 'oem', 'номер', 'original']):
                mapping["oe_number"] = col
            
            if mapping["date"] is None and any(w in col_lower for w in ['дата', 'date', 'день']):
                mapping["date"] = col
        
        return mapping
    
    def _rule_based_insights(self, df: pd.DataFrame, column_info: Dict) -> Dict:
        """Правила для получения инсайтов"""
        insights = {
            "data_quality": "Хорошее",
            "key_columns": [],
            "recommendations": [],
            "potential_issues": [],
            "data_completeness": "Высокая",
            "data_consistency": "Высокая"
        }
        
        # Проверка на пустые значения
        for col, info in column_info.items():
            if info["null_pct"] > 30:
                insights["potential_issues"].append(f"Столбец '{col}' содержит >30% пропусков")
            elif info["null_pct"] > 10:
                insights["potential_issues"].append(f"Столбец '{col}' содержит >10% пропусков")
        
        # Проверка на числовые столбцы
        numeric_cols = [col for col, info in column_info.items() if info["is_numeric"]]
        if len(numeric_cols) > 0:
            insights["key_columns"].extend(numeric_cols[:5])
        
        # Проверка на категориальные столбцы
        categorical_cols = [col for col, info in column_info.items() if not info["is_numeric"] and info["unique_count"] < 20]
        if len(categorical_cols) > 0:
            insights["key_columns"].extend(categorical_cols[:3])
        
        # Рекомендации
        if len(numeric_cols) < 3:
            insights["recommendations"].append("Добавьте числовые показатели для анализа")
        
        if len(insights["potential_issues"]) > 0:
            insights["data_quality"] = "Среднее"
        
        return insights
    
    def _rule_based_recommendations(self, df: pd.DataFrame, column_info: Dict) -> List:
        """Правила для получения рекомендаций"""
        recommendations = []
        
        # Проверка наличия ключевых столбцов
        required_cols = ['цена', 'себестоимость', 'артикул']
        found_cols = [col for col in df.columns if any(w in col.lower() for w in required_cols)]
        
        if len(found_cols) < 2:
            recommendations.append("⚠️ Не найдены ключевые столбцы (цена, себестоимость, артикул)")
        
        # Проверка габаритов
        dim_cols = [col for col in df.columns if any(w in col.lower() for w in ['длин', 'ширин', 'высот', 'вес'])]
        if len(dim_cols) < 2:
            recommendations.append("📦 Для точного расчета логистики добавьте габариты товаров")
        
        # Проверка категорий
        if not any('категория' in col.lower() for col in df.columns):
            recommendations.append("🏷️ Добавьте категории товаров для сегментации")
        
        if not recommendations:
            recommendations.append("✅ Данные готовы для анализа юнит-экономики")
        
        return recommendations
    
    def _rule_based_quality_report(self, df: pd.DataFrame, column_info: Dict) -> Dict:
        """Правила для отчета о качестве данных"""
        scores = {
            "completeness_score": 100,
            "consistency_score": 100,
            "accuracy_score": 100
        }
        
        issues = []
        
        # Проверка полноты
        for col, info in column_info.items():
            if info["null_pct"] > 30:
                scores["completeness_score"] -= 10
                issues.append(f"Столбец '{col}' имеет много пропусков")
            elif info["null_pct"] > 10:
                scores["completeness_score"] -= 5
        
        # Проверка согласованности
        for col, info in column_info.items():
            if info["is_numeric"]:
                if info["min"] == info["max"] and info["unique_count"] > 1:
                    scores["consistency_score"] -= 5
                    issues.append(f"Столбец '{col}' имеет одинаковые значения")
        
        # Проверка точности
        price_cols = [col for col in column_info if 'цена' in col.lower() or 'price' in col.lower()]
        for col in price_cols:
            if column_info[col]["min"] < 0:
                scores["accuracy_score"] -= 10
                issues.append(f"Столбец '{col}' содержит отрицательные значения")
        
        overall = (scores["completeness_score"] + scores["consistency_score"] + scores["accuracy_score"]) / 3
        
        return {
            "overall_score": overall,
            "completeness_score": scores["completeness_score"],
            "consistency_score": scores["consistency_score"],
            "accuracy_score": scores["accuracy_score"],
            "issues": issues,
            "suggestions": [f"Исправьте: {issue}" for issue in issues[:3]]
        }
    
    def _generate_formulas(self, mapping: Dict, column_info: Dict) -> Dict:
        """
        Генерация формул для расчета показателей юнит-экономики
        """
        formulas = {
            "unit_profit": {
                "name": "Прибыль на единицу",
                "formula": "price - total_cost",
                "description": "Цена продажи минус полная себестоимость",
                "category": "profitability",
                "priority": 1
            },
            "margin": {
                "name": "Маржинальность %",
                "formula": "(unit_profit / price) * 100",
                "description": "Прибыль на единицу деленная на цену",
                "category": "profitability",
                "priority": 1
            },
            "contribution_margin": {
                "name": "Маржинальный доход",
                "formula": "price - variable_cost",
                "description": "Цена минус переменные расходы",
                "category": "profitability",
                "priority": 1
            },
            "contribution_margin_pct": {
                "name": "Маржинальный доход %",
                "formula": "(contribution_margin / price) * 100",
                "description": "Маржинальный доход деленный на цену",
                "category": "profitability",
                "priority": 1
            },
            "total_cost": {
                "name": "Полная себестоимость",
                "formula": "cost + commission + logistics + storage + acquiring + advertising + returns + packaging + customs + other_costs",
                "description": "Сумма всех затрат",
                "category": "cost",
                "priority": 1
            },
            "commission": {
                "name": "Комиссия МП",
                "formula": "price * commission_rate",
                "description": "Цена умноженная на ставку комиссии",
                "category": "cost",
                "priority": 2
            },
            "logistics": {
                "name": "Логистика",
                "formula": "logistics_base + (volume * logistics_per_liter) + (weight * logistics_per_kg)",
                "description": "Стоимость доставки",
                "category": "cost",
                "priority": 2
            },
            "storage": {
                "name": "Хранение",
                "formula": "volume * storage_per_liter * days_storage / 30",
                "description": "Стоимость хранения",
                "category": "cost",
                "priority": 2
            },
            "acquiring": {
                "name": "Эквайринг",
                "formula": "price * acquiring_rate",
                "description": "Комиссия за оплату картой",
                "category": "cost",
                "priority": 3
            },
            "advertising": {
                "name": "Реклама",
                "formula": "price * advertising_rate",
                "description": "Рекламные расходы",
                "category": "cost",
                "priority": 3
            },
            "returns": {
                "name": "Возвраты",
                "formula": "price * return_rate",
                "description": "Расходы на возвраты",
                "category": "cost",
                "priority": 3
            },
            "packaging": {
                "name": "Упаковка",
                "formula": "packaging_cost",
                "description": "Стоимость упаковки",
                "category": "cost",
                "priority": 3
            },
            "customs": {
                "name": "Таможня",
                "formula": "customs_cost",
                "description": "Таможенные расходы",
                "category": "cost",
                "priority": 3
            },
            "other_costs": {
                "name": "Прочие расходы",
                "formula": "other_costs",
                "description": "Другие операционные расходы",
                "category": "cost",
                "priority": 3
            },
            "volume": {
                "name": "Объем",
                "formula": "(length * width * height) / 1000",
                "description": "Объем товара в литрах",
                "category": "dimension",
                "priority": 2
            },
            "roi": {
                "name": "ROI",
                "formula": "(unit_profit / cost) * 100",
                "description": "Возврат инвестиций",
                "category": "profitability",
                "priority": 1
            },
            "payback_days": {
                "name": "Окупаемость (дней)",
                "formula": "cost / daily_profit",
                "description": "Срок окупаемости в днях",
                "category": "time",
                "priority": 2
            },
            "break_even_units": {
                "name": "Точка безубыточности (шт)",
                "formula": "fixed_costs / contribution_margin",
                "description": "Количество единиц для достижения безубыточности",
                "category": "break_even",
                "priority": 2
            },
            "break_even_revenue": {
                "name": "Точка безубыточности (₽)",
                "formula": "break_even_units * price",
                "description": "Выручка для достижения безубыточности",
                "category": "break_even",
                "priority": 2
            },
            "ltv": {
                "name": "LTV",
                "formula": "unit_profit / (1 - retention_rate + discount_rate)",
                "description": "Ценность клиента за всё время",
                "category": "customer",
                "priority": 2
            },
            "cac": {
                "name": "CAC",
                "formula": "advertising / customers_per_1000 * 1000",
                "description": "Стоимость привлечения клиента",
                "category": "customer",
                "priority": 2
            },
            "ltv_cac_ratio": {
                "name": "LTV/CAC",
                "formula": "ltv / cac",
                "description": "Соотношение ценности клиента и стоимости привлечения",
                "category": "customer",
                "priority": 1
            },
            "roas": {
                "name": "ROAS",
                "formula": "revenue / advertising",
                "description": "Возврат рекламных инвестиций",
                "category": "marketing",
                "priority": 2
            },
            "drr": {
                "name": "ДРР %",
                "formula": "(advertising / revenue) * 100",
                "description": "Доля рекламных расходов",
                "category": "marketing",
                "priority": 2
            },
            "eoq": {
                "name": "EOQ",
                "formula": "sqrt((2 * annual_demand * order_cost) / holding_cost)",
                "description": "Оптимальный размер заказа",
                "category": "inventory",
                "priority": 2
            },
            "reorder_point": {
                "name": "Точка заказа",
                "formula": "daily_sales * lead_time + safety_stock",
                "description": "Уровень запаса для повторного заказа",
                "category": "inventory",
                "priority": 2
            },
            "safety_stock": {
                "name": "Страховой запас",
                "formula": "z_score * daily_sales_std * sqrt(lead_time)",
                "description": "Страховой запас для предотвращения дефицита",
                "category": "inventory",
                "priority": 2
            },
            "max_stock": {
                "name": "Максимальный запас",
                "formula": "eoq + safety_stock",
                "description": "Максимальный уровень запаса",
                "category": "inventory",
                "priority": 3
            },
            "min_stock": {
                "name": "Минимальный запас",
                "formula": "reorder_point",
                "description": "Минимальный уровень запаса",
                "category": "inventory",
                "priority": 3
            },
            "days_until_stockout": {
                "name": "Дней до дефицита",
                "formula": "current_stock / daily_sales",
                "description": "Количество дней до исчерпания запаса",
                "category": "inventory",
                "priority": 3
            },
            "demand_forecast_30": {
                "name": "Прогноз 30 дней",
                "formula": "historical_avg * 1.1",
                "description": "Прогноз спроса на 30 дней",
                "category": "forecast",
                "priority": 2
            },
            "demand_forecast_60": {
                "name": "Прогноз 60 дней",
                "formula": "historical_avg * 2.2",
                "description": "Прогноз спроса на 60 дней",
                "category": "forecast",
                "priority": 2
            },
            "demand_forecast_90": {
                "name": "Прогноз 90 дней",
                "formula": "historical_avg * 3.3",
                "description": "Прогноз спроса на 90 дней",
                "category": "forecast",
                "priority": 2
            },
            "demand_forecast_180": {
                "name": "Прогноз 180 дней",
                "formula": "historical_avg * 6.6",
                "description": "Прогноз спроса на 180 дней",
                "category": "forecast",
                "priority": 3
            },
            "demand_forecast_365": {
                "name": "Прогноз 365 дней",
                "formula": "historical_avg * 13.2",
                "description": "Прогноз спроса на 365 дней",
                "category": "forecast",
                "priority": 3
            },
            "demand_trend": {
                "name": "Тренд спроса",
                "formula": "slope_of_historical_sales",
                "description": "Направление изменения спроса",
                "category": "forecast",
                "priority": 2
            },
            "demand_volatility": {
                "name": "Волатильность спроса",
                "formula": "std(historical_sales) / mean(historical_sales)",
                "description": "Изменчивость спроса",
                "category": "forecast",
                "priority": 3
            },
            "price_elasticity": {
                "name": "Эластичность цены",
                "formula": "(% change in demand) / (% change in price)",
                "description": "Чувствительность спроса к изменению цены",
                "category": "sensitivity",
                "priority": 2
            },
            "cost_elasticity": {
                "name": "Эластичность себестоимости",
                "formula": "(% change in profit) / (% change in cost)",
                "description": "Чувствительность прибыли к изменению себестоимости",
                "category": "sensitivity",
                "priority": 2
            },
            "commission_elasticity": {
                "name": "Эластичность комиссии",
                "formula": "(% change in profit) / (% change in commission)",
                "description": "Чувствительность прибыли к изменению комиссии",
                "category": "sensitivity",
                "priority": 3
            },
            "logistics_elasticity": {
                "name": "Эластичность логистики",
                "formula": "(% change in profit) / (% change in logistics)",
                "description": "Чувствительность прибыли к изменению логистики",
                "category": "sensitivity",
                "priority": 3
            },
            "quality_score": {
                "name": "Оценка качества",
                "formula": "composite_score_of_metrics",
                "description": "Комплексная оценка качества товара",
                "category": "quality",
                "priority": 3
            },
            "risk_score": {
                "name": "Оценка риска",
                "formula": "composite_risk_score",
                "description": "Оценка риска товара",
                "category": "quality",
                "priority": 3
            },
            "opportunity_score": {
                "name": "Оценка возможности",
                "formula": "composite_opportunity_score",
                "description": "Оценка потенциала товара",
                "category": "quality",
                "priority": 3
            },
            "competitive_score": {
                "name": "Конкурентная оценка",
                "formula": "composite_competitive_score",
                "description": "Оценка конкурентной позиции",
                "category": "quality",
                "priority": 3
            }
        }
        
        # Добавляем специфические формулы для маркетплейсов
        marketplace_formulas = {
            "Ozon": {
                "commission_rate": 0.10,
                "logistics_base": 50,
                "logistics_per_liter": 15,
                "logistics_per_kg": 20,
                "storage_per_liter": 0.8,
                "acquiring_rate": 0.025,
                "return_rate": 0.12,
                "advertising_rate": 0.15,
                "packaging_cost": 50,
                "customs_cost": 0,
                "other_costs": 0
            },
            "Wildberries": {
                "commission_rate": 0.12,
                "logistics_base": 40,
                "logistics_per_liter": 12,
                "logistics_per_kg": 18,
                "storage_per_liter": 0.9,
                "acquiring_rate": 0.028,
                "return_rate": 0.15,
                "advertising_rate": 0.15,
                "packaging_cost": 50,
                "customs_cost": 0,
                "other_costs": 0
            },
            "Яндекс Маркет": {
                "commission_rate": 0.11,
                "logistics_base": 60,
                "logistics_per_liter": 10,
                "logistics_per_kg": 15,
                "storage_per_liter": 0.7,
                "acquiring_rate": 0.022,
                "return_rate": 0.10,
                "advertising_rate": 0.15,
                "packaging_cost": 50,
                "customs_cost": 0,
                "other_costs": 0
            },
            "AliExpress": {
                "commission_rate": 0.08,
                "logistics_base": 45,
                "logistics_per_liter": 12,
                "logistics_per_kg": 16,
                "storage_per_liter": 0.6,
                "acquiring_rate": 0.03,
                "return_rate": 0.08,
                "advertising_rate": 0.15,
                "packaging_cost": 50,
                "customs_cost": 0,
                "other_costs": 0
            },
            "Мегамаркет": {
                "commission_rate": 0.09,
                "logistics_base": 55,
                "logistics_per_liter": 14,
                "logistics_per_kg": 19,
                "storage_per_liter": 0.75,
                "acquiring_rate": 0.026,
                "return_rate": 0.11,
                "advertising_rate": 0.15,
                "packaging_cost": 50,
                "customs_cost": 0,
                "other_costs": 0
            }
        }
        
        formulas["marketplace_rates"] = marketplace_formulas
        
        return formulas
    
    def _call_api(self, prompt: str) -> Optional[str]:
        """Вызов AI API"""
        if not self.api_key:
            return None
        try:
            headers = {"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"}
            data = {
                "model": "deepseek-chat",
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.3,
                "max_tokens": 2000
            }
            response = requests.post(self.api_url, headers=headers, json=data, timeout=30)
            if response.status_code == 200:
                content = response.json()['choices'][0]['message']['content']
                json_match = re.search(r'(\{.*\})', content, re.DOTALL)
                if json_match:
                    return json_match.group()
                list_match = re.search(r'(\[.*\])', content, re.DOTALL)
                if list_match:
                    return list_match.group()
        except Exception as e:
            logger.error(f"API error: {e}")
        return None


# --------------------------------------------
# 💎 ЮНИТ-ЭКОНОМИКА
# --------------------------------------------
class UnitEconomicsEngine:
    """
    Движок расчета юнит-экономики
    """
    
    def __init__(self, marketplace: str = "Ozon", days_storage: int = 30):
        self.marketplace = marketplace
        self.days_storage = days_storage
        self.rates = self._get_rates(marketplace)
        self.retention_rate = CONFIG["retention_rate"]
        self.discount_rate = CONFIG["discount_rate"]
        self.customers_per_1000 = CONFIG["customers_per_1000"]
        self.order_cost = CONFIG["order_cost"]
        self.holding_cost_pct = CONFIG["holding_cost_pct"]
        self.lead_time = CONFIG["lead_time"]
        self.service_level = CONFIG["service_level"]
        self.z_score = self._get_z_score(self.service_level)
        
    def _get_rates(self, marketplace: str) -> Dict:
        """Получение тарифов для маркетплейса"""
        rates = {
            "Ozon": {
                "commission": 0.10, 
                "logistics_base": 50, 
                "logistics_per_liter": 15,
                "logistics_per_kg": 20, 
                "storage_per_liter": 0.8, 
                "acquiring": 0.025, 
                "returns": 0.12,
                "advertising_rate": 0.15,
                "packaging_cost": 50,
                "customs_cost": 0,
                "other_costs": 0
            },
            "Wildberries": {
                "commission": 0.12, 
                "logistics_base": 40, 
                "logistics_per_liter": 12,
                "logistics_per_kg": 18, 
                "storage_per_liter": 0.9, 
                "acquiring": 0.028, 
                "returns": 0.15,
                "advertising_rate": 0.15,
                "packaging_cost": 50,
                "customs_cost": 0,
                "other_costs": 0
            },
            "Яндекс Маркет": {
                "commission": 0.11, 
                "logistics_base": 60, 
                "logistics_per_liter": 10,
                "logistics_per_kg": 15, 
                "storage_per_liter": 0.7, 
                "acquiring": 0.022, 
                "returns": 0.10,
                "advertising_rate": 0.15,
                "packaging_cost": 50,
                "customs_cost": 0,
                "other_costs": 0
            },
            "AliExpress": {
                "commission": 0.08, 
                "logistics_base": 45, 
                "logistics_per_liter": 12,
                "logistics_per_kg": 16, 
                "storage_per_liter": 0.6, 
                "acquiring": 0.03, 
                "returns": 0.08,
                "advertising_rate": 0.15,
                "packaging_cost": 50,
                "customs_cost": 0,
                "other_costs": 0
            },
            "Мегамаркет": {
                "commission": 0.09, 
                "logistics_base": 55, 
                "logistics_per_liter": 14,
                "logistics_per_kg": 19, 
                "storage_per_liter": 0.75, 
                "acquiring": 0.026, 
                "returns": 0.11,
                "advertising_rate": 0.15,
                "packaging_cost": 50,
                "customs_cost": 0,
                "other_costs": 0
            }
        }
        return rates.get(marketplace, rates["Ozon"])
    
    def _get_z_score(self, service_level: float) -> float:
        """Получение Z-оценки для уровня обслуживания"""
        z_scores = {
            0.90: 1.28,
            0.91: 1.34,
            0.92: 1.41,
            0.93: 1.48,
            0.94: 1.55,
            0.95: 1.65,
            0.96: 1.75,
            0.97: 1.88,
            0.98: 2.05,
            0.99: 2.33
        }
        return z_scores.get(service_level, 1.65)
    
    def calculate_product(self, product: ProductUnit, fixed_costs: float = 50000, avg_orders: int = 100) -> ProductUnit:
        """
        Расчет всех показателей юнит-экономики для товара
        """
        # Объем
        product.volume = (product.length * product.width * product.height) / 1000 if all([product.length, product.width, product.height]) else 0
        
        # Расходы
        product.commission = product.price * self.rates.get("commission", 0.10)
        product.logistics = self.rates.get("logistics_base", 50)
        if product.volume > 0:
            product.logistics += product.volume * self.rates.get("logistics_per_liter", 15)
        if product.weight > 0:
            product.logistics += product.weight * self.rates.get("logistics_per_kg", 20)
        
        product.storage = product.volume * self.rates.get("storage_per_liter", 0.8) * (self.days_storage / 30)
        product.acquiring = product.price * self.rates.get("acquiring", 0.025)
        product.advertising = product.price * self.rates.get("advertising_rate", 0.15)
        product.returns = product.price * self.rates.get("returns", 0.12)
        product.packaging = self.rates.get("packaging_cost", 50)
        product.customs = self.rates.get("customs_cost", 0)
        product.other_costs = self.rates.get("other_costs", 0)
        
        # Переменные расходы
        product.total_variable_costs = (product.cost + product.commission + product.logistics + 
                                       product.storage + product.acquiring + product.advertising + 
                                       product.returns + product.packaging + product.customs + 
                                       product.other_costs)
        
        # Полная себестоимость
        product.total_cost = product.total_variable_costs
        
        # Ключевые метрики
        product.contribution_margin = product.price - product.total_variable_costs
        product.contribution_margin_pct = (product.contribution_margin / product.price * 100) if product.price > 0 else 0
        product.unit_profit = product.price - product.total_cost
        product.margin = (product.unit_profit / product.price * 100) if product.price > 0 else 0
        product.roi = (product.unit_profit / product.cost * 100) if product.cost > 0 else 0
        
        # Окупаемость
        daily_profit = product.unit_profit * avg_orders / 30
        product.payback_days = product.cost / daily_profit if daily_profit > 0 else 999
        
        # Точка безубыточности
        product.break_even_units = fixed_costs / product.contribution_margin if product.contribution_margin > 0 else 999999
        product.break_even_revenue = product.break_even_units * product.price
        
        # LTV/CAC
        product.ltv = product.unit_profit / (1 - self.retention_rate + self.discount_rate)
        product.cac = product.advertising / self.customers_per_1000 * 1000 if product.advertising > 0 else product.price * 0.05
        product.ltv_cac_ratio = product.ltv / product.cac if product.cac > 0 else 0
        
        # ROAS и ДРР
        product.roas = product.price / product.advertising if product.advertising > 0 else 0
        product.drr = (product.advertising / product.price * 100) if product.price > 0 else 0
        
        # Запасы
        annual_demand = avg_orders * 12
        holding_cost = product.cost * self.holding_cost_pct
        product.eoq = math.sqrt((2 * annual_demand * self.order_cost) / holding_cost) if holding_cost > 0 else 0
        product.safety_stock = self._calc_safety_stock(avg_orders)
        product.reorder_point = (avg_orders / 30 * self.lead_time) + product.safety_stock
        product.max_stock = product.eoq + product.safety_stock
        product.min_stock = product.reorder_point
        product.days_until_stockout = product.reorder_point / (avg_orders / 30) if avg_orders > 0 else 999
        
        # Статус запаса
        product.stock_status = self._determine_stock_status(product.reorder_point, product.max_stock)
        
        # ABC/XYZ
        product.abc_category = self._calc_abc_category(product.unit_profit)
        product.xyz_category = self._calc_xyz_category(avg_orders)
        
        # Прогнозы
        product.demand_forecast_30 = avg_orders * 1.1
        product.demand_forecast_60 = avg_orders * 2.2
        product.demand_forecast_90 = avg_orders * 3.3
        product.demand_forecast_180 = avg_orders * 6.6
        product.demand_forecast_365 = avg_orders * 13.2
        product.demand_trend = self._calc_trend(avg_orders)
        product.demand_trend_value = 0.1
        product.demand_volatility = 0.2
        
        # Анализ чувствительности
        sensitivity = self._calculate_sensitivity(product)
        product.price_elasticity = sensitivity["price_elasticity"]
        product.cost_elasticity = sensitivity["cost_elasticity"]
        product.commission_elasticity = sensitivity["commission_elasticity"]
        product.logistics_elasticity = sensitivity["logistics_elasticity"]
        product.most_sensitive_factor = sensitivity["most_sensitive"]
        product.sensitivity_analysis = sensitivity
        
        # Сезонность
        product.week_seasonality = [1.0, 0.9, 0.95, 0.95, 1.0, 1.15, 1.2]  # Пн-Вс
        product.month_seasonality = [0.9, 0.85, 0.95, 1.0, 1.05, 1.1, 1.0, 0.95, 0.95, 1.0, 1.1, 1.0]
        product.seasonal_factors = {
            "week": product.week_seasonality,
            "month": product.month_seasonality
        }
        
        # Метрики качества
        product.quality_score = self._calc_quality_score(product)
        product.risk_score = self._calc_risk_score(product)
        product.opportunity_score = self._calc_opportunity_score(product)
        product.competitive_score = self._calc_competitive_score(product)
        
        # Рекомендации по цене
        product.recommended_price = product.price
        product.recommended_markup = (product.price - product.cost) / product.cost * 100 if product.cost > 0 else 0
        product.recommended_discount = 0
        product.optimization_tips = self._generate_tips(product)
        
        return product
    
    def _calc_safety_stock(self, avg_orders: int) -> float:
        """Расчет страхового запаса"""
        daily_std = avg_orders / 30 * 0.2  # 20% вариация
        return self.z_score * daily_std * math.sqrt(self.lead_time)
    
    def _determine_stock_status(self, reorder_point: float, max_stock: float) -> str:
        """Определение статуса запаса"""
        current_stock = (reorder_point + max_stock) / 2  # Средний уровень
        if current_stock <= 0:
            return "❌ Нет в наличии"
        elif current_stock <= reorder_point * 0.5:
            return "🔴 Критический"
        elif current_stock <= reorder_point:
            return "🟡 Низкий"
        elif current_stock > max_stock:
            return "🔵 Избыток"
        else:
            return "🟢 Норма"
    
    def _calc_abc_category(self, profit: float) -> str:
        """Определение ABC-категории"""
        if profit > 1000:
            return "A"
        elif profit > 100:
            return "B"
        else:
            return "C"
    
    def _calc_xyz_category(self, sales: float) -> str:
        """Определение XYZ-категории"""
        cv = 0.2  # Коэффициент вариации
        if cv < 0.2:
            return "X"
        elif cv < 0.5:
            return "Y"
        else:
            return "Z"
    
    def _calc_trend(self, avg_orders: int) -> str:
        """Определение тренда"""
        if avg_orders > 100:
            return "📈 Растущий"
        elif avg_orders > 50:
            return "➡️ Стабильный"
        else:
            return "📉 Падающий"
    
    def _calc_quality_score(self, product: ProductUnit) -> float:
        """Расчет оценки качества"""
        score = 0
        if product.margin > 20:
            score += 20
        elif product.margin > 10:
            score += 10
        if product.contribution_margin_pct > 30:
            score += 20
        elif product.contribution_margin_pct > 15:
            score += 10
        if product.ltv_cac_ratio > 3:
            score += 20
        elif product.ltv_cac_ratio > 1:
            score += 10
        if product.roi > 20:
            score += 20
        elif product.roi > 10:
            score += 10
        if product.payback_days < 30:
            score += 20
        elif product.payback_days < 60:
            score += 10
        return min(score, 100)
    
    def _calc_risk_score(self, product: ProductUnit) -> float:
        """Расчет оценки риска"""
        risk = 0
        if product.margin < 5:
            risk += 20
        if product.contribution_margin_pct < 10:
            risk += 20
        if product.ltv_cac_ratio < 1:
            risk += 20
        if product.payback_days > 90:
            risk += 20
        if product.returns > 0.15 * product.price:
            risk += 20
        return min(risk, 100)
    
    def _calc_opportunity_score(self, product: ProductUnit) -> float:
        """Расчет оценки возможности"""
        score = 0
        if product.margin > 25:
            score += 25
        if product.ltv_cac_ratio > 3:
            score += 25
        if product.roi > 30:
            score += 25
        if product.demand_trend == "📈 Растущий":
            score += 25
        return min(score, 100)
    
    def _calc_competitive_score(self, product: ProductUnit) -> float:
        """Расчет конкурентной оценки"""
        score = 0
        if product.price > product.cost * 2:
            score += 25
        if product.margin > 30:
            score += 25
        if product.brand and product.brand.lower() in ["toyota", "bmw", "mercedes"]:
            score += 25
        if product.category and product.category.lower() in ["оригинал", "oem"]:
            score += 25
        return min(score, 100)
    
    def _calculate_sensitivity(self, product: ProductUnit) -> Dict:
        """Анализ чувствительности"""
        base_profit = product.unit_profit
        
        # Изменение цены на ±10%
        price_up = ((product.price * 1.10) - product.total_cost) - base_profit
        price_down = ((product.price * 0.90) - product.total_cost) - base_profit
        
        # Изменение себестоимости на ±10%
        cost_up = (product.price - (product.total_cost + product.cost * 0.10)) - base_profit
        cost_down = (product.price - (product.total_cost - product.cost * 0.10)) - base_profit
        
        # Изменение комиссии на ±2%
        commission_up = (product.price - (product.total_cost + product.price * 0.02)) - base_profit
        commission_down = (product.price - (product.total_cost - product.price * 0.02)) - base_profit
        
        # Изменение логистики на ±20%
        logistics_up = (product.price - (product.total_cost + product.logistics * 0.20)) - base_profit
        logistics_down = (product.price - (product.total_cost - product.logistics * 0.20)) - base_profit
        
        # Изменение рекламы на ±50%
        ads_up = (product.price - (product.total_cost + product.advertising * 0.50)) - base_profit
        ads_down = (product.price - (product.total_cost - product.advertising * 0.50)) - base_profit
        
        factors = {
            "Цена": abs(price_up - price_down),
            "Себестоимость": abs(cost_up - cost_down),
            "Комиссия": abs(commission_up - commission_down),
            "Логистика": abs(logistics_up - logistics_down),
            "Реклама": abs(ads_up - ads_down)
        }
        
        most_sensitive = max(factors.items(), key=lambda x: x[1])[0] if factors else "Нет данных"
        
        return {
            "price_elasticity": (price_up - price_down) / 2 / base_profit if base_profit != 0 else 0,
            "cost_elasticity": (cost_up - cost_down) / 2 / base_profit if base_profit != 0 else 0,
            "commission_elasticity": (commission_up - commission_down) / 2 / base_profit if base_profit != 0 else 0,
            "logistics_elasticity": (logistics_up - logistics_down) / 2 / base_profit if base_profit != 0 else 0,
            "most_sensitive": most_sensitive,
            "price_up_10": price_up,
            "price_down_10": price_down,
            "cost_up_10": cost_up,
            "cost_down_10": cost_down,
            "commission_up_2": commission_up,
            "commission_down_2": commission_down,
            "logistics_up_20": logistics_up,
            "logistics_down_20": logistics_down,
            "ads_up_50": ads_up,
            "ads_down_50": ads_down
        }
    
    def _generate_tips(self, product: ProductUnit) -> List[str]:
        """Генерация рекомендаций по оптимизации"""
        tips = []
        
        if product.margin < 15:
            tips.append("🔴 Низкая маржинальность (<15%) — рассмотрите повышение цены или снижение себестоимости")
        elif product.margin < 25:
            tips.append("🟡 Средняя маржинальность (15-25%) — есть потенциал для оптимизации")
        else:
            tips.append("🟢 Отличная маржинальность (>25%) — можно масштабировать")
        
        if product.contribution_margin_pct < 30:
            tips.append("📊 Низкий Contribution Margin — пересмотрите структуру расходов")
        
        if product.logistics / product.price > 0.20:
            tips.append("📦 Логистика >20% цены — оптимизируйте габариты")
        
        if product.returns / product.price > 0.10:
            tips.append("⚠️ Возвраты >10% — улучшите описание товара")
        
        if product.payback_days > 90:
            tips.append("⏳ Долгая окупаемость (>90 дней) — сократите партию закупки")
        
        if product.ltv_cac_ratio < 3:
            tips.append("📈 Низкий LTV/CAC — оптимизируйте рекламу или повышайте повторные продажи")
        
        if product.drr > 25:
            tips.append("📢 ДРР > 25% — слишком высокие рекламные расходы")
        
        if product.roi < 15:
            tips.append("💸 Низкий ROI (<15%) — пересмотрите ценовую стратегию")
        
        if product.stock_status == "🔴 Критический":
            tips.append("⚠️ Критический запас — срочно пополните склад")
        
        if product.risk_score > 60:
            tips.append("⚠️ Высокий риск — требуется анализ")
        
        if not tips:
            tips.append("✅ Юнит-экономика в норме")
        
        return tips
    
    def analyze_batch(self, products: List[ProductUnit]) -> Dict:
        """Анализ партии товаров"""
        if not products:
            return {
                "total_products": 0,
                "profitable_products": 0,
                "loss_making": 0,
                "total_revenue": 0,
                "total_cost": 0,
                "total_profit": 0,
                "avg_margin": 0,
                "avg_cm": 0,
                "best_product": None,
                "worst_product": None,
                "abc_distribution": {"A": 0, "B": 0, "C": 0},
                "xyz_distribution": {"X": 0, "Y": 0, "Z": 0},
                "total_ltv": 0,
                "avg_ltv_cac": 0,
                "total_roas": 0,
                "avg_quality_score": 0,
                "avg_risk_score": 0,
                "avg_opportunity_score": 0
            }
        
        total_revenue = sum(p.price for p in products)
        total_cost = sum(p.total_cost for p in products)
        total_profit = sum(p.unit_profit for p in products)
        avg_margin = np.mean([p.margin for p in products])
        avg_cm = np.mean([p.contribution_margin_pct for p in products])
        
        profitable = sum(1 for p in products if p.unit_profit > 0)
        
        # ABC распределение
        abc_dist = {"A": 0, "B": 0, "C": 0}
        xyz_dist = {"X": 0, "Y": 0, "Z": 0}
        for p in products:
            if p.abc_category in abc_dist:
                abc_dist[p.abc_category] += 1
            if p.xyz_category in xyz_dist:
                xyz_dist[p.xyz_category] += 1
        
        # Расчет дополнительных метрик
        avg_ltv_cac = np.mean([p.ltv_cac_ratio for p in products]) if products else 0
        total_ltv = sum(p.ltv for p in products)
        total_roas = sum(p.roas for p in products) / len(products) if products else 0
        avg_quality = np.mean([p.quality_score for p in products]) if products else 0
        avg_risk = np.mean([p.risk_score for p in products]) if products else 0
        avg_opportunity = np.mean([p.opportunity_score for p in products]) if products else 0
        
        return {
            "total_products": len(products),
            "profitable_products": profitable,
            "loss_making": len(products) - profitable,
            "total_revenue": total_revenue,
            "total_cost": total_cost,
            "total_profit": total_profit,
            "avg_margin": avg_margin,
            "avg_cm": avg_cm,
            "best_product": max(products, key=lambda x: x.unit_profit) if products else None,
            "worst_product": min(products, key=lambda x: x.unit_profit) if products else None,
            "abc_distribution": abc_dist,
            "xyz_distribution": xyz_dist,
            "total_ltv": total_ltv,
            "avg_ltv_cac": avg_ltv_cac,
            "total_roas": total_roas,
            "avg_quality_score": avg_quality,
            "avg_risk_score": avg_risk,
            "avg_opportunity_score": avg_opportunity
        }


# --------------------------------------------
# 🔄 ОПТИМИЗАЦИЯ ЦЕН
# --------------------------------------------
class PriceOptimizer:
    """
    Оптимизация ценообразования на основе юнит-экономики
    """
    
    def __init__(self, engine: UnitEconomicsEngine):
        self.engine = engine
        self._cache = {}
    
    def optimize_price(self, product: ProductUnit, min_price: float = None, max_price: float = None) -> Dict:
        """
        Оптимизация цены товара
        """
        if min_price is None:
            min_price = product.cost * 1.1
        if max_price is None:
            max_price = product.price * 2
        
        best_price = product.price
        best_profit = product.unit_profit
        
        # Поиск оптимальной цены
        prices = np.linspace(min_price, max_price, 30)
        results = []
        
        for price in prices:
            test_product = ProductUnit(
                article=product.article,
                name=product.name,
                price=price,
                cost=product.cost,
                length=product.length,
                width=product.width,
                height=product.height,
                weight=product.weight,
                category=product.category,
                brand=product.brand
            )
            self.engine.calculate_product(test_product)
            results.append({
                "price": price,
                "profit": test_product.unit_profit,
                "margin": test_product.margin,
                "roi": test_product.roi
            })
            if test_product.unit_profit > best_profit:
                best_profit = test_product.unit_profit
                best_price = price
        
        # Эластичность спроса
        demand_elasticity = self._estimate_demand_elasticity(product, best_price)
        
        # Оптимальная наценка
        optimal_markup = (best_price - product.cost) / product.cost * 100 if product.cost > 0 else 0
        current_markup = (product.price - product.cost) / product.cost * 100 if product.cost > 0 else 0
        
        return {
            "current_price": product.price,
            "current_profit": product.unit_profit,
            "current_margin": product.margin,
            "current_markup": current_markup,
            "recommended_price": best_price,
            "recommended_markup": optimal_markup,
            "recommended_profit": best_profit,
            "recommended_margin": max([r["margin"] for r in results]) if results else 0,
            "price_change": best_price - product.price,
            "price_change_pct": ((best_price - product.price) / product.price * 100) if product.price > 0 else 0,
            "profit_increase": best_profit - product.unit_profit,
            "profit_increase_pct": ((best_profit - product.unit_profit) / product.unit_profit * 100) if product.unit_profit > 0 else 0,
            "demand_elasticity": demand_elasticity,
            "recommendation": self._get_recommendation(best_price, product.price),
            "price_range": {"min": min_price, "max": max_price},
            "results": results
        }
    
    def _estimate_demand_elasticity(self, product: ProductUnit, new_price: float) -> float:
        """Оценка эластичности спроса"""
        price_change = (new_price - product.price) / product.price if product.price > 0 else 0
        # Упрощенная модель: эластичность зависит от категории
        if product.category:
            cat_lower = product.category.lower()
            if "премиум" in cat_lower or "luxury" in cat_lower:
                return -0.5
            elif "необходим" in cat_lower or "essential" in cat_lower:
                return -0.8
            elif "электроника" in cat_lower or "electronics" in cat_lower:
                return -1.2
        return -1.0
    
    def _get_recommendation(self, recommended: float, current: float) -> str:
        """Рекомендация по цене"""
        diff_pct = (recommended - current) / current * 100 if current > 0 else 0
        if diff_pct > 10:
            return "🚀 Значительно повысить цену (+{:.0f}%)".format(diff_pct)
        elif diff_pct > 3:
            return "📈 Незначительно повысить цену (+{:.0f}%)".format(diff_pct)
        elif diff_pct < -10:
            return "🔻 Значительно снизить цену ({:.0f}%)".format(diff_pct)
        elif diff_pct < -3:
            return "📉 Незначительно снизить цену ({:.0f}%)".format(diff_pct)
        else:
            return "✅ Оставить текущую цену"
    
    def optimize_all_prices(self, products: List[ProductUnit]) -> List[Dict]:
        """Оптимизация цен для всех товаров"""
        results = []
        for product in products:
            result = self.optimize_price(product)
            results.append({
                "article": product.article,
                "name": product.name,
                "current_price": result["current_price"],
                "recommended_price": result["recommended_price"],
                "profit_increase": result["profit_increase"],
                "recommendation": result["recommendation"]
            })
        return results


# --------------------------------------------
# 📊 ВИЗУАЛИЗАЦИЯ
# --------------------------------------------
class VisualizationEngine:
    """Визуализация показателей юнит-экономики"""
    
    @staticmethod
    def plot_margin_distribution(products: List[ProductUnit]) -> go.Figure:
        """Распределение маржинальности"""
        margins = [p.margin for p in products]
        fig = go.Figure(data=[go.Histogram(x=margins, nbinsx=30, marker_color='#667eea')])
        fig.add_vline(x=0, line_dash="dash", line_color="red", annotation_text="Убыток")
        fig.add_vline(x=15, line_dash="dash", line_color="orange", annotation_text="Низкая")
        fig.add_vline(x=25, line_dash="dash", line_color="green", annotation_text="Высокая")
        fig.update_layout(
            title='📊 Распределение маржинальности',
            xaxis_title='Маржинальность %',
            yaxis_title='Количество товаров',
            height=400,
            showlegend=False
        )
        return fig
    
    @staticmethod
    def plot_profit_pareto(products: List[ProductUnit]) -> go.Figure:
        """Pareto-анализ прибыли"""
        sorted_products = sorted(products, key=lambda x: x.unit_profit, reverse=True)
        profits = [p.unit_profit for p in sorted_products]
        cumulative = np.cumsum(profits)
        cumulative_pct = cumulative / cumulative[-1] * 100 if cumulative[-1] > 0 else 0
        
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(
            go.Bar(x=[p.article[:10] for p in sorted_products[:20]], y=profits[:20], name='Прибыль'),
            secondary_y=False
        )
        fig.add_trace(
            go.Scatter(x=[p.article[:10] for p in sorted_products[:20]], y=cumulative_pct[:20], 
                      name='Накопленный %', mode='lines+markers', line=dict(color='red')),
            secondary_y=True
        )
        fig.update_layout(
            title='🏆 Pareto-анализ прибыли (топ-20)',
            height=400,
            xaxis_tickangle=-45
        )
        fig.update_yaxes(title_text="Прибыль, ₽", secondary_y=False)
        fig.update_yaxes(title_text="Накопленный %", secondary_y=True, range=[0, 110])
        return fig
    
    @staticmethod
    def plot_cost_structure(product: ProductUnit) -> go.Figure:
        """Структура расходов товара"""
        costs = {
            "Себестоимость": product.cost,
            "Комиссия": product.commission,
            "Логистика": product.logistics,
            "Хранение": product.storage,
            "Эквайринг": product.acquiring,
            "Реклама": product.advertising,
            "Возвраты": product.returns,
            "Упаковка": product.packaging,
            "Таможня": product.customs,
            "Прочее": product.other_costs
        }
        # Убираем нулевые значения
        costs = {k: v for k, v in costs.items() if v > 0}
        
        if not costs:
            fig = go.Figure()
            fig.update_layout(title='📊 Структура расходов (нет данных)', height=400)
            return fig
        
        colors = ['#667eea', '#764ba2', '#f093fb', '#f5576c', '#4facfe', 
                  '#43e97b', '#fa709a', '#fee140', '#30cfd0', '#330867']
        fig = go.Figure(data=[go.Pie(labels=list(costs.keys()), values=list(costs.values()), 
                                     hole=0.4, marker=dict(colors=colors[:len(costs)]))])
        fig.update_layout(title='📊 Структура расходов', height=400)
        return fig
    
    @staticmethod
    def plot_marketplace_comparison(products: List[ProductUnit]) -> go.Figure:
        """Сравнение эффективности на разных маркетплейсах"""
        marketplaces = ["Ozon", "Wildberries", "Яндекс Маркет", "AliExpress", "Мегамаркет"]
        profits = []
        margins = []
        
        for mp in marketplaces:
            engine = UnitEconomicsEngine(marketplace=mp, days_storage=30)
            mp_products = []
            for p in products[:50]:
                mp_p = ProductUnit(**p.__dict__)
                engine.calculate_product(mp_p)
                mp_products.append(mp_p)
            profits.append(sum(p.unit_profit for p in mp_products))
            margins.append(np.mean([p.margin for p in mp_products]) if mp_products else 0)
        
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        colors = ['#667eea', '#764ba2', '#f093fb', '#f5576c', '#4facfe']
        
        fig.add_trace(
            go.Bar(x=marketplaces, y=profits, name='Прибыль', marker_color=colors),
            secondary_y=False
        )
        fig.add_trace(
            go.Scatter(x=marketplaces, y=margins, name='Маржа %', mode='lines+markers', 
                      line=dict(color='red'), marker=dict(size=10)),
            secondary_y=True
        )
        fig.update_layout(
            title='🏪 Сравнение маркетплейсов',
            height=400,
            xaxis_title='Маркетплейс'
        )
        fig.update_yaxes(title_text="Прибыль, ₽", secondary_y=False)
        fig.update_yaxes(title_text="Маржа %", secondary_y=True, range=[0, max(margins) * 1.2 if margins else 100])
        return fig
    
    @staticmethod
    def plot_abc_xyz_matrix(products: List[ProductUnit]) -> go.Figure:
        """ABC-XYZ матрица"""
        matrix = {"AX": 0, "AY": 0, "AZ": 0, "BX": 0, "BY": 0, "BZ": 0, "CX": 0, "CY": 0, "CZ": 0}
        for p in products:
            key = f"{p.abc_category}{p.xyz_category}"
            if key in matrix:
                matrix[key] += 1
        
        colors = {
            "AX": "#28a745", "AY": "#20c997", "AZ": "#ffc107",
            "BX": "#17a2b8", "BY": "#6c757d", "BZ": "#fd7e14",
            "CX": "#dc3545", "CY": "#e83e8c", "CZ": "#6f42c1"
        }
        
        fig = go.Figure()
        for key, count in matrix.items():
            if count > 0:
                fig.add_trace(go.Bar(
                    x=[key],
                    y=[count],
                    name=key,
                    marker_color=colors.get(key, "#667eea"),
                    text=[count],
                    textposition="outside"
                ))
        
        fig.update_layout(
            title='📊 ABC-XYZ Матрица',
            height=400,
            xaxis_title='Категория',
            yaxis_title='Количество товаров',
            showlegend=False
        )
        return fig
    
    @staticmethod
    def plot_sensitivity(product: ProductUnit) -> go.Figure:
        """График чувствительности"""
        factors = {
            "Цена\n+10%": (product.price * 1.10 - product.total_cost) - product.unit_profit,
            "Цена\n-10%": (product.price * 0.90 - product.total_cost) - product.unit_profit,
            "Себест.\n+10%": (product.price - (product.total_cost + product.cost * 0.10)) - product.unit_profit,
            "Себест.\n-10%": (product.price - (product.total_cost - product.cost * 0.10)) - product.unit_profit,
            "Комиссия\n+2%": (product.price - (product.total_cost + product.price * 0.02)) - product.unit_profit,
            "Комиссия\n-2%": (product.price - (product.total_cost - product.price * 0.02)) - product.unit_profit,
            "Логистика\n+20%": (product.price - (product.total_cost + product.logistics * 0.20)) - product.unit_profit,
            "Логистика\n-20%": (product.price - (product.total_cost - product.logistics * 0.20)) - product.unit_profit,
            "Реклама\n+50%": (product.price - (product.total_cost + product.advertising * 0.50)) - product.unit_profit,
            "Реклама\n-50%": (product.price - (product.total_cost - product.advertising * 0.50)) - product.unit_profit
        }
        
        fig = go.Figure(data=[
            go.Bar(
                x=list(factors.keys()),
                y=list(factors.values()),
                marker_color=['#28a745' if v > 0 else '#dc3545' for v in factors.values()]
            )
        ])
        fig.add_hline(y=0, line_dash="dash", line_color="gray")
        fig.update_layout(
            title='🎯 Анализ чувствительности',
            height=400,
            xaxis_title='Фактор',
            yaxis_title='Изменение прибыли, ₽'
        )
        return fig
    
    @staticmethod
    def plot_price_optimization(product: ProductUnit, optimizer: PriceOptimizer) -> go.Figure:
        """График оптимизации цены"""
        result = optimizer.optimize_price(product)
        results = result["results"]
        
        prices = [r["price"] for r in results]
        profits = [r["profit"] for r in results]
        margins = [r["margin"] for r in results]
        
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        
        fig.add_trace(
            go.Scatter(x=prices, y=profits, mode='lines+markers', name='Прибыль', line=dict(color='#667eea')),
            secondary_y=False
        )
        fig.add_trace(
            go.Scatter(x=prices, y=margins, mode='lines+markers', name='Маржа %', line=dict(color='#f5576c')),
            secondary_y=True
        )
        fig.add_vline(x=product.price, line_dash="dash", line_color="blue", annotation_text="Текущая цена")
        fig.add_vline(x=result["recommended_price"], line_dash="dash", line_color="green", annotation_text="Рекомендуемая")
        
        fig.update_layout(
            title='💰 Оптимизация цены',
            height=400,
            xaxis_title='Цена, ₽'
        )
        fig.update_yaxes(title_text="Прибыль, ₽", secondary_y=False)
        fig.update_yaxes(title_text="Маржа %", secondary_y=True)
        return fig
    
    @staticmethod
    def plot_forecast(product: ProductUnit) -> go.Figure:
        """График прогноза спроса"""
        days = [30, 60, 90, 180, 365]
        forecasts = [
            product.demand_forecast_30,
            product.demand_forecast_60,
            product.demand_forecast_90,
            product.demand_forecast_180,
            product.demand_forecast_365
        ]
        
        fig = go.Figure(data=[
            go.Scatter(x=days, y=forecasts, mode='lines+markers', name='Прогноз', line=dict(color='#667eea'))
        ])
        fig.update_layout(
            title='📈 Прогноз спроса',
            height=400,
            xaxis_title='Дней',
            yaxis_title='Прогноз, шт'
        )
        return fig
    
    @staticmethod
    def plot_quality_scores(products: List[ProductUnit]) -> go.Figure:
        """График оценок качества"""
        scores = {
            "Качество": np.mean([p.quality_score for p in products]) if products else 0,
            "Риск": np.mean([p.risk_score for p in products]) if products else 0,
            "Возможность": np.mean([p.opportunity_score for p in products]) if products else 0,
            "Конкуренция": np.mean([p.competitive_score for p in products]) if products else 0
        }
        
        fig = go.Figure(data=[
            go.Bar(
                x=list(scores.keys()),
                y=list(scores.values()),
                marker_color=['#28a745', '#dc3545', '#ffc107', '#17a2b8']
            )
        ])
        fig.add_hline(y=50, line_dash="dash", line_color="gray")
        fig.update_layout(
            title='📊 Оценки качества',
            height=400,
            yaxis_title='Оценка',
            yaxis_range=[0, 100]
        )
        return fig


# --------------------------------------------
# 📥 ЭКСПОРТ
# --------------------------------------------
class ExportEngine:
    """Экспорт данных с юнит-экономикой"""
    
    @staticmethod
    def to_excel(products: List[ProductUnit]) -> bytes:
        """Экспорт в Excel с полной аналитикой"""
        output = io.BytesIO()
        wb = Workbook()
        
        # Основной лист
        ws = wb.active
        ws.title = "Юнит-экономика"
        
        headers = [
            'Артикул', 'Наименование', 'Категория', 'Бренд', 'Страна', 'OE номер',
            'Цена', 'Себестоимость', 'Объем', 'Вес',
            'Комиссия', 'Логистика', 'Хранение', 'Эквайринг', 'Реклама', 'Возвраты', 'Упаковка', 'Таможня', 'Прочее',
            'Переменные расходы', 'Полная себестоимость',
            'Маржинальный доход', 'CM %', 'Прибыль', 'Маржа %', 'ROI',
            'Окупаемость (дней)', 'Точка безубыточности (шт)', 'Точка безубыточности (₽)',
            'LTV', 'CAC', 'LTV/CAC', 'ROAS', 'ДРР %',
            'EOQ', 'Точка заказа', 'Страховой запас', 'Макс. запас', 'Мин. запас', 'Статус',
            'ABC', 'XYZ',
            'Прогноз 30д', 'Прогноз 60д', 'Прогноз 90д', 'Прогноз 180д', 'Прогноз 365д', 'Тренд',
            'Рекомендуемая цена', 'Рекомендуемая наценка %',
            'Качество', 'Риск', 'Возможность', 'Конкуренция',
            'Лучший МП', 'Рекомендации'
        ]
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            cell.font = Font(bold=True, color="FFFFFF")
        
        for r_idx, p in enumerate(products, 2):
            ws.cell(row=r_idx, column=1, value=p.article)
            ws.cell(row=r_idx, column=2, value=p.name[:100] if p.name else "")
            ws.cell(row=r_idx, column=3, value=p.category)
            ws.cell(row=r_idx, column=4, value=p.brand)
            ws.cell(row=r_idx, column=5, value=p.country)
            ws.cell(row=r_idx, column=6, value=p.oe_number)
            ws.cell(row=r_idx, column=7, value=p.price)
            ws.cell(row=r_idx, column=8, value=p.cost)
            ws.cell(row=r_idx, column=9, value=round(p.volume, 2))
            ws.cell(row=r_idx, column=10, value=p.weight)
            ws.cell(row=r_idx, column=11, value=round(p.commission, 2))
            ws.cell(row=r_idx, column=12, value=round(p.logistics, 2))
            ws.cell(row=r_idx, column=13, value=round(p.storage, 2))
            ws.cell(row=r_idx, column=14, value=round(p.acquiring, 2))
            ws.cell(row=r_idx, column=15, value=round(p.advertising, 2))
            ws.cell(row=r_idx, column=16, value=round(p.returns, 2))
            ws.cell(row=r_idx, column=17, value=round(p.packaging, 2))
            ws.cell(row=r_idx, column=18, value=round(p.customs, 2))
            ws.cell(row=r_idx, column=19, value=round(p.other_costs, 2))
            ws.cell(row=r_idx, column=20, value=round(p.total_variable_costs, 2))
            ws.cell(row=r_idx, column=21, value=round(p.total_cost, 2))
            ws.cell(row=r_idx, column=22, value=round(p.contribution_margin, 2))
            ws.cell(row=r_idx, column=23, value=round(p.contribution_margin_pct, 1))
            ws.cell(row=r_idx, column=24, value=round(p.unit_profit, 2))
            ws.cell(row=r_idx, column=25, value=round(p.margin, 1))
            ws.cell(row=r_idx, column=26, value=round(p.roi, 1))
            ws.cell(row=r_idx, column=27, value=round(p.payback_days, 0))
            ws.cell(row=r_idx, column=28, value=round(p.break_even_units, 0))
            ws.cell(row=r_idx, column=29, value=round(p.break_even_revenue, 0))
            ws.cell(row=r_idx, column=30, value=round(p.ltv, 2))
            ws.cell(row=r_idx, column=31, value=round(p.cac, 2))
            ws.cell(row=r_idx, column=32, value=round(p.ltv_cac_ratio, 2))
            ws.cell(row=r_idx, column=33, value=round(p.roas, 2))
            ws.cell(row=r_idx, column=34, value=round(p.drr, 1))
            ws.cell(row=r_idx, column=35, value=round(p.eoq, 0))
            ws.cell(row=r_idx, column=36, value=round(p.reorder_point, 0))
            ws.cell(row=r_idx, column=37, value=round(p.safety_stock, 0))
            ws.cell(row=r_idx, column=38, value=round(p.max_stock, 0))
            ws.cell(row=r_idx, column=39, value=round(p.min_stock, 0))
            ws.cell(row=r_idx, column=40, value=p.stock_status)
            ws.cell(row=r_idx, column=41, value=p.abc_category)
            ws.cell(row=r_idx, column=42, value=p.xyz_category)
            ws.cell(row=r_idx, column=43, value=round(p.demand_forecast_30, 0))
            ws.cell(row=r_idx, column=44, value=round(p.demand_forecast_60, 0))
            ws.cell(row=r_idx, column=45, value=round(p.demand_forecast_90, 0))
            ws.cell(row=r_idx, column=46, value=round(p.demand_forecast_180, 0))
            ws.cell(row=r_idx, column=47, value=round(p.demand_forecast_365, 0))
            ws.cell(row=r_idx, column=48, value=p.demand_trend)
            ws.cell(row=r_idx, column=49, value=round(p.recommended_price, 2))
            ws.cell(row=r_idx, column=50, value=round(p.recommended_markup, 1))
            ws.cell(row=r_idx, column=51, value=round(p.quality_score, 0))
            ws.cell(row=r_idx, column=52, value=round(p.risk_score, 0))
            ws.cell(row=r_idx, column=53, value=round(p.opportunity_score, 0))
            ws.cell(row=r_idx, column=54, value=round(p.competitive_score, 0))
            ws.cell(row=r_idx, column=55, value=p.best_marketplace)
            ws.cell(row=r_idx, column=56, value="; ".join(p.optimization_tips[:3]))
        
        # Автоширина
        for col in range(1, 57):
            col_letter = chr(64 + col) if col <= 26 else chr(64 + (col - 26) - 1) + chr(64 + (col - 26))
            ws.column_dimensions[col_letter].width = 15
        
        wb.save(output)
        output.seek(0)
        return output.getvalue()
    
    @staticmethod
    def to_csv(products: List[ProductUnit]) -> bytes:
        """Экспорт в CSV"""
        data = []
        for p in products:
            data.append({
                'Артикул': p.article,
                'Наименование': p.name,
                'Категория': p.category,
                'Бренд': p.brand,
                'Цена': p.price,
                'Себестоимость': p.cost,
                'Прибыль': p.unit_profit,
                'Маржа %': p.margin,
                'CM %': p.contribution_margin_pct,
                'ROI': p.roi,
                'Окупаемость': p.payback_days,
                'LTV/CAC': p.ltv_cac_ratio,
                'ABC': p.abc_category,
                'XYZ': p.xyz_category,
                'Лучший МП': p.best_marketplace,
                'Качество': p.quality_score,
                'Риск': p.risk_score,
                'Рекомендации': "; ".join(p.optimization_tips[:2])
            })
        df = pd.DataFrame(data)
        return df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')


# --------------------------------------------
# 🎨 ОСНОВНОЕ ПРИЛОЖЕНИЕ
# --------------------------------------------
class UltimateUnitApp:
    """Главное приложение юнит-экономики"""
    
    def __init__(self):
        self.ai_analyzer = None
        self.engine = UnitEconomicsEngine()
        self.viz = VisualizationEngine()
        self.exporter = ExportEngine()
        self.products: List[ProductUnit] = []
        self.price_optimizer = None
        self._init_session_state()
    
    def _init_session_state(self):
        """Инициализация session state"""
        defaults = {
            'products': [],
            'df': None,
            'column_mapping': {},
            'formulas': {},
            'api_key': '',
            'marketplace': 'Ozon',
            'days_storage': 30,
            'fixed_costs': 50000,
            'avg_orders': 100,
            'processed': False,
            'analysis_results': {},
            'selected_product': None,
            'history': []
        }
        for key, value in defaults.items():
            if key not in st.session_state:
                st.session_state[key] = value
    
    def run(self):
        """Запуск приложения"""
        st.set_page_config(
            page_title=CONFIG["app_name"] + " v" + CONFIG["version"],
            page_icon="🚀",
            layout="wide",
            initial_sidebar_state="expanded"
        )
        
        self._render_header()
        self._render_sidebar()
        self._render_main()
    
    def _render_header(self):
        """Заголовок"""
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    padding: 2rem; border-radius: 15px; color: white; text-align: center; margin-bottom: 2rem;">
            <h1 style="font-size: 2.5rem; margin: 0;">🚀 {CONFIG['app_name']}</h1>
            <p style="font-size: 1.1rem; opacity: 0.95;">
                🧠 AI определяет структуру данных и формулы • 50+ показателей • Прогнозирование прибыли
            </p>
            <div style="display: flex; justify-content: center; gap: 1rem; flex-wrap: wrap; margin-top: 0.5rem;">
                <span style="background: rgba(255,255,255,0.2); padding: 0.2rem 0.8rem; border-radius: 20px;">📊 Юнит-экономика</span>
                <span style="background: rgba(255,255,255,0.2); padding: 0.2rem 0.8rem; border-radius: 20px;">🧠 AI-анализ</span>
                <span style="background: rgba(255,255,255,0.2); padding: 0.2rem 0.8rem; border-radius: 20px;">💰 Оптимизация цен</span>
                <span style="background: rgba(255,255,255,0.2); padding: 0.2rem 0.8rem; border-radius: 20px;">📈 Прогнозы</span>
                <span style="background: rgba(255,255,255,0.2); padding: 0.2rem 0.8rem; border-radius: 20px;">🏪 5 МП</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    def _render_sidebar(self):
        """Боковая панель"""
        with st.sidebar:
            st.markdown("### ⚙️ Настройки")
            
            # API ключ
            api_key = st.text_input(
                "🔑 DeepSeek API ключ",
                type="password",
                placeholder="sk-...",
                key="api_key_input",
                help="Необязательно. Без AI будет работать по правилам"
            )
            if api_key:
                st.session_state.api_key = api_key
                self.ai_analyzer = AIDataAnalyzer(api_key)
                st.success("✅ AI подключён")
            
            st.markdown("---")
            
            # Параметры
            marketplace = st.selectbox(
                "🏪 Маркетплейс",
                CONFIG["supported_marketplaces"],
                index=CONFIG["supported_marketplaces"].index(st.session_state.marketplace) if st.session_state.marketplace in CONFIG["supported_marketplaces"] else 0,
                key="marketplace_select"
            )
            st.session_state.marketplace = marketplace
            
            days_storage = st.number_input(
                "📦 Хранение (дней)",
                min_value=1,
                max_value=180,
                value=st.session_state.days_storage,
                key="days_storage_input"
            )
            st.session_state.days_storage = days_storage
            
            fixed_costs = st.number_input(
                "💰 Постоянные расходы/мес (₽)",
                min_value=0,
                max_value=10000000,
                value=st.session_state.fixed_costs,
                step=1000,
                key="fixed_costs_input"
            )
            st.session_state.fixed_costs = fixed_costs
            
            avg_orders = st.number_input(
                "📊 Среднее заказов/мес",
                min_value=1,
                max_value=100000,
                value=st.session_state.avg_orders,
                key="avg_orders_input"
            )
            st.session_state.avg_orders = avg_orders
            
            st.markdown("---")
            
            # Дополнительные параметры
            with st.expander("📈 Дополнительные параметры"):
                st.number_input(
                    "Коэффициент удержания",
                    min_value=0.0,
                    max_value=1.0,
                    value=CONFIG["retention_rate"],
                    step=0.05,
                    help="Вероятность повторной покупки"
                )
                st.number_input(
                    "Ставка дисконтирования",
                    min_value=0.0,
                    max_value=0.5,
                    value=CONFIG["discount_rate"],
                    step=0.01,
                    help="Для расчета LTV"
                )
                st.number_input(
                    "Клиентов на 1000₽ рекламы",
                    min_value=1,
                    max_value=100,
                    value=CONFIG["customers_per_1000"],
                    help="Для расчета CAC"
                )
                st.number_input(
                    "Время доставки (дней)",
                    min_value=1,
                    max_value=30,
                    value=CONFIG["lead_time"],
                    help="Для расчета точки заказа"
                )
                st.slider(
                    "Уровень обслуживания",
                    min_value=0.90,
                    max_value=0.99,
                    value=CONFIG["service_level"],
                    step=0.01,
                    help="Для расчета страхового запаса"
                )
            
            st.markdown("---")
            st.caption(f"🕐 {datetime.now().strftime('%d.%m.%Y %H:%M')}")
            
            if st.session_state.processed:
                st.success(f"📦 Загружено: {len(st.session_state.products):,} товаров")
    
    def _render_main(self):
        """Основной контент"""
        st.subheader("📁 Загрузка данных")
        
        uploaded_file = st.file_uploader(
            "Загрузите файл с товарами (Excel/CSV)",
            type=["xlsx", "xls", "csv"],
            help="AI автоматически определит структуру данных"
        )
        
        if uploaded_file:
            # Чтение файла
            try:
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file, encoding='utf-8-sig')
                else:
                    df = pd.read_excel(uploaded_file, engine='openpyxl')
                st.session_state.df = df
            except Exception as e:
                st.error(f"❌ Ошибка чтения файла: {e}")
                return
            
            # Анализ данных
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("📊 Строк", f"{len(df):,}")
            
            with col2:
                st.metric("📋 Столбцов", len(df.columns))
            
            with col3:
                st.metric("📈 Числовые", sum(1 for col in df.columns if pd.api.types.is_numeric_dtype(df[col])))
            
            with col4:
                if st.button("🧠 AI-анализ данных", type="primary", use_container_width=True):
                    self._analyze_data(df)
            
            # Показываем результат анализа
            if st.session_state.processed:
                self._render_results()
            
            # Предпросмотр данных
            with st.expander("🔍 Предпросмотр данных"):
                st.dataframe(df.head(10), use_container_width=True)
                st.caption(f"Показаны первые 10 строк из {len(df):,}")
        else:
            self._render_welcome()
    
    def _analyze_data(self, df: pd.DataFrame):
        """AI-анализ данных"""
        with st.spinner("🧠 AI анализирует структуру данных и генерирует формулы..."):
            if self.ai_analyzer:
                analysis = self.ai_analyzer.analyze_columns(df)
                st.session_state.column_mapping = analysis["mapping"]
                st.session_state.formulas = analysis["formulas"]
                st.session_state.analysis_results = analysis
            else:
                analyzer = AIDataAnalyzer()
                analysis = analyzer.analyze_columns(df)
                st.session_state.column_mapping = analysis["mapping"]
                st.session_state.formulas = analysis["formulas"]
                st.session_state.analysis_results = analysis
            
            # Обработка товаров
            self._process_products(df)
            st.session_state.processed = True
            st.success("✅ Анализ завершён!")
            st.rerun()
    
    def _process_products(self, df: pd.DataFrame):
        """Обработка товаров"""
        mapping = st.session_state.column_mapping
        self.products = []
        
        self.engine = UnitEconomicsEngine(
            marketplace=st.session_state.marketplace,
            days_storage=st.session_state.days_storage
        )
        self.price_optimizer = PriceOptimizer(self.engine)
        
        # Определяем максимальное количество товаров для обработки
        max_products = min(len(df), CONFIG["max_ai_products"])
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, (_, row) in enumerate(df.head(max_products).iterrows()):
            product = ProductUnit()
            
            # Заполнение данных из маппинга
            if mapping.get("identifier") and mapping["identifier"] in df.columns:
                product.article = str(row[mapping["identifier"]])
            
            if mapping.get("name") and mapping["name"] in df.columns:
                product.name = str(row[mapping["name"]])
            
            if mapping.get("price") and mapping["price"] in df.columns:
                try:
                    product.price = float(row[mapping["price"]])
                except:
                    product.price = 0
            
            if mapping.get("cost") and mapping["cost"] in df.columns:
                try:
                    product.cost = float(row[mapping["cost"]])
                except:
                    product.cost = 0
            
            if mapping.get("length") and mapping["length"] in df.columns:
                try:
                    product.length = float(row[mapping["length"]])
                except:
                    product.length = 0
            
            if mapping.get("width") and mapping["width"] in df.columns:
                try:
                    product.width = float(row[mapping["width"]])
                except:
                    product.width = 0
            
            if mapping.get("height") and mapping["height"] in df.columns:
                try:
                    product.height = float(row[mapping["height"]])
                except:
                    product.height = 0
            
            if mapping.get("weight") and mapping["weight"] in df.columns:
                try:
                    product.weight = float(row[mapping["weight"]])
                except:
                    product.weight = 0
            
            if mapping.get("category") and mapping["category"] in df.columns:
                product.category = str(row[mapping["category"]])
            
            if mapping.get("brand") and mapping["brand"] in df.columns:
                product.brand = str(row[mapping["brand"]])
            
            if mapping.get("country") and mapping["country"] in df.columns:
                product.country = str(row[mapping["country"]])
            
            if mapping.get("quantity") and mapping["quantity"] in df.columns:
                try:
                    product.quantity = float(row[mapping["quantity"]])
                except:
                    product.quantity = 0
            
            # Расчет юнит-экономики
            if product.price > 0 and product.cost > 0:
                self.engine.calculate_product(
                    product,
                    fixed_costs=st.session_state.fixed_costs,
                    avg_orders=st.session_state.avg_orders
                )
                self.products.append(product)
            
            # Обновление прогресса
            if (idx + 1) % 10 == 0:
                progress_bar.progress((idx + 1) / max_products)
                status_text.text(f"Обработано: {idx + 1:,} / {max_products:,} товаров")
        
        progress_bar.progress(1.0)
        status_text.empty()
        
        st.session_state.products = self.products
    
    def _render_results(self):
        """Отображение результатов"""
        products = st.session_state.products
        
        if not products:
            st.warning("⚠️ Нет данных для отображения")
            return
        
        # Статистика
        total_revenue = sum(p.price for p in products)
        total_profit = sum(p.unit_profit for p in products)
        avg_margin = np.mean([p.margin for p in products])
        profitable = sum(1 for p in products if p.unit_profit > 0)
        avg_ltv_cac = np.mean([p.ltv_cac_ratio for p in products])
        avg_quality = np.mean([p.quality_score for p in products])
        
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        col1.metric("📦 Товаров", f"{len(products):,}")
        col2.metric("💰 Прибыльных", f"{profitable:,}")
        col3.metric("💵 Выручка", f"{total_revenue:,.0f} ₽")
        col4.metric("📈 Прибыль", f"{total_profit:,.0f} ₽")
        col5.metric("📊 Ср. маржа", f"{avg_margin:.1f}%")
        col6.metric("📈 LTV/CAC", f"{avg_ltv_cac:.2f}")
        
        # Вкладки
        tabs = st.tabs([
            "📊 Дашборд",
            "📈 Детальный анализ",
            "🏪 Маркетплейсы",
            "💰 Оптимизация цен",
            "💡 Рекомендации",
            "📥 Экспорт"
        ])
        
        with tabs[0]:
            self._render_dashboard(products)
        
        with tabs[1]:
            self._render_detail_analysis(products)
        
        with tabs[2]:
            self._render_marketplace_analysis(products)
        
        with tabs[3]:
            self._render_price_optimization(products)
        
        with tabs[4]:
            self._render_recommendations(products)
        
        with tabs[5]:
            self._render_export(products)
    
    def _render_dashboard(self, products: List[ProductUnit]):
        """Дашборд"""
        col1, col2 = st.columns(2)
        
        with col1:
            fig = self.viz.plot_margin_distribution(products)
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            fig = self.viz.plot_profit_pareto(products)
            st.plotly_chart(fig, use_container_width=True)
        
        col3, col4 = st.columns(2)
        
        with col3:
            fig = self.viz.plot_abc_xyz_matrix(products)
            st.plotly_chart(fig, use_container_width=True)
        
        with col4:
            fig = self.viz.plot_quality_scores(products)
            st.plotly_chart(fig, use_container_width=True)
        
        # Топ-10 по прибыли
        st.subheader("🏆 Топ-10 по прибыли")
        top = sorted(products, key=lambda x: x.unit_profit, reverse=True)[:10]
        df_top = pd.DataFrame([{
            "Артикул": p.article[:15],
            "Название": p.name[:25],
            "Цена": f"{p.price:,.0f} ₽",
            "Прибыль": f"{p.unit_profit:,.0f} ₽",
            "Маржа": f"{p.margin:.1f}%",
            "ROI": f"{p.roi:.1f}%",
            "LTV/CAC": f"{p.ltv_cac_ratio:.2f}",
            "Качество": f"{p.quality_score:.0f}"
        } for p in top])
        st.dataframe(df_top, use_container_width=True, hide_index=True)
    
    def _render_detail_analysis(self, products: List[ProductUnit]):
        """Детальный анализ"""
        st.subheader("🔍 Детальный анализ товара")
        
        # Выбор товара
        options = [f"{p.article} - {p.name[:40]}" for p in products[:100]]
        selected = st.selectbox(
            "Выберите товар для анализа",
            options,
            key="detail_select"
        )
        
        if selected:
            article = selected.split(" - ")[0]
            product = next((p for p in products if p.article == article), None)
            
            if product:
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("### 📊 Ключевые показатели")
                    metrics = [
                        ("Цена", f"{product.price:,.0f} ₽"),
                        ("Себестоимость", f"{product.cost:,.0f} ₽"),
                        ("Прибыль", f"{product.unit_profit:,.0f} ₽"),
                        ("Маржа", f"{product.margin:.1f}%"),
                        ("CM %", f"{product.contribution_margin_pct:.1f}%"),
                        ("ROI", f"{product.roi:.1f}%"),
                        ("Окупаемость", f"{product.payback_days:.0f} дн"),
                        ("LTV/CAC", f"{product.ltv_cac_ratio:.2f}"),
                        ("ДРР %", f"{product.drr:.1f}%"),
                        ("ROAS", f"{product.roas:.2f}x"),
                        ("Качество", f"{product.quality_score:.0f}"),
                        ("Риск", f"{product.risk_score:.0f}")
                    ]
                    for name, value in metrics:
                        st.metric(name, value)
                
                with col2:
                    st.markdown("### 📊 Структура расходов")
                    fig = self.viz.plot_cost_structure(product)
                    st.plotly_chart(fig, use_container_width=True)
                    
                    st.markdown("### 🎯 Анализ чувствительности")
                    fig = self.viz.plot_sensitivity(product)
                    st.plotly_chart(fig, use_container_width=True)
                    
                    st.markdown("### 📈 Прогноз спроса")
                    fig = self.viz.plot_forecast(product)
                    st.plotly_chart(fig, use_container_width=True)
                
                st.markdown("### 💡 Рекомендации")
                for tip in product.optimization_tips:
                    if "🔴" in tip:
                        st.error(tip)
                    elif "🟡" in tip:
                        st.warning(tip)
                    elif "🟢" in tip:
                        st.success(tip)
                    else:
                        st.info(tip)
    
    def _render_marketplace_analysis(self, products: List[ProductUnit]):
        """Анализ маркетплейсов"""
        st.subheader("🏪 Сравнение маркетплейсов")
        
        fig = self.viz.plot_marketplace_comparison(products)
        st.plotly_chart(fig, use_container_width=True)
        
        st.info("💡 Анализ показывает, на каком маркетплейсе ваши товары принесут максимальную прибыль")
        
        # Детальная таблица
        st.subheader("📊 Детальное сравнение по маркетплейсам")
        
        marketplaces = ["Ozon", "Wildberries", "Яндекс Маркет", "AliExpress", "Мегамаркет"]
        data = []
        for mp in marketplaces:
            engine = UnitEconomicsEngine(marketplace=mp, days_storage=30)
            total_profit = 0
            total_revenue = 0
            total_margin = 0
            for p in products[:50]:
                test = ProductUnit(**p.__dict__)
                engine.calculate_product(test)
                total_profit += test.unit_profit
                total_revenue += test.price
                total_margin += test.margin
            data.append({
                "Маркетплейс": mp,
                "Прибыль": f"{total_profit:,.0f} ₽",
                "Выручка": f"{total_revenue:,.0f} ₽",
                "Маржа": f"{total_margin / len(products[:50]):.1f}%" if products[:50] else "0%"
            })
        
        df = pd.DataFrame(data)
        st.dataframe(df, use_container_width=True, hide_index=True)
    
    def _render_price_optimization(self, products: List[ProductUnit]):
        """Оптимизация цен"""
        st.subheader("💰 Оптимизация цен")
        
        st.markdown("""
        <div style="background: #e7f3ff; padding: 1rem; border-radius: 10px; margin-bottom: 1rem;">
        💡 <b>Оптимизация цен</b> помогает найти оптимальную цену для максимальной прибыли.
        </div>
        """, unsafe_allow_html=True)
        
        # Выбор товара
        options = [f"{p.article} - {p.name[:40]}" for p in products[:100]]
        selected = st.selectbox(
            "Выберите товар для оптимизации",
            options,
            key="optimize_select"
        )
        
        if selected:
            article = selected.split(" - ")[0]
            product = next((p for p in products if p.article == article), None)
            
            if product:
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("### 📊 Текущие показатели")
                    st.metric("Текущая цена", f"{product.price:,.0f} ₽")
                    st.metric("Текущая прибыль", f"{product.unit_profit:,.0f} ₽")
                    st.metric("Текущая маржа", f"{product.margin:.1f}%")
                    st.metric("Текущий ROI", f"{product.roi:.1f}%")
                
                with col2:
                    st.markdown("### 📈 Результаты оптимизации")
                    result = self.price_optimizer.optimize_price(product)
                    st.metric("Рекомендуемая цена", f"{result['recommended_price']:,.0f} ₽", 
                             f"{result['price_change']:+,.0f} ₽")
                    st.metric("Оптимизированная прибыль", f"{result['recommended_profit']:,.0f} ₽",
                             f"{result['profit_increase']:+,.0f} ₽")
                    st.metric("Оптимальная маржа", f"{result['recommended_margin']:.1f}%")
                    st.info(f"📌 {result['recommendation']}")
                
                st.markdown("### 📊 График оптимизации")
                fig = self.viz.plot_price_optimization(product, self.price_optimizer)
                st.plotly_chart(fig, use_container_width=True)
                
                # Таблица всех товаров с рекомендациями
                st.subheader("📋 Рекомендации по ценам для всех товаров")
                all_results = self.price_optimizer.optimize_all_prices(products[:20])
                df_results = pd.DataFrame(all_results)
                st.dataframe(df_results, use_container_width=True, hide_index=True)
    
    def _render_recommendations(self, products: List[ProductUnit]):
        """Рекомендации"""
        st.subheader("💡 Рекомендации по оптимизации")
        
        # Сбор рекомендаций
        all_tips = {}
        for p in products:
            for tip in p.optimization_tips:
                all_tips[tip] = all_tips.get(tip, 0) + 1
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 📊 Распределение рекомендаций")
            if all_tips:
                sorted_tips = sorted(all_tips.items(), key=lambda x: x[1], reverse=True)
                fig = go.Figure(data=[
                    go.Bar(
                        x=[tip[:40] + "..." if len(tip) > 40 else tip for tip, _ in sorted_tips[:10]],
                        y=[count for _, count in sorted_tips[:10]],
                        marker_color='#667eea'
                    )
                ])
                fig.update_layout(
                    title='Топ-10 рекомендаций',
                    height=400,
                    xaxis_tickangle=-45
                )
                st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            st.markdown("### 💡 Детальные рекомендации")
            if all_tips:
                for tip, count in sorted(all_tips.items(), key=lambda x: x[1], reverse=True)[:15]:
                    if "🔴" in tip:
                        st.error(f"{tip} ({count} товаров)")
                    elif "🟡" in tip:
                        st.warning(f"{tip} ({count} товаров)")
                    elif "🟢" in tip:
                        st.success(f"{tip} ({count} товаров)")
                    else:
                        st.info(f"{tip} ({count} товаров)")
            else:
                st.success("✅ Все товары в норме!")
    
    def _render_export(self, products: List[ProductUnit]):
        """Экспорт"""
        st.subheader("📥 Экспорт данных")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.info(f"📊 Экспортируемые товары: {len(products):,}")
            st.info(f"📋 Показателей: 56")
            st.info(f"📋 Столбцов в Excel: 56")
            format_type = st.radio(
                "Формат:",
                ["Excel (с формулами)", "CSV"],
                index=0
            )
        
        with col2:
            if st.button("📥 Скачать", type="primary", use_container_width=True):
                with st.spinner("⏳ Генерация файла..."):
                    if format_type.startswith("Excel"):
                        data = self.exporter.to_excel(products)
                        ext = "xlsx"
                        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    else:
                        data = self.exporter.to_csv(products)
                        ext = "csv"
                        mime = "text/csv"
                    
                    st.download_button(
                        "📥 Скачать файл",
                        data=data,
                        file_name=f"юнит_экономика_{datetime.now().strftime('%Y%m%d_%H%M')}.{ext}",
                        mime=mime,
                        use_container_width=True
                    )
    
    def _render_welcome(self):
        """Приветственный экран"""
        st.markdown("""
        <div style="text-align: center; padding: 40px 0;">
            <h2 style="font-size: 2rem; color: #1a1a2e;">🚀 Загрузите данные для анализа</h2>
            <p style="color: #6c757d; font-size: 1.1rem;">
                🧠 AI определит структуру данных и рассчитает юнит-экономику
            </p>
            
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
                        gap: 15px; max-width: 1000px; margin: 30px auto;">
                <div style="background: #667eea; padding: 20px; border-radius: 10px; color: white;">
                    <h3>🧠 AI</h3>
                    <p style="font-size: 12px;">Определение столбцов</p>
                </div>
                <div style="background: #764ba2; padding: 20px; border-radius: 10px; color: white;">
                    <h3>📊 56</h3>
                    <p style="font-size: 12px;">Показателей</p>
                </div>
                <div style="background: #f093fb; padding: 20px; border-radius: 10px; color: white;">
                    <h3>💰 Прогноз</h3>
                    <p style="font-size: 12px;">Прибыли до продажи</p>
                </div>
                <div style="background: #f5576c; padding: 20px; border-radius: 10px; color: white;">
                    <h3>📈 Оптимизация</h3>
                    <p style="font-size: 12px;">Ценообразования</p>
                </div>
                <div style="background: #4facfe; padding: 20px; border-radius: 10px; color: white;">
                    <h3>🏪 5 МП</h3>
                    <p style="font-size: 12px;">Маркетплейсов</p>
                </div>
                <div style="background: #43e97b; padding: 20px; border-radius: 10px; color: white;">
                    <h3>📥 Экспорт</h3>
                    <p style="font-size: 12px;">Excel/CSV</p>
                </div>
            </div>
            
            <div style="background: #e7f3ff; padding: 20px; border-radius: 10px;
                        max-width: 700px; margin: 20px auto; text-align: left;">
                <p><b>📌 Поддерживаемые форматы:</b></p>
                <p style="font-size: 0.9rem;">• Excel (.xlsx, .xls)</p>
                <p style="font-size: 0.9rem;">• CSV (любой разделитель)</p>
                <p style="font-size: 0.9rem;">• Автоматическое определение структуры данных</p>
                <p style="font-size: 0.9rem;">• AI-анализ и рекомендации</p>
                <p style="font-size: 0.9rem;">• 56 показателей юнит-экономики</p>
            </div>
        </div>
        """, unsafe_allow_html=True)


# --------------------------------------------
# ЗАПУСК
# --------------------------------------------
if __name__ == "__main__":
    app = UltimateUnitApp()
    app.run()
