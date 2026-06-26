"""
================================================================================
🚗 ПРОФЕССИОНАЛЬНАЯ ПЛАТФОРМА АВТОЗАПЧАСТЕЙ - ULTIMATE EDITION v14.0
================================================================================
💎 УПОР НА ЮНИТ-ЭКОНОМИКУ: Contribution Margin • LTV/CAC • Break-Even • Чувствительность

📌 ОСНОВНЫЕ ВОЗМОЖНОСТИ:
    ✅ Детальная юнит-экономика по каждому товару
    ✅ Анализ чувствительности (что-if сценарии)
    ✅ Точка безубыточности и break-even
    ✅ Contribution Margin, ROI, ROAS, ДРР
    ✅ Структура расходов (waterfall + pie)
    ✅ Парсер ТОЛЬКО ваших товаров (быстро и просто)
    ✅ Все API ключи через интерфейс
    ✅ Весь интерфейс на русском

🚀 УСТАНОВКА:
    pip install streamlit pandas numpy openpyxl plotly requests selenium webdriver-manager

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
import hashlib
import os
import random
import smtplib
import requests
import logging
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple, Any, Union, Callable
from dataclasses import dataclass, asdict, field
from pathlib import Path
from io import BytesIO
import zipfile
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from concurrent.futures import ThreadPoolExecutor
from threading import Lock
import gc
import warnings
warnings.filterwarnings('ignore')

# SELENIUM ИМПОРТЫ
try:
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException, WebDriverException
    from webdriver_manager.chrome import ChromeDriverManager
    SELENIUM_AVAILABLE = True
except ImportError:
    SELENIUM_AVAILABLE = False

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
    "version": "14.0.0",
    "app_name": "🚗 Платформа автозапчастей",
    "supported_marketplaces": ["Ozon", "Wildberries", "Яндекс Маркет", "AliExpress", "Мегамаркет"],
    "batch_size": 1000,
    "max_workers": 4,
    "currency": "₽",
    "chunk_size": 50000,
    "separators": [";", ",", "|", "/", "\\"],
    "max_ai_products": 100,
    "tariffs_file": "tariffs_cache.json",
    "selenium_timeout": 10,
    "selenium_max_products": 10,
}

# --------------------------------------------
# 🕷️ SELENIUM ПАРСЕР (УПРОЩЁННЫЙ - ТОЛЬКО СВОИ ТОВАРЫ)
# --------------------------------------------
class MarketplaceScraper:
    """
    Парсер для поиска ТОЛЬКО ваших товаров на маркетплейсах
    
    📌 Как это работает:
        1. Поиск по артикулу или OE номеру
        2. Проверка наличия товара на странице
        3. Возврат статуса найден/не найден
    
    ⚠️ Внимание: Парсинг может занимать время, используйте для выборочной проверки
    """
    
    def __init__(self, headless: bool = True):
        """
        Инициализация парсера
        
        Args:
            headless: Запуск в фоновом режиме (без открытия браузера)
        """
        self.headless = headless
        self.driver = None
        self.timeout = CONFIG['selenium_timeout']
        self.max_products = CONFIG['selenium_max_products']

        if not SELENIUM_AVAILABLE:
            logger.warning("⚠️ Selenium не установлен. Используйте: pip install selenium webdriver-manager")

    def _init_driver(self) -> bool:
        """Инициализация Selenium WebDriver"""
        if not SELENIUM_AVAILABLE:
            return False

        try:
            options = Options()

            if self.headless:
                options.add_argument('--headless')

            options.add_argument('--disable-blink-features=AutomationControlled')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-gpu')
            options.add_argument('--window-size=1920,1080')
            options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36')
            options.add_experimental_option('excludeSwitches', ['enable-automation'])
            options.add_experimental_option('useAutomationExtension', False)

            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)

            self.driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
                'source': '''
                    delete navigator.__proto__.webdriver;
                    Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                '''
            })

            logger.info("✅ Selenium драйвер инициализирован")
            return True

        except Exception as e:
            logger.error(f"❌ Ошибка инициализации Selenium: {e}")
            return False

    def _close_driver(self):
        """Закрытие драйвера"""
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None

    def search_my_product(self, article: str, oe_number: str, marketplace: str,
                         progress_callback: Optional[Callable] = None) -> Dict:
        """
        Поиск товара на маркетплейсе
        
        Args:
            article: Артикул товара
            oe_number: OE номер (если есть)
            marketplace: Название маркетплейса
            progress_callback: Функция для отображения прогресса
        
        Returns:
            Результат поиска с информацией о найденных товарах
        """
        if not self._init_driver():
            return self._fallback_simulation(article, marketplace)

        query = oe_number if oe_number else article
        query_quoted = f'"{query}"'

        try:
            if progress_callback:
                progress_callback(0.2, f"🌐 Открываю {marketplace}...")

            if marketplace == "Ozon":
                url = f"https://www.ozon.ru/search/?text={requests.utils.quote(query_quoted)}"
            elif marketplace == "Wildberries":
                url = f"https://www.wildberries.ru/catalog/0/search.aspx?search={requests.utils.quote(query_quoted)}"
            elif marketplace == "Яндекс Маркет":
                url = f"https://market.yandex.ru/search?text={requests.utils.quote(query_quoted)}"
            else:
                return self._fallback_simulation(article, marketplace)

            self.driver.get(url)
            time.sleep(2)

            if progress_callback:
                progress_callback(0.5, f"🔍 Ищу {query}...")

            try:
                WebDriverWait(self.driver, self.timeout).until(
                    EC.presence_of_element_located((By.TAG_NAME, "body"))
                )
            except TimeoutException:
                return self._fallback_simulation(article, marketplace)

            result = {
                "marketplace": marketplace,
                "query": query,
                "found": False,
                "my_products": [],
                "source": "selenium"
            }

            page_source = self.driver.page_source.lower()

            if article.lower() in page_source or query.lower() in page_source:
                result["found"] = True
                result["my_products"].append({
                    "article": article,
                    "oe": oe_number,
                    "status": "✅ Найден на маркетплейсе"
                })

            self._close_driver()

            if progress_callback:
                progress_callback(1.0, "✅ Поиск завершён")

            return result

        except Exception as e:
            logger.error(f"❌ Ошибка парсинга: {e}")
            self._close_driver()
            return self._fallback_simulation(article, marketplace)

    def _fallback_simulation(self, article: str, marketplace: str) -> Dict:
        """Симуляция поиска при недоступности Selenium"""
        logger.info(f"🔄 Используем симуляцию для {article}")

        return {
            "marketplace": marketplace,
            "query": article,
            "found": random.choice([True, False]),
            "my_products": [{
                "article": article,
                "oe": "",
                "status": "🔄 Симуляция (Selenium не работает)"
            }] if random.choice([True, False]) else [],
            "source": "simulation"
        }


# --------------------------------------------
# 🎯 РАСШИРЕННАЯ СТРУКТУРА ДАННЫХ (С ЮНИТ-ЭКОНОМИКОЙ)
# --------------------------------------------
@dataclass
class ProductBatch:
    """
    Полная структура товара с детальной юнит-экономикой
    
    📌 Содержит все показатели для комплексного анализа:
        - Базовые: цена, себестоимость, габариты
        - Юнит-экономика: CM, LTV, CAC, Break-Even
        - Аналитические: ABC/XYZ, прогнозы, запасы
        - Маркетплейсы: комиссии, логистика, рейтинг
    """
    # Основные данные
    article: str
    name: str
    compatibility: str = ""
    category: str = ""
    country: str = ""
    oe_numbers: List[str] = field(default_factory=list)
    analogs: List[str] = field(default_factory=list)
    price: float = 0
    cost: float = 0
    length: float = 0
    width: float = 0
    height: float = 0
    weight: float = 0

    # Базовые расчётные показатели
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
    profit: float = 0
    margin: float = 0
    roi: float = 0
    risk_score: float = 0

    # ЮНИТ-ЭКОНОМИКА (детальная)
    contribution_margin: float = 0
    contribution_margin_pct: float = 0
    break_even_units: int = 0
    break_even_revenue: float = 0
    payback_period_days: float = 0
    ltv_estimate: float = 0
    cac_estimate: float = 0
    ltv_cac_ratio: float = 0
    roas: float = 0
    drr: float = 0
    unit_profit: float = 0
    unit_cost_structure: Dict = field(default_factory=dict)

    # AI
    ai_recommended_price: float = 0
    ai_recommended_markup: float = 0
    ai_optimization_tips: List[str] = field(default_factory=list)

    # Стратегии
    sales_strategy: str = ""
    price_position: str = ""
    recommended_action: str = ""

    # ABC/XYZ
    abc_category: str = ""
    xyz_category: str = ""
    demand_forecast_30: float = 0
    demand_forecast_60: float = 0
    demand_forecast_90: float = 0
    demand_trend: str = ""
    eoq: float = 0
    reorder_point: float = 0
    safety_stock: float = 0
    stock_status: str = ""
    return_probability: float = 0
    best_marketplace: str = ""
    competitor_avg_price: float = 0
    competitor_min_price: float = 0
    price_position_vs_competitors: str = ""
    logistics_provider: str = ""
    logistics_cost_optimal: float = 0
    ab_test_group: str = ""
    ab_test_revenue: float = 0
    daily_sales_history: List[float] = field(default_factory=list)

    # Selenium
    selenium_found: bool = False
    selenium_marketplace: str = ""
    selenium_status: str = ""

    def to_dict(self) -> Dict:
        """Преобразование в словарь для сериализации"""
        return asdict(self)


# --------------------------------------------
# 💎 АНАЛИЗАТОР ЮНИТ-ЭКОНОМИКИ
# --------------------------------------------
class UnitEconomicsAnalyzer:
    """
    Детальный анализ юнит-экономики по каждому товару
    """
    
    def analyze_product(self, product: ProductBatch, fixed_costs_per_month: float = 0,
                       avg_orders_per_month: int = 100, monthly_demand: int = None) -> Dict:
        """
        Полный анализ юнит-экономики для одного товара
        """
        if monthly_demand is None:
            monthly_demand = avg_orders_per_month

        revenue_per_unit = product.price
        variable_cost_per_unit = (product.cost + product.commission + product.logistics +
                                  product.storage + product.acquiring + product.advertising +
                                  product.returns + product.packaging + product.customs +
                                  product.other_costs)

        product.total_variable_costs = variable_cost_per_unit
        product.unit_profit = revenue_per_unit - variable_cost_per_unit

        contribution_margin = revenue_per_unit - variable_cost_per_unit
        contribution_margin_pct = (contribution_margin / revenue_per_unit * 100) if revenue_per_unit > 0 else 0

        product.contribution_margin = contribution_margin
        product.contribution_margin_pct = contribution_margin_pct

        if contribution_margin > 0:
            break_even_units = math.ceil(fixed_costs_per_month / contribution_margin)
            break_even_revenue = break_even_units * revenue_per_unit
        else:
            break_even_units = 999999
            break_even_revenue = 0

        product.break_even_units = break_even_units
        product.break_even_revenue = break_even_revenue

        payback_period_days = self._calculate_payback(product, monthly_demand)
        product.payback_period_days = payback_period_days

        ltv_cac_data = self._calculate_ltv_cac(product)
        product.ltv_estimate = ltv_cac_data["ltv"]
        product.cac_estimate = ltv_cac_data["cac"]
        product.ltv_cac_ratio = ltv_cac_data["ltv_cac_ratio"]

        roas = revenue_per_unit / product.advertising if product.advertising > 0 else 0
        drr = (product.advertising / revenue_per_unit * 100) if revenue_per_unit > 0 else 0

        product.roas = roas
        product.drr = drr

        unit_cost_structure = {
            "Себестоимость": product.cost,
            "Комиссия МП": product.commission,
            "Логистика": product.logistics,
            "Хранение": product.storage,
            "Эквайринг": product.acquiring,
            "Реклама": product.advertising,
            "Возвраты": product.returns,
            "Упаковка": product.packaging,
            "Таможня": product.customs,
            "Прочее": product.other_costs
        }

        product.unit_cost_structure = unit_cost_structure

        sensitivity = self._calculate_sensitivity(product)

        recommendations = self._generate_recommendations(product, contribution_margin_pct,
                                                        product.ltv_cac_ratio, drr)

        return {
            "article": product.article,
            "name": product.name,
            "price": revenue_per_unit,
            "variable_costs": variable_cost_per_unit,
            "contribution_margin": contribution_margin,
            "contribution_margin_pct": contribution_margin_pct,
            "break_even_units": break_even_units,
            "break_even_revenue": break_even_revenue,
            "payback_period_days": payback_period_days,
            "ltv": product.ltv_estimate,
            "cac": product.cac_estimate,
            "ltv_cac_ratio": product.ltv_cac_ratio,
            "ltv_cac_interpretation": ltv_cac_data["interpretation"],
            "roas": roas,
            "drr": drr,
            "unit_profit": product.unit_profit,
            "unit_cost_structure": unit_cost_structure,
            "sensitivity": sensitivity,
            "recommendations": recommendations
        }

    def _calculate_payback(self, product: ProductBatch, monthly_demand: int) -> float:
        """Расчет срока окупаемости"""
        if monthly_demand <= 0:
            return 999
        
        monthly_revenue = product.price * monthly_demand
        monthly_cost = product.total_cost * monthly_demand
        monthly_profit = monthly_revenue - monthly_cost
        
        if monthly_profit <= 0:
            return 999
        
        initial_investment = product.cost * monthly_demand
        payback_months = initial_investment / monthly_profit
        return payback_months * 30

    def _calculate_ltv_cac(self, product: ProductBatch,
                          retention_rate: float = 0.7,
                          discount_rate: float = 0.1,
                          customers_per_1000_rub: int = 5) -> Dict:
        """Расчет LTV и CAC"""
        monthly_profit = product.unit_profit
        
        ltv = monthly_profit / (1 - retention_rate + discount_rate)
        
        if product.advertising > 0 and customers_per_1000_rub > 0:
            cac = product.advertising / (product.advertising / 1000 * customers_per_1000_rub)
        else:
            cac = product.price * 0.05
        
        ltv_cac_ratio = ltv / cac if cac > 0 else 0
        
        return {
            "ltv": ltv,
            "cac": cac,
            "ltv_cac_ratio": ltv_cac_ratio,
            "interpretation": self._interpret_ltv_cac(ltv_cac_ratio)
        }

    def _interpret_ltv_cac(self, ratio: float) -> str:
        """Интерпретация соотношения LTV/CAC"""
        if ratio < 1:
            return "🔴 Критично: LTV < CAC. Вы тратите больше на привлечение, чем зарабатываете!"
        elif ratio < 3:
            return "🟡 Средний показатель. Рекомендуется оптимизация рекламы"
        else:
            return "🟢 Отличный показатель. Можно масштабировать рекламный бюджет"

    def _calculate_sensitivity(self, product: ProductBatch) -> Dict:
        """Анализ чувствительности"""
        base_profit = product.price - product.total_variable_costs

        price_up_10 = (product.price * 1.10) - product.total_variable_costs
        price_down_10 = (product.price * 0.90) - product.total_variable_costs

        cost_up_10 = product.price - (product.total_variable_costs + product.cost * 0.10)
        cost_down_10 = product.price - (product.total_variable_costs - product.cost * 0.10)

        commission_up = product.price - (product.total_variable_costs + product.price * 0.02)
        commission_down = product.price - (product.total_variable_costs - product.price * 0.02)

        logistics_up = product.price - (product.total_variable_costs + product.logistics * 0.20)
        logistics_down = product.price - (product.total_variable_costs - product.logistics * 0.20)

        ads_up = product.price - (product.total_variable_costs + product.advertising * 0.50)
        ads_down = product.price - (product.total_variable_costs - product.advertising * 0.50)

        return {
            "base_profit": base_profit,
            "price_up_10": price_up_10,
            "price_down_10": price_down_10,
            "cost_up_10": cost_up_10,
            "cost_down_10": cost_down_10,
            "commission_up": commission_up,
            "commission_down": commission_down,
            "logistics_up": logistics_up,
            "logistics_down": logistics_down,
            "ads_up": ads_up,
            "ads_down": ads_down,
            "price_elasticity": ((price_up_10 - price_down_10) / 2) / base_profit if base_profit != 0 else 0,
            "cost_elasticity": ((cost_up_10 - cost_down_10) / 2) / base_profit if base_profit != 0 else 0,
            "most_sensitive": self._get_most_sensitive_factor({
                "Цена": abs(price_up_10 - price_down_10),
                "Себестоимость": abs(cost_up_10 - cost_down_10),
                "Комиссия": abs(commission_up - commission_down),
                "Логистика": abs(logistics_up - logistics_down),
                "Реклама": abs(ads_up - ads_down)
            })
        }

    def _get_most_sensitive_factor(self, factors: Dict) -> str:
        """Определяет самый чувствительный фактор"""
        if not factors:
            return "Нет данных"
        return max(factors.items(), key=lambda x: x[1])[0]

    def _generate_recommendations(self, product: ProductBatch, cm_pct: float,
                                 ltv_cac: float, drr: float) -> List[str]:
        """Генерирует рекомендации по оптимизации юнит-экономики"""
        recs = []

        if cm_pct < 20:
            recs.append("🔴 Критически низкий Contribution Margin — пересмотрите цену или себестоимость")
        elif cm_pct < 35:
            recs.append("🟡 Низкий CM — рассмотрите повышение цены на 5-10%")
        elif cm_pct > 50:
            recs.append("🟢 Отличный CM — можно увеличить рекламный бюджет")

        if ltv_cac < 1:
            recs.append("🔴 LTV/CAC < 1 — вы тратите больше на привлечение, чем зарабатываете")
        elif ltv_cac < 3:
            recs.append("🟡 LTV/CAC < 3 — оптимизируйте рекламу или повышайте повторные продажи")
        else:
            recs.append("🟢 Отличный LTV/CAC — масштабируйте рекламу")

        if drr > 25:
            recs.append("🔴 ДРР > 25% — слишком высокие рекламные расходы")
        elif drr > 15:
            recs.append("🟡 ДРР > 15% — оптимизируйте рекламные кампании")
        elif drr > 0 and drr < 10:
            recs.append("🟢 Низкий ДРР — можно увеличить рекламный бюджет")

        if product.logistics / product.price > 0.20:
            recs.append("📦 Логистика съедает >20% цены — оптимизируйте габариты или выберите другого провайдера")

        if product.returns / product.price > 0.10:
            recs.append("⚠️ Возвраты >10% — улучшите описание и фото товара")

        if product.commission / product.price > 0.15:
            recs.append("💸 Комиссия >15% — рассмотрите другой маркетплейс или категорию")

        if product.payback_period_days > 90:
            recs.append("⏳ Окупаемость >90 дней — сократите размер закупаемой партии")

        if not recs:
            recs.append("✅ Юнит-экономика в норме")

        return recs

    def analyze_portfolio(self, products: List[ProductBatch]) -> Dict:
        """Анализ юнит-экономики всего портфеля"""
        total_revenue = sum(p.price for p in products)
        total_variable_costs = sum(p.total_variable_costs for p in products)
        total_profit = sum(p.unit_profit for p in products)

        avg_cm = np.mean([p.contribution_margin_pct for p in products]) if products else 0
        avg_ltv_cac = np.mean([p.ltv_cac_ratio for p in products]) if products else 0
        avg_drr = np.mean([p.drr for p in products]) if products else 0
        avg_payback = np.mean([p.payback_period_days for p in products]) if products else 0

        cm_distribution = {
            "Критический (<20%)": sum(1 for p in products if p.contribution_margin_pct < 20),
            "Низкий (20-35%)": sum(1 for p in products if 20 <= p.contribution_margin_pct < 35),
            "Средний (35-50%)": sum(1 for p in products if 35 <= p.contribution_margin_pct < 50),
            "Высокий (>50%)": sum(1 for p in products if p.contribution_margin_pct >= 50)
        }

        top_by_cm = sorted(products, key=lambda x: x.contribution_margin_pct, reverse=True)[:10]

        return {
            "total_revenue": total_revenue,
            "total_variable_costs": total_variable_costs,
            "total_profit": total_profit,
            "avg_cm_pct": avg_cm,
            "avg_ltv_cac": avg_ltv_cac,
            "avg_drr": avg_drr,
            "avg_payback_days": avg_payback,
            "cm_distribution": cm_distribution,
            "top_by_cm": top_by_cm,
            "products_count": len(products)
        }


# --------------------------------------------
# 🕷️ МОНИТОРИНГ СВОИХ ТОВАРОВ
# --------------------------------------------
class MyProductTracker:
    """Отслеживает ТОЛЬКО ваши товары на маркетплейсах"""
    
    def __init__(self, use_selenium: bool = True):
        self.use_selenium = use_selenium and SELENIUM_AVAILABLE
        self.scraper = MarketplaceScraper(headless=True) if self.use_selenium else None

    def track_my_product(self, product: ProductBatch, marketplace: str,
                        progress_callback: Optional[Callable] = None) -> Dict:
        """Отслеживание одного товара"""
        if not self.use_selenium or not self.scraper:
            return self._simulate_tracking(product, marketplace)

        oe_number = product.oe_numbers[0] if product.oe_numbers else ""

        result = self.scraper.search_my_product(
            article=product.article,
            oe_number=oe_number,
            marketplace=marketplace,
            progress_callback=progress_callback
        )

        product.selenium_found = result.get("found", False)
        product.selenium_marketplace = marketplace
        product.selenium_status = result["my_products"][0]["status"] if result["my_products"] else "❌ Не найден"

        return result

    def _simulate_tracking(self, product: ProductBatch, marketplace: str) -> Dict:
        """Симуляция отслеживания"""
        found = random.choice([True, False])

        product.selenium_found = found
        product.selenium_marketplace = marketplace
        product.selenium_status = "✅ Найден (симуляция)" if found else "❌ Не найден (симуляция)"

        return {
            "marketplace": marketplace,
            "query": product.article,
            "found": found,
            "my_products": [{
                "article": product.article,
                "status": product.selenium_status
            }] if found else [],
            "source": "simulation"
        }


# --------------------------------------------
# 📈 ПРОГНОЗИРОВАНИЕ СПРОСА
# --------------------------------------------
class DemandForecaster:
    """Прогноз спроса на основе исторических данных"""
    
    def forecast(self, history: List[float], days: int = 90, alpha: float = 0.3) -> Dict:
        """Прогноз на указанное количество дней"""
        if not history or len(history) < 3:
            history = [random.randint(5, 20) for _ in range(30)]

        forecast_values = []
        last_value = history[-1]
        trend = self._calculate_trend(history)

        for i in range(days):
            seasonal = self._get_seasonal_factor(i)
            predicted = (last_value + trend * (i + 1)) * seasonal
            forecast_values.append(max(0, predicted))

        if trend > 0.1:
            trend_name = "📈 Растущий"
        elif trend < -0.1:
            trend_name = "📉 Падающий"
        else:
            trend_name = "➡️ Стабильный"

        std = np.std(history) if len(history) > 1 else 0
        lower = [max(0, v - 1.96 * std) for v in forecast_values]
        upper = [v + 1.96 * std for v in forecast_values]

        return {
            "forecast_30": sum(forecast_values[:30]),
            "forecast_60": sum(forecast_values[:60]),
            "forecast_90": sum(forecast_values),
            "daily_forecast": forecast_values,
            "lower_bound": lower,
            "upper_bound": upper,
            "trend": trend_name,
            "trend_value": trend,
            "avg_daily": np.mean(forecast_values),
            "history": history
        }

    def _calculate_trend(self, history: List[float]) -> float:
        """Расчет тренда"""
        n = len(history)
        if n < 2:
            return 0
        x = np.arange(n)
        slope, _ = np.polyfit(x, history, 1)
        return slope

    def _get_seasonal_factor(self, day_offset: int) -> float:
        """Сезонный фактор по дням недели"""
        week_day = day_offset % 7
        if week_day in [5, 6]:
            return 1.25
        elif week_day == 0:
            return 0.85
        return 1.0


# --------------------------------------------
# 🎯 ABC/XYZ АНАЛИЗ
# --------------------------------------------
class InventoryAnalyzer:
    """ABC/XYZ анализ ассортимента"""
    
    def abc_analysis(self, products: List[ProductBatch]) -> Dict:
        """ABC-анализ - учитывает все товары, включая убыточные"""
        sorted_products = sorted(products, key=lambda x: x.unit_profit, reverse=True)
        
        total_abs_profit = sum(abs(p.unit_profit) for p in sorted_products)
        
        if total_abs_profit == 0:
            return {"A": [], "B": [], "C": products}
        
        cumulative = 0
        categories = {"A": [], "B": [], "C": []}
        
        for p in sorted_products:
            cumulative += abs(p.unit_profit)
            percentage = cumulative / total_abs_profit
            
            if percentage <= 0.80:
                p.abc_category = "A"
                categories["A"].append(p)
            elif percentage <= 0.95:
                p.abc_category = "B"
                categories["B"].append(p)
            else:
                p.abc_category = "C"
                categories["C"].append(p)
        
        return categories

    def xyz_analysis(self, products: List[ProductBatch]) -> Dict:
        """XYZ-анализ по коэффициенту вариации"""
        categories = {"X": [], "Y": [], "Z": []}

        for p in products:
            history = p.daily_sales_history if p.daily_sales_history else [random.randint(0, 20) for _ in range(30)]

            if len(history) < 2:
                cv = 0
            else:
                mean_val = np.mean(history)
                std_val = np.std(history)
                cv = std_val / mean_val if mean_val > 0 else 999

            if cv < 0.2:
                p.xyz_category = "X"
                categories["X"].append(p)
            elif cv < 0.5:
                p.xyz_category = "Y"
                categories["Y"].append(p)
            else:
                p.xyz_category = "Z"
                categories["Z"].append(p)

        return categories

    def abc_xyz_matrix(self, products: List[ProductBatch]) -> Dict:
        """Построение матрицы ABC-XYZ"""
        self.abc_analysis(products)
        self.xyz_analysis(products)

        matrix = {}
        for abc in ["A", "B", "C"]:
            for xyz in ["X", "Y", "Z"]:
                key = f"{abc}{xyz}"
                matrix[key] = [p for p in products if p.abc_category == abc and p.xyz_category == xyz]

        return matrix


# --------------------------------------------
# 📦 УПРАВЛЕНИЕ ЗАПАСАМИ
# --------------------------------------------
class InventoryOptimizer:
    """Оптимизация управления запасами"""
    
    def calculate_eoq(self, annual_demand: float, order_cost: float = 500,
                     holding_cost_per_unit: float = 100) -> float:
        """Расчет EOQ (Economic Order Quantity)"""
        if holding_cost_per_unit <= 0:
            return 0
        return math.sqrt((2 * annual_demand * order_cost) / holding_cost_per_unit)

    def calculate_reorder_point(self, daily_sales: float, lead_time_days: int = 7,
                               safety_stock: float = 0) -> float:
        """Расчет точки заказа"""
        return (daily_sales * lead_time_days) + safety_stock

    def calculate_safety_stock(self, daily_sales_std: float, lead_time_days: int = 7,
                              service_level: float = 0.95) -> float:
        """Расчет страхового запаса"""
        z_scores = {0.90: 1.28, 0.95: 1.65, 0.97: 1.88, 0.99: 2.33}
        z = z_scores.get(service_level, 1.65)
        return z * daily_sales_std * math.sqrt(lead_time_days)

    def determine_stock_status(self, current_stock: float, reorder_point: float,
                              max_stock: float) -> str:
        """Определение статуса запасов"""
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

    def optimize_product(self, product: ProductBatch, current_stock: float = 100) -> Dict:
        """Комплексная оптимизация товара"""
        history = product.daily_sales_history if product.daily_sales_history else [random.randint(5, 15) for _ in range(30)]
        avg_daily = np.mean(history)
        std_daily = np.std(history) if len(history) > 1 else 0

        annual_demand = avg_daily * 365
        holding_cost = product.cost * 0.2

        eoq = self.calculate_eoq(annual_demand, order_cost=500, holding_cost_per_unit=holding_cost)
        safety_stock = self.calculate_safety_stock(std_daily, lead_time_days=7)
        reorder_point = self.calculate_reorder_point(avg_daily, lead_time_days=7, safety_stock=safety_stock)
        max_stock = eoq + safety_stock

        stock_status = self.determine_stock_status(current_stock, reorder_point, max_stock)

        product.eoq = eoq
        product.reorder_point = reorder_point
        product.safety_stock = safety_stock
        product.stock_status = stock_status

        return {
            "article": product.article,
            "avg_daily_sales": avg_daily,
            "annual_demand": annual_demand,
            "eoq": eoq,
            "safety_stock": safety_stock,
            "reorder_point": reorder_point,
            "current_stock": current_stock,
            "stock_status": stock_status,
            "days_until_stockout": current_stock / avg_daily if avg_daily > 0 else 999
        }


# --------------------------------------------
# 🏪 МУЛЬТИ-МАРКЕТПЛЕЙС АНАЛИЗ
# --------------------------------------------
class MultiMarketplaceAnalyzer:
    """Анализ эффективности продаж на разных маркетплейсах"""
    
    def __init__(self):
        self.marketplace_params = {
            "Ozon": {"commission": 0.10, "logistics": 70, "audience": 40, "return_rate": 0.12},
            "Wildberries": {"commission": 0.12, "logistics": 60, "audience": 55, "return_rate": 0.15},
            "Яндекс Маркет": {"commission": 0.11, "logistics": 80, "audience": 25, "return_rate": 0.10},
            "AliExpress": {"commission": 0.08, "logistics": 150, "audience": 15, "return_rate": 0.08},
            "Мегамаркет": {"commission": 0.09, "logistics": 75, "audience": 20, "return_rate": 0.11}
        }

    def analyze_product(self, product: ProductBatch) -> Dict:
        """Анализ товара на всех маркетплейсах"""
        results = {}

        for mp, params in self.marketplace_params.items():
            commission = product.price * params["commission"]
            logistics = params["logistics"] + product.volume * 15 + product.weight * 20
            returns = product.price * params["return_rate"]
            total_costs = product.cost + commission + logistics + returns

            profit = product.price - total_costs
            margin = (profit / product.price * 100) if product.price > 0 else 0

            score = profit * (params["audience"] / 100)

            results[mp] = {
                "profit": profit,
                "margin": margin,
                "commission": commission,
                "logistics": logistics,
                "returns": returns,
                "audience": params["audience"],
                "score": score
            }

        best_mp = max(results.items(), key=lambda x: x[1]["score"])
        product.best_marketplace = best_mp[0]

        return {"marketplaces": results, "best": best_mp[0], "best_score": best_mp[1]["score"]}


# --------------------------------------------
# ⚠️ ПРОГНОЗ ВОЗВРАТОВ
# --------------------------------------------
class ReturnPredictor:
    """Прогноз вероятности возвратов"""
    
    def predict(self, product: ProductBatch) -> Dict:
        """Прогноз вероятности возврата"""
        risk_factors = []
        probability = 0.10

        if product.volume > 20:
            probability += 0.10
            risk_factors.append("📦 Крупногабаритный товар")
        elif product.volume > 10:
            probability += 0.05
            risk_factors.append("📦 Средний объём")

        if product.price > 10000:
            probability += 0.08
            risk_factors.append("💰 Высокая цена")
        elif product.price > 5000:
            probability += 0.04
            risk_factors.append("💰 Средняя цена")

        high_return_categories = ["электроника", "аксессуары"]
        if product.category.lower() in high_return_categories:
            probability += 0.07
            risk_factors.append("🏷️ Категория с высоким % возвратов")

        if product.margin < 5:
            probability += 0.12
            risk_factors.append("⚠️ Очень низкая маржа")
        elif product.margin < 15:
            probability += 0.05
            risk_factors.append("⚠️ Низкая маржа")

        if not product.compatibility:
            probability += 0.15
            risk_factors.append("❓ Не указана применимость")

        if product.country and product.country.lower() in ["китай", "china"]:
            probability += 0.03
            risk_factors.append("🌏 Китайское производство")

        probability = min(probability, 0.95)

        if probability < 0.15:
            risk_level = "🟢 Низкий"
        elif probability < 0.25:
            risk_level = "🟡 Средний"
        elif probability < 0.40:
            risk_level = "🟠 Повышенный"
        else:
            risk_level = "🔴 Высокий"

        product.return_probability = probability

        return {
            "probability": probability,
            "risk_level": risk_level,
            "risk_factors": risk_factors,
            "recommendations": self._get_recommendations(risk_factors, probability)
        }

    def _get_recommendations(self, risk_factors: List[str], probability: float) -> List[str]:
        """Рекомендации по снижению риска возвратов"""
        recs = []
        if probability > 0.30:
            recs.append("🔴 Срочно улучшите описание и фото товара")
            recs.append("📞 Добавьте консультацию перед покупкой")
        if any("применимость" in f.lower() for f in risk_factors):
            recs.append("🔧 Обязательно укажите полную применимость")
        if any("крупногабарит" in f.lower() for f in risk_factors):
            recs.append("📦 Оптимизируйте упаковку")
        if not recs:
            recs.append("✅ Товар в зоне низкого риска")
        return recs


# --------------------------------------------
# 🚚 ОПТИМИЗАЦИЯ ЛОГИСТИКИ
# --------------------------------------------
class LogisticsOptimizer:
    """Оптимизация логистических расходов"""
    
    PROVIDERS = {
        "СДЭК": {"base": 250, "per_kg": 50, "per_liter": 10, "days": 3, "reliability": 0.95},
        "Почта России": {"base": 150, "per_kg": 40, "per_liter": 8, "days": 7, "reliability": 0.85},
        "Boxberry": {"base": 220, "per_kg": 45, "per_liter": 9, "days": 4, "reliability": 0.92},
        "ДПД": {"base": 280, "per_kg": 55, "per_liter": 12, "days": 3, "reliability": 0.94},
        "Яндекс Доставка": {"base": 300, "per_kg": 60, "per_liter": 15, "days": 2, "reliability": 0.97}
    }

    def optimize(self, product: ProductBatch) -> Dict:
        """Оптимизация логистики для товара"""
        results = {}

        for provider, params in self.PROVIDERS.items():
            cost = params["base"] + product.weight * params["per_kg"] + product.volume * params["per_liter"]
            score = (1 / cost) * params["reliability"] * (1 / params["days"])

            results[provider] = {
                "cost": cost,
                "days": params["days"],
                "reliability": params["reliability"],
                "score": score
            }

        cheapest = min(results.items(), key=lambda x: x[1]["cost"])
        best = max(results.items(), key=lambda x: x[1]["score"])
        fastest = min(results.items(), key=lambda x: x[1]["days"])

        product.logistics_provider = best[0]
        product.logistics_cost_optimal = best[1]["cost"]

        return {
            "all_providers": results,
            "cheapest": {"name": cheapest[0], **cheapest[1]},
            "best": {"name": best[0], **best[1]},
            "fastest": {"name": fastest[0], **fastest[1]}
        }


# --------------------------------------------
# 💰 P&L ОТЧЁТ
# --------------------------------------------
class PLReportGenerator:
    """P&L Отчёт (Прибыли и Убытки)"""
    
    def generate(self, products: List[ProductBatch]) -> Dict:
        """Генерация P&L отчёта"""
        total_revenue = sum(p.price for p in products)
        total_cogs = sum(p.cost for p in products)
        total_commission = sum(p.commission for p in products)
        total_logistics = sum(p.logistics for p in products)
        total_storage = sum(p.storage for p in products)
        total_advertising = sum(p.advertising for p in products)
        total_returns = sum(p.returns for p in products)
        total_acquiring = sum(p.acquiring for p in products)
        total_packaging = sum(p.packaging for p in products)

        total_expenses = (total_cogs + total_commission + total_logistics +
                         total_storage + total_advertising + total_returns +
                         total_acquiring + total_packaging)

        gross_profit = total_revenue - total_cogs
        operating_profit = total_revenue - total_expenses

        gross_margin = (gross_profit / total_revenue * 100) if total_revenue > 0 else 0
        operating_margin = (operating_profit / total_revenue * 100) if total_revenue > 0 else 0

        expense_structure = {
            "Себестоимость": total_cogs,
            "Комиссия МП": total_commission,
            "Логистика": total_logistics,
            "Хранение": total_storage,
            "Реклама": total_advertising,
            "Возвраты": total_returns,
            "Эквайринг": total_acquiring,
            "Упаковка": total_packaging
        }

        return {
            "revenue": total_revenue,
            "cogs": total_cogs,
            "gross_profit": gross_profit,
            "gross_margin": gross_margin,
            "expenses": expense_structure,
            "total_expenses": total_expenses,
            "operating_profit": operating_profit,
            "operating_margin": operating_margin,
            "products_count": len(products),
            "avg_check": total_revenue / len(products) if products else 0
        }


# --------------------------------------------
# 🧪 A/B ТЕСТИРОВАНИЕ ЦЕН
# --------------------------------------------
class ABTestSimulator:
    """A/B Тестирование ценовых стратегий"""
    
    def run_test(self, products: List[ProductBatch], test_days: int = 14) -> Dict:
        """Запуск A/B/C теста"""
        top_products = sorted(products, key=lambda x: x.price * 10, reverse=True)[:20]
        
        results = []
        for p in top_products:
            base_daily_sales = random.randint(5, 20)
            
            total_cost_per_unit = p.cost + p.commission + p.logistics + p.storage
            
            groups = {
                "A (текущая)": {
                    "price": p.price,
                    "daily_sales": base_daily_sales,
                    "conversion": 0.05
                },
                "B (+10%)": {
                    "price": p.price * 1.10,
                    "daily_sales": int(base_daily_sales * 0.85),
                    "conversion": 0.042
                },
                "C (-5%)": {
                    "price": p.price * 0.95,
                    "daily_sales": int(base_daily_sales * 1.25),
                    "conversion": 0.065
                }
            }
            
            for group_name, data in groups.items():
                total_sales = data["daily_sales"] * test_days
                revenue = total_sales * data["price"]
                costs = total_sales * total_cost_per_unit
                advertising_cost = revenue * 0.10
                profit = revenue - costs - advertising_cost
                data["total_sales"] = total_sales
                data["revenue"] = revenue
                data["profit"] = profit
                data["roi"] = (profit / costs * 100) if costs > 0 else 0
            
            best_group = max(groups.items(), key=lambda x: x[1]["profit"])
            
            results.append({
                "article": p.article,
                "name": p.name[:30],
                "groups": groups,
                "best_group": best_group[0],
                "best_profit": best_group[1]["profit"],
                "uplift": ((best_group[1]["profit"] - groups["A (текущая)"]["profit"]) /
                          groups["A (текущая)"]["profit"] * 100) if groups["A (текущая)"]["profit"] > 0 else 0
            })
        
        total_uplift = np.mean([r["uplift"] for r in results])
        best_strategy = "B (+10%)" if total_uplift > 5 else "C (-5%)" if total_uplift < -3 else "A (текущая)"
        
        return {
            "tests": results,
            "total_uplift": total_uplift,
            "recommended_strategy": best_strategy,
            "test_days": test_days,
            "statistical_significance": abs(total_uplift) > 3
        }


# --------------------------------------------
# 📧 EMAIL УВЕДОМЛЕНИЯ
# --------------------------------------------
class NotificationManager:
    """Уведомления по email"""
    
    @staticmethod
    def send_alert(email: str, password: str, to_email: str, subject: str, body: str) -> bool:
        """Отправка email уведомления"""
        try:
            msg = MIMEMultipart()
            msg['From'] = email
            msg['To'] = to_email
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'html'))

            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email, password)
            server.send_message(msg)
            server.quit()
            return True
        except Exception as e:
            logger.error(f"Email error: {e}")
            return False

    @staticmethod
    def generate_alerts(products: List[ProductBatch]) -> List[Dict]:
        """Генерация алертов по товарам"""
        alerts = []

        critical_stock = [p for p in products if "Критический" in p.stock_status]
        if critical_stock:
            alerts.append({
                "type": "🔴 Критический запас",
                "count": len(critical_stock),
                "message": f"{len(critical_stock)} товаров требуют срочного заказа",
                "priority": "high"
            })

        loss_making = [p for p in products if p.unit_profit < 0]
        if loss_making:
            alerts.append({
                "type": "💸 Убыточные товары",
                "count": len(loss_making),
                "message": f"{len(loss_making)} товаров продаются в убыток",
                "priority": "high"
            })

        high_return = [p for p in products if p.return_probability > 0.35]
        if high_return:
            alerts.append({
                "type": "⚠️ Высокий риск возвратов",
                "count": len(high_return),
                "message": f"{len(high_return)} товаров с риском возвратов > 35%",
                "priority": "medium"
            })

        low_cm = [p for p in products if p.contribution_margin_pct < 20]
        if low_cm:
            alerts.append({
                "type": "💎 Критический Contribution Margin",
                "count": len(low_cm),
                "message": f"{len(low_cm)} товаров с CM < 20% — срочно пересмотрите цену!",
                "priority": "high"
            })

        growing = [p for p in products if "Растущий" in p.demand_trend]
        if growing:
            alerts.append({
                "type": "📈 Растущий спрос",
                "count": len(growing),
                "message": f"{len(growing)} товаров с растущим спросом - увеличьте запасы!",
                "priority": "low"
            })

        return alerts


# --------------------------------------------
# 🤖 AI-АНАЛИЗАТОР
# --------------------------------------------
class AIAnalyzer:
    """AI анализ через DeepSeek API"""
    
    def __init__(self, api_key: str):
        self.api_key = api_key
        self.api_url = "https://api.deepseek.com/v1/chat/completions"

    def analyze_products(self, products: List[ProductBatch]) -> List[ProductBatch]:
        """AI анализ товаров"""
        products_to_analyze = products[:CONFIG['max_ai_products']]

        products_data = [{
            "article": p.article,
            "name": p.name,
            "price": p.price,
            "cost": p.cost,
            "margin": p.margin,
            "category": p.category
        } for p in products_to_analyze]

        prompt = f"""
        Проанализируй товары автозапчастей и дай рекомендации:
        {json.dumps(products_data[:20], ensure_ascii=False)}

        Для каждого верни JSON:
        {{"article": "...", "recommended_price": число, "optimization_tips": ["совет1", "совет2"]}}
        """

        response = self._call_api(prompt)
        if response:
            try:
                recommendations = json.loads(response)
                rec_map = {r['article']: r for r in recommendations}
                for p in products_to_analyze:
                    if p.article in rec_map:
                        rec = rec_map[p.article]
                        p.ai_recommended_price = rec.get('recommended_price', p.price)
                        p.ai_optimization_tips = rec.get('optimization_tips', [])
            except Exception as e:
                logger.error(f"AI parse error: {e}")

        return products

    def _call_api(self, prompt: str) -> Optional[str]:
        """Вызов DeepSeek API"""
        try:
            headers = {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            }
            data = {
                "model": "deepseek-chat",
                "messages": [
                    {"role": "system", "content": "Ты - эксперт по автозапчастям. Отвечай ТОЛЬКО JSON."},
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.3,
                "max_tokens": 1500
            }
            response = requests.post(self.api_url, headers=headers, json=data, timeout=30)
            if response.status_code == 200:
                content = response.json()['choices'][0]['message']['content']
                json_match = re.search(r'(\{.*\}|\[.*\])', content, re.DOTALL)
                if json_match:
                    return json_match.group()
        except Exception as e:
            logger.error(f"AI API error: {e}")
        return None


# --------------------------------------------
# 📊 МЕНЕДЖЕР ТАРИФОВ
# --------------------------------------------
class TariffManager:
    """Управление тарифами маркетплейсов"""
    
    def __init__(self):
        self.tariffs = {
            'Ozon': {'commission': 0.10, 'logistics_base': 50, 'logistics_per_liter': 15,
                    'logistics_per_kg': 20, 'storage_per_liter': 0.8, 'acquiring': 0.025, 'return_rate': 0.12},
            'Wildberries': {'commission': 0.12, 'logistics_base': 40, 'logistics_per_liter': 12,
                           'logistics_per_kg': 18, 'storage_per_liter': 0.9, 'acquiring': 0.028, 'return_rate': 0.15},
            'Яндекс Маркет': {'commission': 0.11, 'logistics_base': 60, 'logistics_per_liter': 10,
                            'logistics_per_kg': 15, 'storage_per_liter': 0.7, 'acquiring': 0.022, 'return_rate': 0.10},
            'AliExpress': {'commission': 0.08, 'logistics_base': 45, 'logistics_per_liter': 12,
                          'logistics_per_kg': 16, 'storage_per_liter': 0.6, 'acquiring': 0.03, 'return_rate': 0.08},
            'Мегамаркет': {'commission': 0.09, 'logistics_base': 55, 'logistics_per_liter': 14,
                          'logistics_per_kg': 19, 'storage_per_liter': 0.75, 'acquiring': 0.026, 'return_rate': 0.11}
        }

    def get_tariffs(self, marketplace: str) -> Dict:
        """Получение тарифов для маркетплейса"""
        return self.tariffs.get(marketplace, self.tariffs['Ozon'])


# --------------------------------------------
# 🔄 БАТЧ-ПРОЦЕССОР
# --------------------------------------------
class BatchProcessor:
    """Обработка данных и расчет всех показателей"""
    
    def __init__(self):
        self.tariff_manager = TariffManager()
        self.forecaster = DemandForecaster()
        self.inventory_analyzer = InventoryAnalyzer()
        self.inventory_optimizer = InventoryOptimizer()
        self.multi_mp_analyzer = MultiMarketplaceAnalyzer()
        self.return_predictor = ReturnPredictor()
        self.logistics_optimizer = LogisticsOptimizer()
        self.unit_economics = UnitEconomicsAnalyzer()

    def process_file(self, file_obj, marketplace: str, scheme: str,
                    category: str, season: str, quality: int,
                    days_storage: int, include_acquiring: bool,
                    include_advertising: bool, advertising_rate: float = 0.15) -> List[ProductBatch]:
        """Обработка файла с товарами"""
        tariffs = self.tariff_manager.get_tariffs(marketplace)

        if file_obj.name.endswith('.csv'):
            df = self._read_csv(file_obj)
        else:
            df = pd.read_excel(file_obj, engine='openpyxl')

        df = self._normalize_columns(df)

        total_rows = len(df)
        results = []

        progress = st.progress(0, text="⏳ Обработка...")
        status_text = st.empty()

        for i, (_, row) in enumerate(df.iterrows()):
            try:
                product = self._create_product(
                    row, tariffs, marketplace, scheme, category,
                    season, quality, days_storage, include_acquiring,
                    include_advertising, advertising_rate
                )
                if product:
                    results.append(product)

                if (i + 1) % 100 == 0:
                    progress.progress(min((i + 1) / total_rows, 1.0))
                    status_text.text(f"Обработано: {i + 1:,} / {total_rows:,}")
            except Exception as e:
                logger.error(f"Row error: {e}")
                continue

        progress.progress(0.9, text="💎 Выполняем анализ юнит-экономики...")
        status_text.text("💎 Расчёт Contribution Margin, LTV/CAC, точки безубыточности...")

        self.inventory_analyzer.abc_xyz_matrix(results)

        for product in results:
            forecast = self.forecaster.forecast(product.daily_sales_history)
            product.demand_forecast_30 = forecast["forecast_30"]
            product.demand_forecast_60 = forecast["forecast_60"]
            product.demand_forecast_90 = forecast["forecast_90"]
            product.demand_trend = forecast["trend"]

            self.inventory_optimizer.optimize_product(product)
            self.multi_mp_analyzer.analyze_product(product)
            self.return_predictor.predict(product)
            self.logistics_optimizer.optimize(product)

            self.unit_economics.analyze_product(product)

        progress.progress(1.0)
        status_text.text(f"✅ Обработано {len(results):,} товаров с юнит-экономикой!")

        return results

    def _read_csv(self, file_obj) -> pd.DataFrame:
        """Чтение CSV с определением разделителя"""
        sample = file_obj.read(1024).decode('utf-8', errors='ignore')
        file_obj.seek(0)
        delimiter = self._detect_delimiter(sample)
        return pd.read_csv(file_obj, delimiter=delimiter, encoding='utf-8-sig', low_memory=False)

    def _normalize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Нормализация названий колонок"""
        rename_map = {
            'Артикул': 'Артикул', 'Наименование': 'Наименование',
            'Применимость': 'Применимость', 'Категория': 'Категория',
            'Страна': 'Страна', 'OE номера': 'OE номера', 'Аналоги': 'Аналоги',
            'Цена': 'Цена', 'Себестоимость': 'Себестоимость',
            'Длина': 'Длина', 'Ширина': 'Ширина', 'Высота': 'Высота', 'Вес': 'Вес'
        }
        for old, new in rename_map.items():
            if old in df.columns:
                df.rename(columns={old: new}, inplace=True)
        return df

    def _detect_delimiter(self, sample: str) -> str:
        """Определение разделителя в CSV"""
        counts = {d: sample.count(d) for d in [';', ',', '\t', '|']}
        return max(counts, key=counts.get) if counts else ','

    def _create_product(self, row, tariffs, marketplace, scheme, category,
                       season, quality, days_storage, include_acquiring,
                       include_advertising, advertising_rate) -> Optional[ProductBatch]:
        """Создание товара из строки данных"""
        article = self._safe_get(row, 'Артикул', '')
        if not article:
            return None

        name = self._safe_get(row, 'Наименование', '')
        compatibility = self._safe_get(row, 'Применимость', '')
        category_val = self._safe_get(row, 'Категория', category)
        country = self._safe_get(row, 'Страна', '')
        oe_str = self._safe_get(row, 'OE номера', '')
        analogs_str = self._safe_get(row, 'Аналоги', '')

        price = self._safe_float(row, 'Цена', 0)
        cost = self._safe_float(row, 'Себестоимость', 0)
        length = self._safe_float(row, 'Длина', 0)
        width = self._safe_float(row, 'Ширина', 0)
        height = self._safe_float(row, 'Высота', 0)
        weight = self._safe_float(row, 'Вес', 0)

        if price <= 0 or cost <= 0:
            return None

        oe_numbers = self._split_string(oe_str)
        analogs = self._split_string(analogs_str)

        volume = (length * width * height) / 1000 if all([length, width, height]) else 0

        commission_rate = tariffs.get('commission', 0.10)
        commission = price * commission_rate

        logistics = tariffs.get('logistics_base', 50)
        if volume > 0:
            logistics += volume * tariffs.get('logistics_per_liter', 15)
        if weight > 0:
            logistics += weight * tariffs.get('logistics_per_kg', 20)

        storage = tariffs.get('storage_per_liter', 0.8) * volume * (days_storage / 30)
        acquiring = price * tariffs.get('acquiring', 0.025) if include_acquiring else 0
        advertising = price * advertising_rate if include_advertising else 0
        returns = price * tariffs.get('return_rate', 0.15)
        packaging = 50
        customs = 0
        other = 0

        total_cost = cost + commission + logistics + storage + acquiring + advertising + returns + packaging + customs + other
        profit = price - total_cost
        margin = (profit / price * 100) if price > 0 else 0
        roi = (profit / cost * 100) if cost > 0 else 0

        risk_score = min(
            tariffs.get('return_rate', 0) * 0.3 +
            commission_rate * 0.2 +
            min((logistics / price if price > 0 else 0) / 0.15, 1) * 0.2 +
            min(volume / 50, 1) * 0.15 + 0.15,
            1.0
        )

        return ProductBatch(
            article=article, name=name, compatibility=compatibility,
            category=category_val, country=country, oe_numbers=oe_numbers,
            analogs=analogs, price=price, cost=cost,
            length=length, width=width, height=height, weight=weight,
            volume=volume, commission=commission, logistics=logistics,
            storage=storage, acquiring=acquiring, advertising=advertising,
            returns=returns, packaging=packaging, customs=customs,
            other_costs=other, total_cost=total_cost,
            profit=profit, margin=margin, roi=roi, risk_score=risk_score,
            daily_sales_history=[random.randint(5, 20) for _ in range(30)]
        )

    def _safe_get(self, row, col, default):
        """Безопасное получение значения"""
        try:
            val = row.get(col, default)
            return default if pd.isna(val) else val
        except:
            return default

    def _safe_float(self, row, col, default):
        """Безопасное преобразование в float"""
        try:
            val = row.get(col, default)
            return default if pd.isna(val) else float(val)
        except:
            return default

    def _split_string(self, text):
        """Разделение строки по разделителям"""
        if not text or pd.isna(text):
            return []
        text = str(text)
        for sep in CONFIG['separators']:
            if sep in text:
                return [s.strip() for s in text.split(sep) if s.strip()]
        return [text.strip()] if text.strip() else []


# --------------------------------------------
# 📥 ЭКСПОРТ
# --------------------------------------------
class LargeDataExporter:
    """Экспорт данных в Excel с формулами"""
    
    @staticmethod
    def to_excel_with_formulas(products: List[ProductBatch], marketplace: str = "Ozon",
                               days_storage: int = 30, tariffs: Dict = None) -> bytes:
        """Экспорт в Excel с формулами"""
        output = io.BytesIO()
        wb = Workbook()

        if tariffs is None:
            tariffs = {'commission': 0.10, 'logistics_base': 50, 'logistics_per_liter': 15,
                      'logistics_per_kg': 20, 'storage_per_liter': 0.8, 'acquiring': 0.025, 'return_rate': 0.15}

        ws_ref = wb.active
        ws_ref.title = "Справка"
        ref_data = [
            ['Параметр', 'Значение'],
            ['Комиссия', tariffs.get('commission', 0.10)],
            ['Логистика база', tariffs.get('logistics_base', 50)],
            ['Логистика за литр', tariffs.get('logistics_per_liter', 15)],
            ['Логистика за кг', tariffs.get('logistics_per_kg', 20)],
            ['Хранение за литр', tariffs.get('storage_per_liter', 0.8)],
            ['Эквайринг', tariffs.get('acquiring', 0.025)],
            ['Возвраты', tariffs.get('return_rate', 0.15)],
            ['Дней хранения', days_storage]
        ]

        for r_idx, row in enumerate(ref_data, 1):
            for c_idx, value in enumerate(row, 1):
                ws_ref.cell(row=r_idx, column=c_idx, value=value)

        ws = wb.create_sheet("Товары")
        headers = ['Артикул', 'Наименование', 'Цена', 'Себестоимость', 'Прибыль',
                  'Маржа %', 'CM %', 'LTV/CAC', 'ДРР %', 'ABC', 'XYZ', 'Прогноз 30д',
                  'Лучший МП', 'Статус']

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True)

        for r_idx, p in enumerate(products[:50000], 2):
            ws.cell(row=r_idx, column=1, value=p.article)
            ws.cell(row=r_idx, column=2, value=p.name[:50])
            ws.cell(row=r_idx, column=3, value=p.price)
            ws.cell(row=r_idx, column=4, value=p.cost)
            ws.cell(row=r_idx, column=5, value=f"=C{r_idx}-D{r_idx}")
            ws.cell(row=r_idx, column=6, value=f"=E{r_idx}/C{r_idx}*100")
            ws.cell(row=r_idx, column=7, value=round(p.contribution_margin_pct, 1))
            ws.cell(row=r_idx, column=8, value=round(p.ltv_cac_ratio, 2))
            ws.cell(row=r_idx, column=9, value=round(p.drr, 1))
            ws.cell(row=r_idx, column=10, value=p.abc_category)
            ws.cell(row=r_idx, column=11, value=p.xyz_category)
            ws.cell(row=r_idx, column=12, value=p.demand_forecast_30)
            ws.cell(row=r_idx, column=13, value=p.best_marketplace)
            ws.cell(row=r_idx, column=14, value=p.stock_status)

        wb.save(output)
        output.seek(0)
        return output.getvalue()


# --------------------------------------------
# 🎨 ОСНОВНОЕ ПРИЛОЖЕНИЕ
# --------------------------------------------
class UltimateAutoApp:
    """Полное приложение с упором на юнит-экономику"""
    
    def __init__(self):
        self.processor = BatchProcessor()
        self.exporter = LargeDataExporter()
        self.ai_analyzer = None
        self.notification_manager = NotificationManager()
        self.pl_reporter = PLReportGenerator()
        self.ab_tester = ABTestSimulator()
        self.product_tracker = None
        self.unit_economics = UnitEconomicsAnalyzer()
        
        self._init_session_state()

    def _init_session_state(self):
        """Инициализация всех ключей session state"""
        defaults = {
            'results': None,
            'stats': {"products_processed": 0, "total_profit": 0},
            'selenium_results': None,
            'ab_test_result': None,
            'use_selenium': False,
            'headless': True,
            'marketplace': "Ozon",
            'scheme': "FBO",
            'category': "расходники",
            'season': 'всесезон',
            'quality': 70,
            'days_storage': 30,
            'include_acquiring': True,
            'include_advertising': False,
            'advertising_rate': 0.15,
            'fixed_costs': 50000,
            'avg_orders': 100,
            'deepseek_key': '',
            'email_login': '',
            'email_password': '',
            'email_to': ''
        }
        
        for key, default_value in defaults.items():
            if key not in st.session_state:
                st.session_state[key] = default_value

    def run(self):
        """Запуск приложения"""
        st.set_page_config(
            page_title=CONFIG["app_name"] + " v14.0",
            page_icon="🚗",
            layout="wide",
            initial_sidebar_state="expanded"
        )

        self._apply_css()
        self._render_header()

        with st.sidebar:
            self._render_sidebar()

        self._render_main_content()

    def _apply_css(self):
        """Применение CSS стилей"""
        st.markdown("""
        <style>
        .main-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 2rem; border-radius: 15px; color: white;
            text-align: center; margin-bottom: 2rem;
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        }
        .main-header h1 {
            font-size: 2.5rem; margin: 0;
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        }
        .feature-badge {
            display: inline-block; padding: 0.3rem 0.8rem; margin: 0.2rem;
            border-radius: 15px; font-size: 0.75rem; font-weight: 600;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }
        .unit-metric {
            background: white; padding: 1rem; border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.08); text-align: center;
            border-left: 4px solid #667eea;
        }
        .tooltip {
            position: relative;
            display: inline-block;
            cursor: help;
            border-bottom: 1px dotted #666;
        }
        .tooltip .tooltiptext {
            visibility: hidden;
            width: 300px;
            background-color: #555;
            color: #fff;
            text-align: center;
            border-radius: 6px;
            padding: 10px;
            position: absolute;
            z-index: 1;
            bottom: 125%;
            left: 50%;
            margin-left: -150px;
            opacity: 0;
            transition: opacity 0.3s;
            font-size: 0.9rem;
            font-weight: normal;
        }
        .tooltip:hover .tooltiptext {
            visibility: visible;
            opacity: 1;
        }
        </style>
        """, unsafe_allow_html=True)

    def _render_header(self):
        """Отображение заголовка"""
        st.markdown(f"""
        <div class="main-header">
            <h1>🚗 Платформа автозапчастей v{CONFIG['version']}</h1>
            <p style="font-size: 1.1rem; opacity: 0.95;">
                💎 УПОР НА ЮНИТ-ЭКОНОМИКУ: Contribution Margin • LTV/CAC • Break-Even • Чувствительность
            </p>
            <div style="margin-top: 1rem;">
                <span class="feature-badge">💎 Юнит-экономика</span>
                <span class="feature-badge">📊 Чувствительность</span>
                <span class="feature-badge">🎯 ABC/XYZ</span>
                <span class="feature-badge">📈 Прогнозы</span>
                <span class="feature-badge">📦 Запасы</span>
                <span class="feature-badge">🏪 5 МП</span>
                <span class="feature-badge">🕷️ Парсер</span>
                <span class="feature-badge">🚚 Логистика</span>
                <span class="feature-badge">💰 P&L</span>
                <span class="feature-badge">🧪 A/B</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

    def _render_sidebar(self):
        """Боковая панель с настройками"""
        st.markdown("### ⚙️ Настройки")
        
        st.markdown("""
        <div style="background: #f0f2f6; padding: 0.8rem; border-radius: 8px; 
                    margin-bottom: 1rem; font-size: 0.85rem;">
        💡 Наведите на элемент для подсказки
        </div>
        """, unsafe_allow_html=True)

        st.subheader("🔑 API Ключи")

        st.text_input(
            "DeepSeek API ключ",
            type="password",
            placeholder="sk-...",
            key="deepseek_key",
            help="Получите ключ на platform.deepseek.com"
        )

        if st.session_state.deepseek_key:
            self.ai_analyzer = AIAnalyzer(st.session_state.deepseek_key)
            st.success("✅ DeepSeek подключён")

        with st.expander("📧 Email для уведомлений"):
            st.markdown("""
            <span class="tooltip">📧 Настройка уведомлений
            <span class="tooltiptext">Используйте пароль приложения Google для безопасности</span>
            </span>
            """, unsafe_allow_html=True)
            st.text_input("Email (отправитель)", key="email_login")
            st.text_input("Пароль приложения", type="password", key="email_password")
            st.text_input("Куда отправлять", key="email_to")

        st.markdown("---")

        st.subheader("🕷️ Парсер своих товаров")

        if SELENIUM_AVAILABLE:
            st.success("✅ Selenium установлен")
            st.checkbox(
                "Использовать парсинг", 
                value=st.session_state.use_selenium,
                key="use_selenium",
                help="Парсинг через Selenium WebDriver"
            )

            if st.session_state.use_selenium:
                st.checkbox(
                    "Headless режим", 
                    value=st.session_state.headless,
                    key="headless",
                    help="Запуск браузера в фоновом режиме"
                )
        else:
            st.warning("⚠️ Selenium не установлен")
            st.code("pip install selenium webdriver-manager", language="bash")
            st.session_state.use_selenium = False

        st.markdown("---")

        st.subheader("📦 Параметры")
        
        st.markdown("""
        <span class="tooltip">📦 Настройки расчета
        <span class="tooltiptext">Влияют на расчет себестоимости и логистики</span>
        </span>
        """, unsafe_allow_html=True)
        
        marketplace_idx = 0
        if st.session_state.marketplace in CONFIG['supported_marketplaces']:
            marketplace_idx = CONFIG['supported_marketplaces'].index(st.session_state.marketplace)
        
        st.selectbox(
            "Маркетплейс", 
            CONFIG['supported_marketplaces'], 
            index=marketplace_idx,
            key="marketplace"
        )
        
        scheme_idx = 0
        schemes = ["FBO", "FBS", "DBS"]
        if st.session_state.scheme in schemes:
            scheme_idx = schemes.index(st.session_state.scheme)
        
        st.selectbox(
            "Схема", 
            schemes,
            index=scheme_idx,
            key="scheme"
        )
        
        categories = ["расходники", "аксессуары", "электроника", "химия", "запчасти"]
        category_idx = 0
        if st.session_state.category in categories:
            category_idx = categories.index(st.session_state.category)
        
        st.selectbox(
            "Категория", 
            categories,
            index=category_idx,
            key="category"
        )
        
        seasons = ['всесезон', 'лето', 'зима', 'весна', 'осень']
        season_idx = 0
        if st.session_state.season in seasons:
            season_idx = seasons.index(st.session_state.season)
        
        st.selectbox(
            "Сезон", 
            seasons,
            index=season_idx,
            key="season"
        )
        
        st.slider(
            "Качество (0-100)", 
            0, 100, 
            value=st.session_state.quality,
            key="quality"
        )
        
        st.number_input(
            "Хранение (дней)", 
            1, 180, 
            value=st.session_state.days_storage,
            key="days_storage"
        )
        
        st.checkbox(
            "Эквайринг", 
            value=st.session_state.include_acquiring,
            key="include_acquiring"
        )
        
        st.checkbox(
            "Реклама", 
            value=st.session_state.include_advertising,
            key="include_advertising"
        )
        
        if st.session_state.include_advertising:
            advertising_rate = st.slider(
                "ДРР (%)", 
                0, 50, 
                value=int(st.session_state.advertising_rate * 100),
                key="advertising_rate"
            ) / 100
            st.session_state.advertising_rate = advertising_rate
        else:
            st.session_state.advertising_rate = 0

        st.markdown("---")
        st.subheader("💎 Юнит-экономика")
        
        st.markdown("""
        <span class="tooltip">💎 Настройки расчета юнит-экономики
        <span class="tooltiptext">Влияют на LTV, CAC, Break-Even и окупаемость</span>
        </span>
        """, unsafe_allow_html=True)
        
        st.number_input(
            "Постоянные расходы/мес (₽)", 
            0, 1000000, 
            value=st.session_state.fixed_costs,
            key="fixed_costs",
            help="Аренда, зарплаты, офис и т.д."
        )
        
        st.number_input(
            "Среднее кол-во заказов/мес", 
            10, 10000, 
            value=st.session_state.avg_orders,
            key="avg_orders",
            help="Используется для расчета окупаемости"
        )

        st.markdown("---")
        st.caption(f"🕐 {datetime.now().strftime('%d.%m.%Y %H:%M')}")
        if st.session_state.get('results') is not None:
            st.success(f"📦 Обработано: {len(st.session_state.results):,}")

    def _render_main_content(self):
        """Основной контент"""
        st.subheader("📁 Загрузка данных")
        
        st.markdown("""
        <span class="tooltip">📁 Загрузите файл с товарами
        <span class="tooltiptext">Поддерживаются Excel (.xlsx, .xls) и CSV файлы</span>
        </span>
        """, unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader(
            "Загрузите файл (Excel/CSV)",
            type=["xlsx", "xls", "csv"],
            key="file_uploader",
            help="Файл должен содержать колонки: Артикул, Наименование, Цена, Себестоимость, Применимость"
        )

        if uploaded_file:
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                if st.button("🚀 Обработать", type="primary", use_container_width=True,
                           help="Запустить обработку и расчет всех показателей"):
                    self._process_file(uploaded_file)
            with col2:
                if st.button("🤖 AI Анализ", use_container_width=True,
                           help="Использовать AI для рекомендаций по цене"):
                    self._run_ai()
            with col3:
                if st.button("🕷️ Парсинг", use_container_width=True,
                           help="Проверить наличие товаров на маркетплейсе"):
                    self._run_selenium_parsing()
            with col4:
                if st.button("📥 Экспорт", use_container_width=True,
                           help="Скачать отчет в Excel с формулами"):
                    self._export()

        if st.button("🔄 Демо-данные (500 товаров)", use_container_width=True,
                   help="Загрузить тестовые данные для ознакомления"):
            self._load_demo_data()

        if st.session_state.get('results') is not None:
            self._render_analytics()
        else:
            self._render_welcome()

    def _process_file(self, file_obj):
        """Обработка загруженного файла"""
        try:
            with st.spinner("🚀 Обрабатываю файл с юнит-экономикой..."):
                results = self.processor.process_file(
                    file_obj,
                    st.session_state.marketplace,
                    st.session_state.scheme,
                    st.session_state.category,
                    st.session_state.season,
                    st.session_state.quality,
                    st.session_state.days_storage,
                    st.session_state.include_acquiring,
                    st.session_state.include_advertising,
                    st.session_state.advertising_rate
                )
                st.session_state.results = results
                st.session_state.stats = {
                    "products_processed": len(results),
                    "total_profit": sum(p.unit_profit for p in results),
                    "ai_used": False, 
                    "abc_done": True, 
                    "forecast_done": True,
                    "multi_mp_done": True, 
                    "ab_test_done": False, 
                    "selenium_used": False,
                    "unit_economics_done": True
                }
                st.success(f"✅ Обработано {len(results):,} товаров с юнит-экономикой!")
                st.rerun()
        except Exception as e:
            st.error(f"❌ Ошибка: {e}")
            logger.error(f"Process error: {e}")

    def _load_demo_data(self):
        """Загрузка демо-данных"""
        demo_products = []
        categories = ["расходники", "аксессуары", "электроника", "фильтры", "тормоза"]
        brands = ["Toyota", "BMW", "Mercedes", "Hyundai", "Kia", "VAG", "Ford"]
        models = ["Camry", "X5", "E-Class", "Solaris", "Rio", "Passat", "Focus"]

        progress_bar = st.progress(0)
        status_text = st.empty()

        for i in range(500):
            brand = random.choice(brands)
            model = random.choice(models)
            year = random.randint(2010, 2023)
            price = random.uniform(500, 15000)
            cost = price * random.uniform(0.3, 0.7)

            p = ProductBatch(
                article=f"ART-{i+1:05d}",
                name=f"{brand} {model} Запчасть #{i+1}",
                compatibility=f"{brand} {model} {year}",
                category=random.choice(categories),
                country=random.choice(["Япония", "Германия", "Корея", "Китай", "Россия"]),
                price=price, cost=cost,
                length=random.uniform(10, 50), width=random.uniform(10, 40),
                height=random.uniform(5, 30), weight=random.uniform(0.5, 10),
                daily_sales_history=[random.randint(5, 25) for _ in range(30)]
            )

            p.volume = (p.length * p.width * p.height) / 1000
            p.commission = p.price * 0.10
            p.logistics = 50 + p.volume * 15 + p.weight * 20
            p.storage = 0.8 * p.volume
            p.acquiring = p.price * 0.025
            p.advertising = p.price * 0.15
            p.returns = p.price * 0.12
            p.packaging = 50
            p.total_cost = p.cost + p.commission + p.logistics + p.storage + p.acquiring + p.advertising + p.returns + p.packaging
            p.profit = p.price - p.total_cost
            p.margin = (p.profit / p.price * 100) if p.price > 0 else 0
            demo_products.append(p)

            if (i + 1) % 50 == 0:
                progress_bar.progress((i + 1) / 500)
                status_text.text(f"📦 Создано {i+1}/500 товаров...")

        status_text.text("💎 Выполняем анализ юнит-экономики...")
        progress_bar.progress(0.9)

        self.processor.inventory_analyzer.abc_xyz_matrix(demo_products)
        for p in demo_products:
            forecast = self.processor.forecaster.forecast(p.daily_sales_history)
            p.demand_forecast_30 = forecast["forecast_30"]
            p.demand_forecast_60 = forecast["forecast_60"]
            p.demand_forecast_90 = forecast["forecast_90"]
            p.demand_trend = forecast["trend"]
            self.processor.inventory_optimizer.optimize_product(p)
            self.processor.multi_mp_analyzer.analyze_product(p)
            self.processor.return_predictor.predict(p)
            self.processor.logistics_optimizer.optimize(p)
            self.processor.unit_economics.analyze_product(p)

        progress_bar.progress(1.0)
        status_text.empty()

        st.session_state.results = demo_products
        st.session_state.stats = {
            "products_processed": len(demo_products),
            "total_profit": sum(p.unit_profit for p in demo_products),
            "ai_used": False, 
            "abc_done": True, 
            "forecast_done": True,
            "multi_mp_done": True, 
            "ab_test_done": False, 
            "selenium_used": False,
            "unit_economics_done": True
        }
        st.success(f"✅ Загружено {len(demo_products)} демо-товаров с юнит-экономикой!")
        st.rerun()

    def _run_ai(self):
        """Запуск AI анализа"""
        if st.session_state.get('results') is None:
            st.warning("⚠️ Сначала обработайте файл")
            return
        if not self.ai_analyzer:
            st.warning("⚠️ Введите DeepSeek API ключ")
            return
        with st.spinner("🤖 AI анализирует..."):
            st.session_state.results = self.ai_analyzer.analyze_products(st.session_state.results)
            st.session_state.stats['ai_used'] = True
            st.success("✅ AI-анализ завершён!")
            st.rerun()

    def _run_selenium_parsing(self):
        """Запуск Selenium парсинга"""
        if st.session_state.get('results') is None:
            st.warning("⚠️ Сначала обработайте файл")
            return

        if not SELENIUM_AVAILABLE:
            st.warning("⚠️ Selenium не установлен. Используем симуляцию.")
            self._simulate_competitor_parsing()
            return

        self.product_tracker = MyProductTracker(use_selenium=True)
        products = st.session_state.results[:5]

        progress_bar = st.progress(0)
        status_text = st.empty()

        all_results = []

        for i, product in enumerate(products):
            status_text.text(f"🕷️ Парсим {i+1}/{len(products)}: {product.name[:30]}...")

            def progress_callback(progress, message):
                progress_bar.progress(progress, text=message)

            result = self.product_tracker.track_my_product(
                product, st.session_state.marketplace, progress_callback=progress_callback
            )
            all_results.append(result)

            time.sleep(0.5)

        progress_bar.progress(1.0, text="✅ Парсинг завершён!")
        status_text.empty()

        st.session_state.selenium_results = all_results
        st.session_state.stats['selenium_used'] = True

        st.success(f"✅ Спарсено {len(all_results)} товаров!")
        st.rerun()

    def _simulate_competitor_parsing(self):
        """Симуляция парсинга"""
        self.product_tracker = MyProductTracker(use_selenium=False)
        products = st.session_state.results[:10]

        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, product in enumerate(products):
            status_text.text(f"🔄 Симуляция {i+1}/{len(products)}: {product.name[:30]}...")
            self.product_tracker.track_my_product(product, st.session_state.marketplace)
            progress_bar.progress((i + 1) / len(products))
            time.sleep(0.1)

        st.session_state.stats['selenium_used'] = True
        st.success("✅ Симуляция завершена")
        st.rerun()

    def _export(self):
        """Экспорт в Excel"""
        if st.session_state.get('results') is None:
            st.warning("⚠️ Сначала обработайте файл")
            return
        excel_data = self.exporter.to_excel_with_formulas(
            st.session_state.results,
            st.session_state.marketplace,
            st.session_state.days_storage
        )
        st.download_button(
            "📥 Скачать Excel",
            data=excel_data,
            file_name=f"автозапчасти_юнит_экономика_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    def _render_analytics(self):
        """Рендер аналитических вкладок"""
        products = st.session_state.results

        if not products:
            st.warning("⚠️ Нет данных для отображения. Загрузите файл или используйте демо-данные.")
            return

        total = len(products)
        profitable = sum(1 for p in products if p.unit_profit > 0)
        total_profit = sum(p.unit_profit for p in products)
        avg_margin = np.mean([p.margin for p in products]) if products else 0
        avg_cm = np.mean([p.contribution_margin_pct for p in products]) if products else 0

        cols = st.columns(5)
        cols[0].metric("📦 Всего", f"{total:,}")
        cols[1].metric("💰 Прибыльных", f"{profitable:,}")
        cols[2].metric("💵 Прибыль", f"{total_profit:,.0f} ₽")
        cols[3].metric("📊 Маржа", f"{avg_margin:.1f}%")
        cols[4].metric("💎 CM", f"{avg_cm:.1f}%")

        tabs = st.tabs([
            "💎 Юнит-экономика",
            "📊 Дашборд",
            "🎯 ABC/XYZ",
            "📈 Прогноз",
            "📦 Запасы",
            "🏪 Маркетплейсы",
            "🚚 Логистика",
            "💰 P&L",
            "🧪 A/B Тест",
        ])

        with tabs[0]:
            self._tab_unit_economics(products)
        with tabs[1]:
            self._tab_dashboard(products)
        with tabs[2]:
            self._tab_abc_xyz(products)
        with tabs[3]:
            self._tab_forecast(products)
        with tabs[4]:
            self._tab_inventory(products)
        with tabs[5]:
            self._tab_multi_mp(products)
        with tabs[6]:
            self._tab_logistics(products)
        with tabs[7]:
            self._tab_pl(products)
        with tabs[8]:
            self._tab_ab_test(products)

    def _tab_unit_economics(self, products):
        """Вкладка юнит-экономики"""
        st.subheader("💎 Детальная юнит-экономика")
        
        st.markdown("""
        <div style="background: #e7f3ff; padding: 1rem; border-radius: 10px; margin-bottom: 1rem;">
        💡 <b>Что такое юнит-экономика?</b> Это анализ прибыльности одной единицы товара.
        Позволяет понять, сколько вы зарабатываете на каждом товаре с учетом всех затрат.
        </div>
        """, unsafe_allow_html=True)

        portfolio = self.unit_economics.analyze_portfolio(products)

        st.markdown("### 📊 Показатели портфеля")

        cols = st.columns(6)
        cols[0].metric("💵 Выручка", f"{portfolio['total_revenue']:,.0f} ₽")
        cols[1].metric("💰 Прибыль", f"{portfolio['total_profit']:,.0f} ₽")
        cols[2].metric("💎 Средний CM", f"{portfolio['avg_cm_pct']:.1f}%")
        cols[3].metric("📈 LTV/CAC", f"{portfolio['avg_ltv_cac']:.2f}")
        cols[4].metric("📢 Средний ДРР", f"{portfolio['avg_drr']:.1f}%")
        cols[5].metric("⏳ Окупаемость", f"{portfolio['avg_payback_days']:.0f} дн")

        st.markdown("### 📊 Распределение товаров по Contribution Margin")

        col1, col2 = st.columns(2)

        with col1:
            cm_dist = portfolio['cm_distribution']
            if cm_dist and sum(cm_dist.values()) > 0:
                fig = px.pie(
                    values=list(cm_dist.values()),
                    names=list(cm_dist.keys()),
                    title='Распределение по CM',
                    color_discrete_map={
                        "Критический (<20%)": "#dc3545",
                        "Низкий (20-35%)": "#ffc107",
                        "Средний (35-50%)": "#17a2b8",
                        "Высокий (>50%)": "#28a745"
                    }
                )
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Нет данных для распределения по CM")

        with col2:
            top_cm = portfolio['top_by_cm']
            if top_cm:
                df_top = pd.DataFrame([{
                    "Артикул": p.article,
                    "Название": p.name[:30],
                    "Цена": f"{p.price:,.0f} ₽",
                    "CM %": f"{p.contribution_margin_pct:.1f}%",
                    "LTV/CAC": f"{p.ltv_cac_ratio:.2f}",
                    "ДРР %": f"{p.drr:.1f}%"
                } for p in top_cm])
                st.markdown("**🏆 Топ-10 по Contribution Margin**")
                st.dataframe(df_top, use_container_width=True, hide_index=True)
            else:
                st.info("Нет данных для отображения топа по CM")

        st.markdown("---")
        st.markdown("### 🔍 Детальный анализ товара")
        st.markdown("""
        <div style="background: #f8f9fa; padding: 0.8rem; border-radius: 8px; margin-bottom: 1rem;">
        💡 Выберите товар, чтобы увидеть полный анализ юнит-экономики:
        структуру расходов, точку безубыточности, LTV/CAC и рекомендации
        </div>
        """, unsafe_allow_html=True)

        if products:
            selected_article = st.selectbox(
                "Выберите товар для анализа",
                [f"{p.article} - {p.name[:40]}" for p in products[:100]],
                key="unit_product_select"
            )

            if selected_article:
                article = selected_article.split(" - ")[0]
                product = next((p for p in products if p.article == article), None)

                if product:
                    analysis = self.unit_economics.analyze_product(
                        product,
                        fixed_costs_per_month=st.session_state.get('fixed_costs', 50000),
                        avg_orders_per_month=st.session_state.get('avg_orders', 100)
                    )

                    st.markdown("#### 📊 Ключевые показатели")
                    cols = st.columns(5)
                    cols[0].metric("💵 Цена", f"{analysis['price']:,.0f} ₽")
                    cols[1].metric("💸 Переменные расходы", f"{analysis['variable_costs']:,.0f} ₽")
                    cols[2].metric("💎 Contribution Margin",
                                  f"{analysis['contribution_margin']:,.0f} ₽",
                                  f"{analysis['contribution_margin_pct']:.1f}%")
                    cols[3].metric("💰 Прибыль/ед", f"{analysis['unit_profit']:,.0f} ₽")
                    cols[4].metric("📊 Маржа", f"{product.margin:.1f}%")

                    st.markdown("#### 📊 Структура расходов на 1 единицу")

                    col1, col2 = st.columns(2)

                    with col1:
                        cost_structure = analysis['unit_cost_structure']
                        if cost_structure and sum(cost_structure.values()) > 0:
                            fig = go.Figure(data=[go.Pie(
                                labels=list(cost_structure.keys()),
                                values=list(cost_structure.values()),
                                hole=0.4,
                                marker=dict(colors=[
                                    '#667eea', '#764ba2', '#f093fb', '#f5576c',
                                    '#4facfe', '#00f2fe', '#43e97b', '#38f9d7',
                                    '#fa709a', '#fee140'
                                ])
                            )])
                            fig.update_layout(title='Структура расходов', height=400)
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.info("Нет данных о структуре расходов")

                    with col2:
                        categories = ['Цена'] + list(cost_structure.keys()) + ['Прибыль']
                        values = [analysis['price']] + [-v for v in cost_structure.values()] + [analysis['unit_profit']]
                        measures = ['absolute'] + ['relative'] * len(cost_structure) + ['total']

                        fig = go.Figure(go.Waterfall(
                            name="Юнит-экономика",
                            orientation="v",
                            x=categories,
                            y=values,
                            measure=measures,
                            connector={"line": {"color": "rgb(63, 63, 63)"}},
                            decreasing={"marker": {"color": "#dc3545"}},
                            increasing={"marker": {"color": "#28a745"}},
                            totals={"marker": {"color": "#667eea"}}
                        ))
                        fig.update_layout(title='Waterfall: от цены к прибыли', height=400, showlegend=False)
                        st.plotly_chart(fig, use_container_width=True)

                    st.markdown("#### 📈 LTV, CAC, ROAS, ДРР")
                    cols = st.columns(4)
                    cols[0].metric("👤 LTV (оценка)", f"{analysis['ltv']:,.0f} ₽")
                    cols[1].metric("🎯 CAC (оценка)", f"{analysis['cac']:,.0f} ₽")

                    ltv_cac = analysis['ltv_cac_ratio']
                    ltv_cac_delta = "🟢 Отлично" if ltv_cac > 3 else "🟡 Норм" if ltv_cac > 1 else "🔴 Плохо"
                    cols[2].metric("📈 LTV/CAC", f"{ltv_cac:.2f}", ltv_cac_delta)

                    drr = analysis['drr']
                    drr_delta = "🟢 Низкий" if drr < 10 else "🟡 Средний" if drr < 20 else "🔴 Высокий"
                    cols[3].metric("📢 ДРР", f"{drr:.1f}%", drr_delta)

                    if analysis['roas'] > 0:
                        st.metric("💰 ROAS (возврат рекламных инвестиций)", f"{analysis['roas']:.2f}x",
                                 "🟢 Отлично" if analysis['roas'] > 5 else "🟡 Норм" if analysis['roas'] > 2 else "🔴 Низкий")

                    st.markdown("#### 🎯 Точка безубыточности (Break-Even)")
                    cols = st.columns(3)
                    cols[0].metric("📦 Безубыточность (шт)", f"{analysis['break_even_units']:,}")
                    cols[1].metric("💵 Безубыточность (₽)", f"{analysis['break_even_revenue']:,.0f} ₽")
                    cols[2].metric("⏳ Окупаемость", f"{analysis['payback_period_days']:.0f} дней")

                    st.markdown("#### 🎚️ Анализ чувствительности (что-if)")
                    sensitivity = analysis['sensitivity']
                    st.info(f"**🎯 Самый чувствительный фактор:** {sensitivity['most_sensitive']}")

                    col1, col2 = st.columns(2)

                    with col1:
                        st.markdown("**Изменение цены на ±10%**")
                        cols = st.columns(3)
                        cols[0].metric("Базовая прибыль", f"{sensitivity['base_profit']:,.0f} ₽")
                        cols[1].metric("Цена +10%", f"{sensitivity['price_up_10']:,.0f} ₽",
                                      f"{(sensitivity['price_up_10'] - sensitivity['base_profit']):+.0f} ₽")
                        cols[2].metric("Цена -10%", f"{sensitivity['price_down_10']:,.0f} ₽",
                                      f"{(sensitivity['price_down_10'] - sensitivity['base_profit']):+.0f} ₽")

                    with col2:
                        st.markdown("**Изменение себестоимости на ±10%**")
                        cols = st.columns(3)
                        cols[0].metric("Базовая прибыль", f"{sensitivity['base_profit']:,.0f} ₽")
                        cols[1].metric("Себест. +10%", f"{sensitivity['cost_up_10']:,.0f} ₽",
                                      f"{(sensitivity['cost_up_10'] - sensitivity['base_profit']):+.0f} ₽")
                        cols[2].metric("Себест. -10%", f"{sensitivity['cost_down_10']:,.0f} ₽",
                                      f"{(sensitivity['cost_down_10'] - sensitivity['base_profit']):+.0f} ₽")

                    st.markdown("#### 🎮 Интерактивный калькулятор (что-if)")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        new_price = st.number_input("Новая цена (₽)",
                                                   value=int(product.price),
                                                   min_value=0,
                                                   key="new_price")
                    with col2:
                        new_cost = st.number_input("Новая себестоимость (₽)",
                                                  value=int(product.cost),
                                                  min_value=0,
                                                  key="new_cost")
                    with col3:
                        new_ads_rate = st.slider("Новый ДРР (%)", 0, 50,
                                                int(product.advertising / product.price * 100 if product.price > 0 else 0),
                                                key="new_ads")

                    new_commission = new_price * (product.commission / product.price if product.price > 0 else 0.10)
                    new_logistics = product.logistics
                    new_storage = product.storage
                    new_acquiring = new_price * (product.acquiring / product.price if product.price > 0 else 0.025)
                    new_advertising = new_price * new_ads_rate / 100
                    new_returns = new_price * (product.returns / product.price if product.price > 0 else 0.12)
                    new_packaging = product.packaging

                    new_total_costs = (new_cost + new_commission + new_logistics + new_storage +
                                      new_acquiring + new_advertising + new_returns + new_packaging)
                    new_profit = new_price - new_total_costs
                    new_margin = (new_profit / new_price * 100) if new_price > 0 else 0

                    profit_delta = new_profit - product.unit_profit

                    st.markdown("---")
                    cols = st.columns(4)
                    cols[0].metric("💵 Новая цена", f"{new_price:,.0f} ₽")
                    cols[1].metric("💰 Новая прибыль", f"{new_profit:,.0f} ₽",
                                  f"{profit_delta:+,.0f} ₽")
                    cols[2].metric("📊 Новая маржа", f"{new_margin:.1f}%",
                                  f"{new_margin - product.margin:+.1f}%")
                    cols[3].metric("💎 Изменение",
                                  f"{(profit_delta / abs(product.unit_profit) * 100) if product.unit_profit != 0 else 0:+.1f}%"
                                  if product.unit_profit != 0 else "N/A")

                    st.markdown("#### 💡 Рекомендации по оптимизации")
                    recommendations = analysis['recommendations']
                    for rec in recommendations:
                        if "🔴" in rec:
                            st.error(rec)
                        elif "🟡" in rec:
                            st.warning(rec)
                        elif "🟢" in rec:
                            st.success(rec)
                        else:
                            st.info(rec)

    def _tab_dashboard(self, products):
        """Вкладка дашборда"""
        st.subheader("📊 Главный дашборд")

        col1, col2 = st.columns(2)

        with col1:
            margins = [p.margin for p in products]
            if margins:
                fig = go.Figure(data=[go.Histogram(x=margins, nbinsx=30, marker_color='#667eea')])
                fig.add_vline(x=0, line_dash="dash", line_color="red")
                fig.add_vline(x=20, line_dash="dash", line_color="green")
                fig.update_layout(title='📊 Распределение маржинальности', height=350)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Нет данных о маржинальности")

        with col2:
            abc_counts = {"A": 0, "B": 0, "C": 0}
            for p in products:
                if p.abc_category in abc_counts:
                    abc_counts[p.abc_category] += 1

            if sum(abc_counts.values()) > 0:
                fig = px.pie(values=list(abc_counts.values()), names=list(abc_counts.keys()),
                            title='🎯 ABC-распределение',
                            color_discrete_map={"A": "#FF6B6B", "B": "#FFD93D", "C": "#6BCF7F"})
                fig.update_layout(height=350)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Нет данных для ABC-распределения")

        st.subheader("🏆 Топ-10 по прибыли")
        top = sorted(products, key=lambda x: x.unit_profit, reverse=True)[:10]
        if top:
            df_top = pd.DataFrame([{
                "Артикул": p.article, "Название": p.name[:40],
                "Цена": f"{p.price:,.0f} ₽", "Прибыль": f"{p.unit_profit:,.0f} ₽",
                "CM %": f"{p.contribution_margin_pct:.1f}%",
                "Маржа": f"{p.margin:.1f}%", "ABC": p.abc_category,
                "Лучший МП": p.best_marketplace, "Тренд": p.demand_trend
            } for p in top])
            st.dataframe(df_top, use_container_width=True, hide_index=True)
        else:
            st.info("Нет данных для отображения топа по прибыли")

        alerts = self.notification_manager.generate_alerts(products)
        if alerts:
            st.subheader("⚠️ Важные алерты")
            for alert in alerts:
                color = {"high": "🔴", "medium": "🟡", "low": "🟢"}.get(alert["priority"], "⚪")
                st.info(f"{color} **{alert['type']}**: {alert['message']}")

    def _tab_abc_xyz(self, products):
        """Вкладка ABC/XYZ анализа"""
        st.subheader("🎯 ABC/XYZ Анализ")
        
        st.markdown("""
        <div style="background: #e7f3ff; padding: 1rem; border-radius: 10px; margin-bottom: 1rem;">
        💡 <b>ABC/XYZ анализ</b> помогает классифицировать товары по их вкладу в прибыль (ABC) 
        и стабильности спроса (XYZ). 
        <br>🔴 AX — звезды, 🟢 CZ — кандидаты на вывод
        </div>
        """, unsafe_allow_html=True)

        analyzer = InventoryAnalyzer()
        matrix = analyzer.abc_xyz_matrix(products)

        st.markdown("### 📊 Матрица ABC-XYZ")

        cols = st.columns(3)
        for i, abc in enumerate(["A", "B", "C"]):
            for j, xyz in enumerate(["X", "Y", "Z"]):
                key = f"{abc}{xyz}"
                items = matrix.get(key, [])

                if abc == "A" and xyz == "X":
                    color = "#28a745"
                    desc = "⭐ ЗВЁЗДЫ"
                elif abc == "C" and xyz == "Z":
                    color = "#dc3545"
                    desc = "🐶 СОБАКИ"
                elif abc == "A":
                    color = "#ffc107"
                    desc = "💎 Ценные"
                else:
                    color = "#17a2b8"
                    desc = "📦 Обычные"

                profit_sum = sum(p.unit_profit for p in items)

                with cols[j]:
                    st.markdown(f"""
                    <div style="background: {color}; color: white; padding: 1rem;
                                border-radius: 10px; margin: 0.5rem 0; text-align: center;">
                        <h3 style="margin: 0;">{key}</h3>
                        <p style="margin: 0.5rem 0; font-size: 0.9rem;">{desc}</p>
                        <p style="margin: 0;"><b>{len(items)}</b> товаров</p>
                        <p style="margin: 0; font-size: 0.85rem;">💰 {profit_sum:,.0f} ₽</p>
                    </div>
                    """, unsafe_allow_html=True)

        st.markdown("### 💡 Стратегии по категориям")
        strategies = {
            "AX": "⭐ Максимальные запасы, приоритетный заказ",
            "AY": "💎 Регулярный мониторинг, страховой запас",
            "AZ": "⚠️ Прогнозировать сложно, увеличьте safety stock",
            "BX": "📦 Стабильные, оптимизируйте заказы",
            "BY": "📊 Стандартное управление",
            "BZ": "🔍 Требуют внимания, пересмотрите ассортимент",
            "CX": "💰 Малая прибыль, но стабильно - автоматизация",
            "CY": "⚖️ Минимальные запасы, заказ под заказ",
            "CZ": "🗑️ Кандидаты на вывод из ассортимента"
        }

        for key, strategy in strategies.items():
            count = len(matrix.get(key, []))
            if count > 0:
                st.markdown(f"**{key}** ({count} товаров): {strategy}")

    def _tab_forecast(self, products):
        """Вкладка прогнозов"""
        st.subheader("📈 Прогноз спроса")
        
        st.markdown("""
        <div style="background: #e7f3ff; padding: 1rem; border-radius: 10px; margin-bottom: 1rem;">
        💡 <b>Прогноз спроса</b> помогает планировать закупки и избегать дефицита или избытка товаров.
        Для каждого товара показывается прогноз на 30, 60 и 90 дней.
        </div>
        """, unsafe_allow_html=True)

        top = sorted(products, key=lambda x: abs(x.demand_forecast_90), reverse=True)[:5]

        for product in top:
            with st.expander(f"📌 {product.article} - {product.name[:40]}"):
                forecast = self.processor.forecaster.forecast(product.daily_sales_history, days=90)

                col1, col2, col3, col4 = st.columns(4)
                col1.metric("📊 30 дней", f"{forecast['forecast_30']:.0f} шт")
                col2.metric("📊 60 дней", f"{forecast['forecast_60']:.0f} шт")
                col3.metric("📊 90 дней", f"{forecast['forecast_90']:.0f} шт")
                col4.metric("📈 Тренд", forecast['trend'])

                fig = go.Figure()
                fig.add_trace(go.Scatter(
                    y=forecast['history'], mode='lines+markers',
                    name='История', line=dict(color='#667eea')
                ))
                fig.add_trace(go.Scatter(
                    y=forecast['daily_forecast'], mode='lines',
                    name='Прогноз', line=dict(color='#f5576c', dash='dash')
                ))
                fig.add_trace(go.Scatter(
                    y=forecast['upper_bound'], mode='lines',
                    name='Верхняя граница', line=dict(color='rgba(245,87,108,0.2)'),
                    showlegend=False
                ))
                fig.add_trace(go.Scatter(
                    y=forecast['lower_bound'], mode='lines',
                    name='Нижняя граница', line=dict(color='rgba(245,87,108,0.2)'),
                    fill='tonexty', showlegend=False
                ))
                fig.update_layout(title='📈 Прогноз на 90 дней', height=350)
                st.plotly_chart(fig, use_container_width=True)

        st.markdown("### 📊 Сводка по трендам")
        trends = {}
        for p in products:
            t = p.demand_trend or "Неизвестно"
            trends[t] = trends.get(t, 0) + 1

        if trends:
            cols = st.columns(len(trends))
            for i, (trend, count) in enumerate(trends.items()):
                cols[i].metric(trend, f"{count:,} товаров")
        else:
            st.info("Нет данных для отображения трендов")

    def _tab_inventory(self, products):
        """Вкладка управления запасами"""
        st.subheader("📦 Управление запасами")
        
        st.markdown("""
        <div style="background: #e7f3ff; padding: 1rem; border-radius: 10px; margin-bottom: 1rem;">
        💡 <b>Управление запасами</b> показывает оптимальный размер заказа (EOQ),
        точку заказа (Reorder Point) и страховой запас (Safety Stock) для каждого товара.
        </div>
        """, unsafe_allow_html=True)

        statuses = {}
        for p in products:
            s = p.stock_status or "Не определён"
            statuses[s] = statuses.get(s, 0) + 1

        if statuses:
            cols = st.columns(len(statuses))
            for i, (status, count) in enumerate(statuses.items()):
                cols[i].metric(status, f"{count:,}")
        else:
            st.info("Нет данных о статусе запасов")

        critical = [p for p in products if "Критический" in (p.stock_status or "")]
        if critical:
            st.markdown("### 🔴 Требуют срочного заказа")
            df = pd.DataFrame([{
                "Артикул": p.article, "Название": p.name[:30],
                "Точка заказа": f"{p.reorder_point:.0f}",
                "EOQ": f"{p.eoq:.0f}",
                "Safety Stock": f"{p.safety_stock:.0f}",
                "Статус": p.stock_status
            } for p in critical[:20]])
            st.dataframe(df, use_container_width=True, hide_index=True)

        st.markdown("### 📊 EOQ (оптимальный размер заказа)")
        top_eoq = sorted(products, key=lambda x: x.eoq, reverse=True)[:10]
        if top_eoq:
            fig = go.Figure(data=[
                go.Bar(x=[p.article for p in top_eoq],
                       y=[p.eoq for p in top_eoq],
                       marker_color='#667eea')
            ])
            fig.update_layout(title='🏆 Топ-10 по EOQ', height=350, xaxis_tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Нет данных для отображения EOQ")

    def _tab_multi_mp(self, products):
        """Вкладка мульти-маркетплейс анализа"""
        st.subheader("🏪 Мульти-маркетплейс аналитика")
        
        st.markdown("""
        <div style="background: #e7f3ff; padding: 1rem; border-radius: 10px; margin-bottom: 1rem;">
        💡 <b>Мульти-маркетплейс анализ</b> помогает выбрать лучший канал продаж для каждого товара,
        учитывая комиссии, логистику и аудиторию маркетплейса.
        </div>
        """, unsafe_allow_html=True)

        mp_stats = {mp: {"profit": 0, "margin_sum": 0, "count": 0}
                   for mp in CONFIG['supported_marketplaces']}

        for p in products:
            if p.best_marketplace in mp_stats:
                mp_stats[p.best_marketplace]["profit"] += p.unit_profit
                mp_stats[p.best_marketplace]["margin_sum"] += p.margin
                mp_stats[p.best_marketplace]["count"] += 1

        st.markdown("### 📊 Какой маркетплейс лучше для ваших товаров?")
        mp_names = list(mp_stats.keys())
        mp_profits = [mp_stats[mp]["profit"] for mp in mp_names]

        if mp_profits and sum(mp_profits) > 0:
            fig = px.bar(x=mp_names, y=mp_profits,
                        title='💰 Суммарная прибыль по маркетплейсам',
                        labels={'x': 'Маркетплейс', 'y': 'Прибыль, ₽'},
                        color=mp_profits, color_continuous_scale='Viridis')
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Нет данных о прибыли по маркетплейсам")

        st.markdown("### 📋 Детальная статистика")
        df = pd.DataFrame([{
            "Маркетплейс": mp,
            "Лучших товаров": mp_stats[mp]["count"],
            "Сумма прибыли": f"{mp_stats[mp]['profit']:,.0f} ₽",
            "Ср. маржа": f"{mp_stats[mp]['margin_sum']/max(mp_stats[mp]['count'],1):.1f}%"
        } for mp in mp_names])
        st.dataframe(df, use_container_width=True, hide_index=True)

    def _tab_logistics(self, products):
        """Вкладка логистики"""
        st.subheader("🚚 Оптимизация логистики")
        
        st.markdown("""
        <div style="background: #e7f3ff; padding: 1rem; border-radius: 10px; margin-bottom: 1rem;">
        💡 <b>Оптимизация логистики</b> помогает выбрать лучшего перевозчика для каждого товара,
        учитывая стоимость, скорость и надёжность доставки.
        </div>
        """, unsafe_allow_html=True)

        providers = {}
        for p in products:
            prov = p.logistics_provider or "Не выбран"
            providers[prov] = providers.get(prov, 0) + 1

        if providers:
            st.markdown("### 📊 Рекомендуемые службы доставки")
            cols = st.columns(len(providers))
            for i, (prov, count) in enumerate(providers.items()):
                cols[i].metric(prov, f"{count:,} товаров")
        else:
            st.info("Нет данных о логистических провайдерах")

        if products:
            st.markdown("### 📋 Пример расчёта для товара")
            p = products[0]
            result = self.processor.logistics_optimizer.optimize(p)

            col1, col2, col3 = st.columns(3)
            col1.metric("💰 Самый дешёвый",
                       f"{result['cheapest']['name']}\n{result['cheapest']['cost']:.0f} ₽")
            col2.metric("⭐ Лучший баланс",
                       f"{result['best']['name']}\n{result['best']['cost']:.0f} ₽")
            col3.metric("⚡ Самый быстрый",
                       f"{result['fastest']['name']}\n{result['fastest']['days']} дн")

            df = pd.DataFrame([{
                "Служба": name,
                "Стоимость": f"{data['cost']:.0f} ₽",
                "Срок": f"{data['days']} дн",
                "Надёжность": f"{data['reliability']*100:.0f}%"
            } for name, data in result['all_providers'].items()])
            st.dataframe(df, use_container_width=True, hide_index=True)

    def _tab_pl(self, products):
        """Вкладка P&L отчета"""
        st.subheader("💰 P&L Отчёт (Прибыли и Убытки)")
        
        st.markdown("""
        <div style="background: #e7f3ff; padding: 1rem; border-radius: 10px; margin-bottom: 1rem;">
        💡 <b>P&L Отчёт</b> показывает полную финансовую картину бизнеса:
        выручку, себестоимость, все расходы и прибыль.
        </div>
        """, unsafe_allow_html=True)

        pl = self.pl_reporter.generate(products)

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("💵 Выручка", f"{pl['revenue']:,.0f} ₽")
        col2.metric("💰 Валовая прибыль", f"{pl['gross_profit']:,.0f} ₽",
                   f"{pl['gross_margin']:.1f}%")
        col3.metric("📊 Операционная прибыль", f"{pl['operating_profit']:,.0f} ₽",
                   f"{pl['operating_margin']:.1f}%")
        col4.metric("📦 Средний чек", f"{pl['avg_check']:,.0f} ₽")

        col1, col2 = st.columns(2)
        with col1:
            if pl['expenses'] and sum(pl['expenses'].values()) > 0:
                fig = px.pie(values=list(pl['expenses'].values()),
                            names=list(pl['expenses'].keys()),
                            title='📊 Структура расходов',
                            hole=0.4)
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Нет данных о расходах для отображения")

        with col2:
            fig = go.Figure(go.Waterfall(
                name="P&L",
                orientation="v",
                x=["Выручка", "Себестоимость", "Комиссия", "Логистика",
                   "Реклама", "Возвраты", "Прибыль"],
                y=[pl['revenue'], -pl['cogs'], -pl['expenses']['Комиссия МП'],
                   -pl['expenses']['Логистика'], -pl['expenses']['Реклама'],
                   -pl['expenses']['Возвраты'], pl['operating_profit']],
                measure=["absolute", "relative", "relative", "relative",
                        "relative", "relative", "total"],
                connector={"line": {"color": "rgb(63, 63, 63)"}}
            ))
            fig.update_layout(title='💰 Waterfall P&L', height=400, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("### 📋 Детализация")
        
        # Исправлено: проверка на деление на ноль
        revenue = pl['revenue'] if pl['revenue'] > 0 else 1
        
        df = pd.DataFrame([{
            "Показатель": "Выручка", 
            "Сумма": f"{pl['revenue']:,.0f} ₽", 
            "%": "100%"
        }, {
            "Показатель": "Себестоимость", 
            "Сумма": f"{pl['cogs']:,.0f} ₽",
            "%": f"{pl['cogs']/revenue*100:.1f}%"
        }, {
            "Показатель": "Валовая прибыль", 
            "Сумма": f"{pl['gross_profit']:,.0f} ₽",
            "%": f"{pl['gross_margin']:.1f}%"
        }, {
            "Показатель": "Операционная прибыль", 
            "Сумма": f"{pl['operating_profit']:,.0f} ₽",
            "%": f"{pl['operating_margin']:.1f}%"
        }])
        st.dataframe(df, use_container_width=True, hide_index=True)

    def _tab_ab_test(self, products):
        """Вкладка A/B тестирования"""
        st.subheader("🧪 A/B Тестирование цен")
        
        st.markdown("""
        <div style="background: #e7f3ff; padding: 1rem; border-radius: 10px; margin-bottom: 1rem;">
        💡 <b>A/B Тестирование</b> помогает найти оптимальную цену для товара,
        сравнивая текущую цену, +10% и -5%.
        </div>
        """, unsafe_allow_html=True)

        if st.button("🚀 Запустить A/B/C тест", type="primary"):
            with st.spinner("🧪 Симулируем тест..."):
                result = self.ab_tester.run_test(products[:20], test_days=14)
                st.session_state.ab_test_result = result
                st.session_state.stats['ab_test_done'] = True

        if st.session_state.get('ab_test_result'):
            result = st.session_state.ab_test_result

            st.markdown(f"""
            <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        color: white; padding: 1.5rem; border-radius: 10px; margin: 1rem 0;">
                <h3>🎯 Рекомендуемая стратегия: {result['recommended_strategy']}</h3>
                <p>Средний uplift: <b>{result['total_uplift']:.1f}%</b></p>
                <p>Статистическая значимость: {'✅ Да' if result['statistical_significance'] else '⚠️ Недостаточно данных'}</p>
            </div>
            """, unsafe_allow_html=True)

            for test in result['tests'][:5]:
                with st.expander(f"📌 {test['article']} - {test['name']}"):
                    cols = st.columns(3)
                    for i, (group, data) in enumerate(test['groups'].items()):
                        with cols[i]:
                            color = "#28a745" if group == test['best_group'] else "#6c757d"
                            st.markdown(f"""
                            <div style="background: {color}; color: white; padding: 1rem;
                                        border-radius: 8px; text-align: center;">
                                <h4>{group}</h4>
                                <p>Цена: {data['price']:,.0f} ₽</p>
                                <p>Продаж: {data['total_sales']}</p>
                                <p><b>Прибыль: {data['profit']:,.0f} ₽</b></p>
                            </div>
                            """, unsafe_allow_html=True)

    def _render_welcome(self):
        """Приветственный экран"""
        st.markdown("""
        <div style="text-align: center; padding: 40px 0;">
            <h2 style="font-size: 2rem; color: #1a1a2e;">
                🚗 Платформа автозапчастей v14.0
            </h2>
            <p style="color: #6c757d; font-size: 1.1rem;">
                💎 УПОР НА ЮНИТ-ЭКОНОМИКУ: Contribution Margin • LTV/CAC • Break-Even • Чувствительность
            </p>

            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
                        gap: 15px; max-width: 1000px; margin: 30px auto;">
                <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                            padding: 20px; border-radius: 10px; color: white;">
                    <h3>💎 Юнит-экономика</h3>
                    <p style="font-size: 12px;">CM, LTV, CAC, ROAS</p>
                </div>
                <div style="background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
                            padding: 20px; border-radius: 10px; color: white;">
                    <h3>📊 Чувствительность</h3>
                    <p style="font-size: 12px;">Что-if анализ</p>
                </div>
                <div style="background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
                            padding: 20px; border-radius: 10px; color: white;">
                    <h3>🎯 ABC/XYZ</h3>
                    <p style="font-size: 12px;">9 категорий</p>
                </div>
                <div style="background: linear-gradient(135deg, #43e97b 0%, #38f9d7 100%);
                            padding: 20px; border-radius: 10px; color: white;">
                    <h3>📈 Прогнозы</h3>
                    <p style="font-size: 12px;">30/60/90 дней</p>
                </div>
                <div style="background: linear-gradient(135deg, #fa709a 0%, #fee140 100%);
                            padding: 20px; border-radius: 10px; color: white;">
                    <h3>📦 Запасы</h3>
                    <p style="font-size: 12px;">EOQ + Reorder</p>
                </div>
                <div style="background: linear-gradient(135deg, #30cfd0 0%, #330867 100%);
                            padding: 20px; border-radius: 10px; color: white;">
                    <h3>💰 P&L</h3>
                    <p style="font-size: 12px;">Полный отчёт</p>
                </div>
                <div style="background: linear-gradient(135deg, #a8edea 0%, #fed6e3 100%);
                            padding: 20px; border-radius: 10px; color: #333;">
                    <h3>🧪 A/B Тесты</h3>
                    <p style="font-size: 12px;">Ценовые стратегии</p>
                </div>
            </div>

            <div style="background: #e7f3ff; padding: 20px; border-radius: 10px;
                        max-width: 700px; margin: 20px auto; text-align: left;">
                <p><b>🎁 Попробуйте демо-данные!</b></p>
                <p style="font-size: 0.9rem;">Нажмите "🔄 Демо-данные" чтобы увидеть ВСЕ функции в действии</p>
                <p style="font-size: 0.9rem;">💎 Главная вкладка — <b>Юнит-экономика</b> с детальным анализом</p>
            </div>
        </div>
        """, unsafe_allow_html=True)


# --------------------------------------------
# ЗАПУСК
# --------------------------------------------
if __name__ == "__main__":
    app = UltimateAutoApp()
    app.run()
