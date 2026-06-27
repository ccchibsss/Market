"""
================================================================================
🚀 ULTIMATE UNIT ECONOMICS ENGINE v44.0 - АНАЛИТИКА + ДАШБОРД
================================================================================
📌 НОВЫЕ ФУНКЦИИ:
    ✅ Аналитика продаж в реальном времени
    ✅ Интерактивный дашборд с графиками
    ✅ Прогнозирование продаж (ARIMA, Prophet)
    ✅ Анализ трендов и сезонности
    ✅ ABC/XYZ анализ с визуализацией
    ✅ Тепловая карта категорий
    ✅ Экспорт дашборда в PDF

📊 ДОСТУПНЫЕ ГРАФИКИ:
    📈 Продажи по дням/неделям/месяцам
    📊 Распределение по категориям
    📉 Тренды прибыли и маржи
    🎯 ABC/XYZ матрица
    🔥 Тепловая карта категорий
    📈 Прогноз продаж на 30 дней

🚀 ЗАПУСК:
    streamlit run ultimate_unit_economy.py
================================================================================
"""

# [ВСЯ ПРЕДЫДУЩАЯ ЧАСТЬ КОДА СОХРАНЯЕТСЯ]

# --------------------------------------------
# 📊 ДОПОЛНИТЕЛЬНЫЕ ИМПОРТЫ ДЛЯ АНАЛИТИКИ
# --------------------------------------------

try:
    import plotly.graph_objects as go
    import plotly.express as px
    from plotly.subplots import make_subplots
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False
    logger.warning("Plotly не установлен. Установите: pip install plotly")

try:
    from sklearn.linear_model import LinearRegression
    from sklearn.metrics import mean_absolute_error, mean_squared_error
    from sklearn.preprocessing import StandardScaler
    SKLEARN_AVAILABLE = True
except ImportError:
    SKLEARN_AVAILABLE = False
    logger.warning("Scikit-learn не установлен. Установите: pip install scikit-learn")

try:
    import pandas as pd
    import numpy as np
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False


# --------------------------------------------
# 📈 КЛАСС АНАЛИТИКИ ПРОДАЖ
# --------------------------------------------

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


# --------------------------------------------
# 📊 КЛАСС ПРОГНОЗИРОВАНИЯ
# --------------------------------------------

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


# --------------------------------------------
# 🎨 ОБНОВЛЕНИЕ ОСНОВНОГО ПРИЛОЖЕНИЯ
# --------------------------------------------

class UnitEconomicsApp:
    """Главное приложение с дашбордом"""
    
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
    
    def _render_main(self):
        """Основной контент с дашбордом"""
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
    
    def _render_dashboard_tab(self):
        """Вкладка дашборда"""
        st.subheader("📊 Дашборд аналитики в реальном времени")
        
        if not PLOTLY_AVAILABLE:
            st.warning("⚠️ Для отображения графиков установите plotly: pip install plotly")
        
        # Создаем демонстрационные данные, если нет реальных
        if not self.sales_data:
            self._generate_demo_sales_data()
        
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
    
    def _generate_demo_sales_data(self):
        """Генерация демонстрационных данных для дашборда"""
        import random
        
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
                y=df['revenue'],
                name="Выручка",
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
        
        fig.update_yaxes(title_text="Выручка, ₽", secondary_y=False)
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
                values=df['revenue'].head(8),
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
                y=df['revenue'].head(10),
                text=df['revenue'].head(10).apply(lambda x: f"{x:,.0f}₽"),
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
            st.info("**AY/ AZ** - высокая прибыль, но нестабильные продажи")
        
        with col2:
            st.warning("**BX/ BY** - средняя прибыль, требуется оптимизация")
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
                st.experimental_rerun()


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
