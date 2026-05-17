# Анализ основной аудитории

!apt-get update -q
!apt-get install -y -q chromium-browser chromium-chromedriver
!pip install -q selenium==4.18.1
!pip install -q beautifulsoup4 requests openpyxl
!pip install -q pandas numpy matplotlib plotly
!pip install -q dash dash-bootstrap-components

import io
import time
import warnings
import re

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import plotly.express as px
import plotly.graph_objects as go
import requests
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

from dash import Dash, dcc, html, Input, Output
import dash_bootstrap_components as dbc

warnings.filterwarnings("ignore")

class MSPSeleniumParser:
    BASE_URL = "https://rmsp.nalog.ru/statistics.html"

    FEDERAL_DISTRICTS = {
        "Все округа (РФ)":   {"fo": "0", "level": "0"},
        "Центральный":       {"fo": "1", "level": "1"},
        "Северо-Западный":   {"fo": "2", "level": "1"},
        "Южный":             {"fo": "3", "level": "1"},
        "Северо-Кавказский": {"fo": "4", "level": "1"},
        "Приволжский":       {"fo": "5", "level": "1"},
        "Уральский":         {"fo": "6", "level": "1"},
        "Сибирский":         {"fo": "7", "level": "1"},
        "Дальневосточный":   {"fo": "8", "level": "1"},
    }

    def __init__(self, headless=True):
        self.headless = headless
        self.driver = None
        self.collected_links = []

    def _init_driver(self):
        options = Options()
        if self.headless:
            options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("user-agent=Mozilla/5.0")
        self.driver = webdriver.Chrome(
            service=Service("/usr/bin/chromedriver"), options=options
        )
        print("Браузер запущен")

    def _open_page(self):
        self.driver.get(self.BASE_URL)
        try:
            WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.TAG_NAME, "table"))
            )
            print("Страница загружена")
        except TimeoutException:
            print("Таблица не найдена за 15 сек")

    def _get_stat_date(self):
        soup = BeautifulSoup(self.driver.page_source, "html.parser")
        for tag in soup.find_all(["p", "span", "div", "h3", "b"]):
            match = re.search(r"\d{2}\.\d{2}\.\d{4}", tag.get_text(strip=True))
            if match:
                print(f"Дата статистики: {match.group()}")
                return match.group()
        return "10.04.2025"

    def _build_xlsx_urls(self, stat_date):
        links = []
        for name, params in self.FEDERAL_DISTRICTS.items():
            url = (
                f"https://rmsp.nalog.ru/statistics.xlsx"
                f"?statDate={stat_date}&level={params['level']}&fo={params['fo']}&ssrf=0"
            )
            links.append({"district": name, "url": url})
        return links

    def collect_links(self):
        self._init_driver()
        try:
            self._open_page()
            stat_date = self._get_stat_date()
            self.collected_links = self._build_xlsx_urls(stat_date)
            print(f"Ссылок собрано: {len(self.collected_links)}")
        finally:
            self.driver.quit()
            print("Браузер закрыт")
        return self.collected_links

        class MSPStaticParser:
    HEADERS = {
        "User-Agent": "Mozilla/5.0",
        "Referer": "https://rmsp.nalog.ru/statistics.html",
    }

    def __init__(self, links):
        self.links = links
        self.raw_frames = {}

    def parse_html_meta(self):
        try:
            resp = requests.get("https://rmsp.nalog.ru/statistics.html",
                                headers=self.HEADERS, timeout=15)
            soup = BeautifulSoup(resp.text, "html.parser")
            print("Заголовок:", soup.title.get_text(strip=True) if soup.title else "—")
            print("Таблиц на странице:", len(soup.find_all("table")))
        except Exception as e:
            print(f"HTML не получен: {e}")

    def _download_xlsx(self, url, district):
        try:
            resp = requests.get(url, headers=self.HEADERS, timeout=30)
            resp.raise_for_status()
            df = pd.read_excel(io.BytesIO(resp.content), header=None)
            print(f"  {district}: {len(df)} строк")
            return df
        except Exception as e:
            print(f"  {district}: ошибка — {e}")
            return None

    def download_all(self):
        print("Скачиваем xlsx...")
        for item in self.links:
            df = self._download_xlsx(item["url"], item["district"])
            if df is not None:
                self.raw_frames[item["district"]] = df
            time.sleep(0.5)
        print(f"Скачано: {len(self.raw_frames)}")
        return self.raw_frames

        class MSPDataProcessor:
    CATEGORY_KEYWORDS = {
        "Микропредприятия":    ["микро"],
        "Малые предприятия":   ["малые", "малой"],
        "Средние предприятия": ["средние", "средней"],
    }
    TYPE_KEYWORDS = {
        "Юридические лица": ["юридических лиц", "юридические лица"],
        "ИП":               ["индивидуальных предпринимателей", "ип"],
    }

    def __init__(self, raw_frames):
        self.raw_frames = raw_frames
        self.summary_df = pd.DataFrame()

    @staticmethod
    def _find_value(df, keywords):
        for _, row in df.iterrows():
            for cell in row:
                if isinstance(cell, str) and any(k in cell.lower() for k in keywords):
                    for val in row:
                        if isinstance(val, (int, float)) and not pd.isna(val) and val > 0:
                            return float(val)
        return 0.0

    def _process_one(self, district, df):
        row = {"district": district}
        for name, kw in self.CATEGORY_KEYWORDS.items():
            row[name] = self._find_value(df, kw)
        for name, kw in self.TYPE_KEYWORDS.items():
            row[name] = self._find_value(df, kw)
        row["Итого МСП"] = row["Микропредприятия"] + row["Малые предприятия"] + row["Средние предприятия"]
        return row

    def process_all(self):
        rows = [self._process_one(d, df) for d, df in self.raw_frames.items()]
        self.summary_df = pd.DataFrame(rows)
        print(self.summary_df.to_string(index=False))
        return self.summary_df

        class MSPDemoData:
    DEMO = {
        "Все округа (РФ)":   {"Микропредприятия": 5_820_000, "Малые предприятия": 190_000, "Средние предприятия": 18_000, "Итого МСП": 6_028_000, "Юридические лица": 1_850_000, "ИП": 4_178_000},
        "Центральный":       {"Микропредприятия": 1_750_000, "Малые предприятия":  58_000, "Средние предприятия":  5_400, "Итого МСП": 1_813_400, "Юридические лица":   620_000, "ИП": 1_193_400},
        "Приволжский":       {"Микропредприятия":   960_000, "Малые предприятия":  32_000, "Средние предприятия":  3_100, "Итого МСП":   995_100, "Юридические лица":   280_000, "ИП":   715_100},
        "Северо-Западный":   {"Микропредприятия":   680_000, "Малые предприятия":  22_000, "Средние предприятия":  2_100, "Итого МСП":   704_100, "Юридические лица":   230_000, "ИП":   474_100},
        "Сибирский":         {"Микропредприятия":   540_000, "Малые предприятия":  17_000, "Средние предприятия":  1_600, "Итого МСП":   558_600, "Юридические лица":   145_000, "ИП":   413_600},
        "Южный":             {"Микропредприятия":   490_000, "Малые предприятия":  15_000, "Средние предприятия":  1_400, "Итого МСП":   506_400, "Юридические лица":   125_000, "ИП":   381_400},
        "Уральский":         {"Микропредприятия":   420_000, "Малые предприятия":  14_000, "Средние предприятия":  1_350, "Итого МСП":   435_350, "Юридические лица":   120_000, "ИП":   315_350},
        "Дальневосточный":   {"Микропредприятия":   290_000, "Малые предприятия":   9_500, "Средние предприятия":    900, "Итого МСП":   300_400, "Юридические лица":    75_000, "ИП":   225_400},
        "Северо-Кавказский": {"Микропредприятия":   210_000, "Малые предприятия":   6_500, "Средние предприятия":    620, "Итого МСП":   217_120, "Юридические лица":    45_000, "ИП":   172_120},
    }

    @classmethod
    def get_dataframe(cls):
        rows = [{"district": k, **v} for k, v in cls.DEMO.items()]
        print("Используются демо-данные (ФНС 2024)")
        return pd.DataFrame(rows)

        class MSPVisualizer:
    PALETTE = ["#2563EB", "#16A34A", "#DC2626", "#D97706",
               "#7C3AED", "#0891B2", "#BE185D", "#65A30D"]

    def __init__(self, summary_df):
        self.df = summary_df.copy()
        self.df_regions = self.df[self.df["district"] != "Все округа (РФ)"].copy()
        self.df_total   = self.df[self.df["district"] == "Все округа (РФ)"].copy()

    def plot_all(self):
        self._bar_by_district()
        self._pie_categories()
        self._bar_ul_ip()
        self._stacked_pct()

    def _bar_by_district(self):
        df = self.df_regions.sort_values("Итого МСП")
        fig, ax = plt.subplots(figsize=(12, 6))
        bars = ax.barh(df["district"], df["Итого МСП"] / 1000,
                       color=self.PALETTE[:len(df)], edgecolor="white")
        for bar in bars:
            w = bar.get_width()
            ax.text(w + 5, bar.get_y() + bar.get_height() / 2,
                    f"{w:,.0f} тыс.", va="center", fontsize=9)
        ax.set_xlabel("тыс. субъектов")
        ax.set_title("МСП по федеральным округам")
        ax.spines[["top", "right"]].set_visible(False)
        plt.tight_layout()
        plt.show()

    def _pie_categories(self):
        total_row = self.df_total.iloc[0] if not self.df_total.empty \
                    else self.df_regions.sum(numeric_only=True)
        cats = ["Микропредприятия", "Малые предприятия", "Средние предприятия"]
        fig, ax = plt.subplots(figsize=(7, 7))
        ax.pie([total_row[c] for c in cats],
               labels=cats, autopct="%1.1f%%",
               colors=self.PALETTE[:3],
               wedgeprops={"edgecolor": "white", "linewidth": 2})
        ax.set_title("Структура МСП по категориям")
        plt.tight_layout()
        plt.show()

    def _bar_ul_ip(self):
        df = self.df_regions.sort_values("Итого МСП", ascending=False)
        x = np.arange(len(df))
        fig, ax = plt.subplots(figsize=(12, 6))
        ul = df["Юридические лица"].values / 1000
        ip = df["ИП"].values / 1000
        ax.bar(x, ul, 0.5, label="ЮЛ", color="#6366F1")
        ax.bar(x, ip, 0.5, bottom=ul, label="ИП", color="#EC4899")
        ax.set_xticks(x)
        ax.set_xticklabels(df["district"], rotation=25, ha="right")
        ax.set_ylabel("тыс.")
        ax.set_title("ЮЛ и ИП по округам")
        ax.legend()
        ax.spines[["top", "right"]].set_visible(False)
        plt.tight_layout()
        plt.show()

    def _stacked_pct(self):
        df = self.df_regions.copy()
        cats = ["Микропредприятия", "Малые предприятия", "Средние предприятия"]
        total = df[cats].sum(axis=1)
        x = np.arange(len(df))
        fig, ax = plt.subplots(figsize=(12, 6))
        bottom = np.zeros(len(df))
        for i, c in enumerate(cats):
            vals = df[c] / total * 100
            ax.bar(x, vals, 0.6, bottom=bottom, label=c, color=self.PALETTE[i])
            for j, (v, b) in enumerate(zip(vals, bottom)):
                if v > 3:
                    ax.text(x[j], b + v / 2, f"{v:.1f}%",
                            ha="center", va="center", fontsize=8,
                            color="white", fontweight="bold")
            bottom += vals
        ax.set_xticks(x)
        ax.set_xticklabels(df["district"], rotation=25, ha="right")
        ax.set_ylabel("%")
        ax.set_ylim(0, 100)
        ax.set_title("Доля категорий МСП по округам")
        ax.legend()
        ax.spines[["top", "right"]].set_visible(False)
        plt.tight_layout()
        plt.show()

        class MSPDashboard:
    COLORS = {
        "primary": "#1E40AF", "micro": "#3B82F6",
        "small":   "#10B981", "medium": "#F59E0B",
        "ul":      "#6366F1", "ip":    "#EC4899",
        "bg":      "#F8FAFC", "card":  "#FFFFFF",
        "text":    "#1E293B", "muted": "#64748B",
    }

    def __init__(self, summary_df):
        self.df = summary_df.copy()
        self.df_regions = self.df[self.df["district"] != "Все округа (РФ)"].copy()
        self.app = Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])
        self._build_layout()
        self._register_callbacks()

    def _kpi_card(self, title, value, color):
        return dbc.Card(dbc.CardBody([
            html.H6(title, style={"color": self.COLORS["muted"], "fontSize": "0.8rem"}),
            html.H3(value, style={"color": color, "fontWeight": "800"}),
        ]), style={"borderRadius": "12px", "border": "none",
                   "boxShadow": "0 4px 16px rgba(0,0,0,0.06)",
                   "textAlign": "center"})

    def _build_layout(self):
        df = self.df_regions
        districts_opts = [{"label": "Все", "value": "ALL"}] + \
                         [{"label": d, "value": d} for d in df["district"].unique()]
        cats_opts = [
            {"label": " Микро",   "value": "Микропредприятия"},
            {"label": " Малые",   "value": "Малые предприятия"},
            {"label": " Средние", "value": "Средние предприятия"},
        ]

        self.app.layout = dbc.Container([
            html.H4("Субъекты МСП России", className="mt-4 mb-3",
                    style={"color": self.COLORS["primary"], "fontWeight": "800"}),

            dbc.Row([
                dbc.Col(self._kpi_card("Всего МСП",
                    f'{int(df["Итого МСП"].sum()):,}'.replace(",", " "),
                    self.COLORS["primary"]), md=3),
                dbc.Col(self._kpi_card("ИП",
                    f'{int(df["ИП"].sum()):,}'.replace(",", " "),
                    self.COLORS["ip"]), md=3),
                dbc.Col(self._kpi_card("Юр. лиц",
                    f'{int(df["Юридические лица"].sum()):,}'.replace(",", " "),
                    self.COLORS["ul"]), md=3),
                dbc.Col(self._kpi_card("Микропредприятий",
                    f'{int(df["Микропредприятия"].sum()):,}'.replace(",", " "),
                    self.COLORS["micro"]), md=3),
            ], className="g-3 mb-4"),

            dbc.Row([
                dbc.Col([
                    html.Label("Округ"),
                    dcc.Dropdown(id="dd", options=districts_opts,
                                 value="ALL", clearable=False),
                ], md=4),
                dbc.Col([
                    html.Label("Категории"),
                    dcc.Checklist(id="cats", options=cats_opts,
                                  value=["Микропредприятия",
                                         "Малые предприятия",
                                         "Средние предприятия"],
                                  inline=True,
                                  inputStyle={"marginRight": "4px", "marginLeft": "12px"}),
                ], md=8),
            ], className="mb-4"),

            dbc.Row([
                dbc.Col(dcc.Graph(id="g1"), md=8),
                dbc.Col(dcc.Graph(id="g2"), md=4),
            ], className="mb-3"),
            dbc.Row([
                dbc.Col(dcc.Graph(id="g3"), md=6),
                dbc.Col(dcc.Graph(id="g4"), md=6),
            ]),
        ], fluid=True, style={"background": self.COLORS["bg"], "minHeight": "100vh"})

    def _register_callbacks(self):
        @self.app.callback(
            Output("g1", "figure"), Output("g2", "figure"),
            Output("g3", "figure"), Output("g4", "figure"),
            Input("dd", "value"), Input("cats", "value"),
        )
        def update(district, cats):
            df = self.df_regions if district == "ALL" \
                 else self.df_regions[self.df_regions["district"] == district]
            cats = cats or ["Микропредприятия", "Малые предприятия", "Средние предприятия"]
            colors = {"Микропредприятия": self.COLORS["micro"],
                      "Малые предприятия": self.COLORS["small"],
                      "Средние предприятия": self.COLORS["medium"]}

            g1 = px.bar(df.sort_values("Итого МСП"), x="Итого МСП", y="district",
                        orientation="h", color="district",
                        template="plotly_white", title="Всего МСП")
            g1.update_layout(showlegend=False, margin=dict(l=10,r=10,t=40,b=10))

            g2 = go.Figure(go.Pie(
                labels=cats, values=[df[c].sum() for c in cats], hole=0.4,
                marker_colors=[colors[c] for c in cats]))
            g2.update_layout(title="Категории", margin=dict(l=10,r=10,t=40,b=10))

            g3 = go.Figure()
            g3.add_trace(go.Bar(x=df["district"], y=df["Юридические лица"],
                                name="ЮЛ", marker_color=self.COLORS["ul"]))
            g3.add_trace(go.Bar(x=df["district"], y=df["ИП"],
                                name="ИП", marker_color=self.COLORS["ip"]))
            g3.update_layout(barmode="group", template="plotly_white",
                             title="ЮЛ и ИП", xaxis_tickangle=-30,
                             margin=dict(l=10,r=10,t=40,b=80))

            g4 = go.Figure()
            total = df[cats].sum(axis=1)
            for c in cats:
                g4.add_trace(go.Bar(x=df["district"],
                                    y=(df[c] / total * 100).round(1),
                                    name=c, marker_color=colors[c]))
            g4.update_layout(barmode="stack", template="plotly_white",
                             title="Доля %", yaxis_ticksuffix="%",
                             xaxis_tickangle=-30,
                             margin=dict(l=10,r=10,t=40,b=80))
            return g1, g2, g3, g4

    def run(self, port=8050):
        self.app.run(port=port, debug=False, jupyter_mode="inline")

        def main():
    summary_df = None
    try:
        parser = MSPSeleniumParser(headless=True)
        links = parser.collect_links()

        static = MSPStaticParser(links)
        static.parse_html_meta()
        raw_frames = static.download_all()

        if raw_frames:
            summary_df = MSPDataProcessor(raw_frames).process_all()
        else:
            raise ValueError("Нет данных")

    except Exception as e:
        print(f"Парсинг не удался: {e}\nЗагружаем демо-данные...")
        summary_df = MSPDemoData.get_dataframe()

    MSPVisualizer(summary_df).plot_all()
    MSPDashboard(summary_df).run()

main()
