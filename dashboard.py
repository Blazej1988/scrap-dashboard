#!/usr/bin/env python3
"""
SCRAP ANALYSIS DASHBOARD
Interactive web dashboard dla analizy złomów - trendy, filtry, wykresy
"""

import streamlit as st
import pandas as pd
import openpyxl
from pathlib import Path
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime

# Konfiguracja strony
st.set_page_config(
    page_title="SCRAP ANALYSIS DASHBOARD",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS dla lepszego wyglądu
st.markdown("""
    <style>
        .metric-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 10px;
            text-align: center;
        }
        .section-header {
            font-size: 24px;
            font-weight: bold;
            color: #1f4e78;
            margin-top: 30px;
            margin-bottom: 10px;
            border-bottom: 3px solid #1f4e78;
            padding-bottom: 10px;
        }
    </style>
""", unsafe_allow_html=True)

# ==================== LOAD DATA ====================
@st.cache_data
def load_excel_data():
    """Załaduj dane z najnowszego raportu Excel"""
    output_dir = Path('OUTPUT')
    if not output_dir.exists():
        st.error("❌ Folder OUTPUT nie znaleziony!")
        return None, None

    files = sorted(output_dir.glob('SCRAP_ANALYSIS_*.xlsx'))
    if not files:
        st.error("❌ Brak raportów Excel!")
        return None, None

    latest_file = files[-1]
    st.sidebar.info(f"📄 Plik: {latest_file.name}")

    try:
        # Czytaj Data sheet (monthly aggregation)
        data_monthly = pd.read_excel(latest_file, sheet_name='Data', engine='openpyxl')

        # Czytaj DataWeekly sheet (weekly aggregation)
        data_weekly = pd.read_excel(latest_file, sheet_name='DataWeekly', engine='openpyxl')

        return data_monthly, data_weekly
    except Exception as e:
        st.error(f"❌ Błąd czytania Excel: {str(e)}")
        return None, None

# Załaduj dane
data_monthly, data_weekly = load_excel_data()

if data_monthly is None or data_weekly is None:
    st.stop()

# ==================== SIDEBAR FILTERS ====================
st.sidebar.markdown("## 🔍 FILTRY")

# Typ analizy
analysis_type = st.sidebar.radio(
    "📊 Analiza:",
    options=["📅 Miesiące", "📆 Tygodnie"],
    index=0,
    help="Wybierz czy chcesz analizę po miesiącach czy tygodniach"
)

# Projekt
project_filter = st.sidebar.multiselect(
    "🏭 Projekt:",
    options=['U-PJT', 'B-PJT'],
    default=['U-PJT', 'B-PJT'],
    help="Wybierz które projekty chcesz zobaczyć"
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 📦 Komponenty")
st.sidebar.markdown("Wszystkie 3 komponenty są zawsze pokazane:")
st.sidebar.markdown("- 🔹 **FRAME** (Obudowy)")
st.sidebar.markdown("- 🔹 **FRONT** (Fronty)")
st.sidebar.markdown("- 🔹 **WRAPPER** (Wrapery)")

# ==================== DATA PREPARATION ====================
def prepare_data_for_analysis(data, analysis_type, project_filter):
    """Przygotuj dane na podstawie wybranego typu analizy"""

    # Filtruj po projektach
    filtered_data = data[data['Project'].isin(project_filter)].copy()

    if analysis_type == "📅 Miesiące":
        # Agreguj po miesiącach
        agg_data = filtered_data.groupby(['Month', 'Component', 'Project']).agg({
            'Plan_Qty': 'sum',
            'Scrap_Qty': 'sum'
        }).reset_index()

        # Oblicz ratio
        agg_data['Ratio %'] = (agg_data['Scrap_Qty'] / agg_data['Plan_Qty'] * 100).round(2)
        agg_data['Ratio %'] = agg_data['Ratio %'].fillna(0)

        # Sortuj po miesiącach
        month_order = [f'M{i}' for i in range(1, 10)]
        agg_data['Month'] = pd.Categorical(agg_data['Month'], categories=month_order, ordered=True)
        agg_data = agg_data.sort_values('Month')

        period_col = 'Month'
    else:
        # Agreguj po tygodniach
        agg_data = filtered_data.groupby(['Month', 'Week', 'Component', 'Project']).agg({
            'Plan_Qty': 'sum',
            'Scrap_Qty': 'sum'
        }).reset_index()

        agg_data['Ratio %'] = (agg_data['Scrap_Qty'] / agg_data['Plan_Qty'] * 100).round(2)
        agg_data['Ratio %'] = agg_data['Ratio %'].fillna(0)

        # Sortuj po tygodniach (numerycznie, nie alfabetycznie!)
        agg_data['Week_Num'] = agg_data['Week'].str.extract('(\d+)').astype(int)
        agg_data = agg_data.sort_values(['Month', 'Week_Num'])
        agg_data = agg_data.drop('Week_Num', axis=1)
        period_col = 'Week'

    return agg_data, period_col

# ==================== MAIN CONTENT ====================
st.markdown("# 📊 SCRAP ANALYSIS DASHBOARD")
st.markdown("Analiza trendów złomów - rzeczywisty czas")

# Przygotuj dane
agg_data, period_col = prepare_data_for_analysis(data_monthly if analysis_type == "📅 Miesiące" else data_weekly, analysis_type, project_filter)

# ==================== WYŚWIETL KOMPONENTY ====================
components = ['Frame', 'Front', 'Wrapper']

for component in components:
    # ===== NAGŁÓWEK KOMPONENTU =====
    component_names = {
        'Frame': '🔹 FRAME (Obudowy)',
        'Front': '🔹 FRONT (Fronty)',
        'Wrapper': '🔹 WRAPPER (Wrapery)'
    }

    st.markdown(f"<div class='section-header'>{component_names.get(component, component)}</div>", unsafe_allow_html=True)

    # Filtruj dane dla komponentu
    comp_data = agg_data[agg_data['Component'] == component].copy()

    if len(comp_data) == 0:
        st.info(f"❌ Brak danych dla {component}")
        continue

    # ===== TABELA Z DANYMI =====
    st.markdown("#### 📊 Tabela danych")

    # Przygotuj tabelę do wyświetlenia
    table_data = comp_data[[period_col, 'Project', 'Plan_Qty', 'Scrap_Qty', 'Ratio %']].copy()
    table_data.columns = [period_col, 'Projekt', 'Plan Qty', 'Scrap Qty', 'Ratio %']
    table_data['Plan Qty'] = table_data['Plan Qty'].astype(int)
    table_data['Scrap Qty'] = table_data['Scrap Qty'].astype(int)

    # Pivot tabela - pokaż U-PJT i B-PJT obok siebie
    pivot_data = table_data.pivot_table(
        index=period_col,
        columns='Projekt',
        values=['Plan Qty', 'Scrap Qty', 'Ratio %'],
        aggfunc='first'
    )

    st.dataframe(pivot_data, use_container_width=True)

    # ===== WYKRESY =====
    st.markdown("#### 📈 Wykresy trendów")

    col1, col2 = st.columns(2)

    # Wykres 1: Plan vs Scrap
    with col1:
        fig1 = go.Figure()

        for project in comp_data['Project'].unique():
            proj_data = comp_data[comp_data['Project'] == project].sort_values(period_col)

            fig1.add_trace(go.Scatter(
                x=proj_data[period_col],
                y=proj_data['Scrap_Qty'],
                mode='lines+markers',
                name=f'{project} - Scrap',
                line=dict(width=3),
                marker=dict(size=8)
            ))

        fig1.update_layout(
            title=f"Trend Złomów - {component}",
            xaxis_title=period_col,
            yaxis_title="Scrap Qty",
            hovermode='x unified',
            height=400,
            template='plotly_white'
        )

        st.plotly_chart(fig1, use_container_width=True)

    # Wykres 2: Ratio % (jakość)
    with col2:
        fig2 = go.Figure()

        for project in comp_data['Project'].unique():
            proj_data = comp_data[comp_data['Project'] == project].sort_values(period_col)

            fig2.add_trace(go.Scatter(
                x=proj_data[period_col],
                y=proj_data['Ratio %'],
                mode='lines+markers',
                name=f'{project} - Ratio %',
                line=dict(width=3),
                marker=dict(size=8)
            ))

        fig2.update_layout(
            title=f"Ratio % Złomów - {component}",
            xaxis_title=period_col,
            yaxis_title="Ratio %",
            hovermode='x unified',
            height=400,
            template='plotly_white'
        )

        st.plotly_chart(fig2, use_container_width=True)

    # Wykres 3: Porównanie Plan vs Scrap (słupki)
    fig3 = go.Figure()

    for project in comp_data['Project'].unique():
        proj_data = comp_data[comp_data['Project'] == project].sort_values(period_col)

        fig3.add_trace(go.Bar(
            x=proj_data[period_col],
            y=proj_data['Plan_Qty'],
            name=f'{project} - Plan',
            marker=dict(opacity=0.7)
        ))

        fig3.add_trace(go.Bar(
            x=proj_data[period_col],
            y=proj_data['Scrap_Qty'],
            name=f'{project} - Scrap',
            marker=dict(opacity=0.9)
        ))

    fig3.update_layout(
        title=f"Plan vs Scrap - {component}",
        xaxis_title=period_col,
        yaxis_title="Quantity",
        barmode='group',
        hovermode='x unified',
        height=400,
        template='plotly_white'
    )

    st.plotly_chart(fig3, use_container_width=True)

    st.markdown("---")

# ==================== FOOTER ====================
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; font-size: 12px;'>
    <p>📊 SCRAP ANALYSIS DASHBOARD | Aktualizacja w czasie rzeczywistym</p>
    <p>Dane czytane z Excel report | Filtrowanie dynamiczne</p>
</div>
""", unsafe_allow_html=True)
