"""
CDC Provisional Natality Dashboard (2025)
=========================================
A sophisticated, pedagogical Streamlit dashboard designed for undergraduate
business analytics students to explore geographic, monthly, and sex-based
patterns in provisional 2025 US live birth counts.

Author: Development Agent (Pair Programmer)
Audience: Undergraduate Business Analytics Students
Dataset: CDC National Center for Health Statistics (NCHS) Provisional Natality 2025
"""

import os
from pathlib import Path
from typing import Dict, List, Tuple

import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# -----------------------------------------------------------------------------
# Configuration and Constants
# -----------------------------------------------------------------------------

PAGE_TITLE = "CDC Provisional Natality Dashboard (2025)"
PAGE_ICON = "👶"

# Chronological order for calendar months
MONTH_ORDER: List[str] = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

# Standard 2-letter postal abbreviations for all 50 states + District of Columbia
STATE_ABBREV_MAP: Dict[str, str] = {
    "Alabama": "AL", "Alaska": "AK", "Arizona": "AZ", "Arkansas": "AR",
    "California": "CA", "Colorado": "CO", "Connecticut": "CT", "Delaware": "DE",
    "District of Columbia": "DC", "Florida": "FL", "Georgia": "GA", "Hawaii": "HI",
    "Idaho": "ID", "Illinois": "IL", "Indiana": "IN", "Iowa": "IA",
    "Kansas": "KS", "Kentucky": "KY", "Louisiana": "LA", "Maine": "ME",
    "Maryland": "MD", "Massachusetts": "MA", "Michigan": "MI", "Minnesota": "MN",
    "Mississippi": "MS", "Missouri": "MO", "Montana": "MT", "Nebraska": "NE",
    "Nevada": "NV", "New Hampshire": "NH", "New Jersey": "NJ", "New Mexico": "NM",
    "New York": "NY", "North Carolina": "NC", "North Dakota": "ND", "Ohio": "OH",
    "Oklahoma": "OK", "Oregon": "OR", "Pennsylvania": "PA", "Rhode Island": "RI",
    "South Carolina": "SC", "South Dakota": "SD", "Tennessee": "TN", "Texas": "TX",
    "Utah": "UT", "Vermont": "VT", "Virginia": "VA", "Washington": "WA",
    "West Virginia": "WV", "Wisconsin": "WI", "Wyoming": "WY"
}

# Accessible, WCAG-compliant color palette
COLORS = {
    "primary": "#1E3A8A",      # Deep Navy
    "secondary": "#0D9488",    # Teal
    "accent": "#D97706",       # Warm Amber
    "male": "#3B82F6",         # Balanced Blue
    "female": "#0D9488",       # Accessible Teal
    "total": "#1E293B",        # Slate
    "highlight_top": "#047857",# Emerald
    "highlight_bot": "#B91C1C",# Crimson
    "card_bg": "#F8FAFC",      # Light background
    "border": "#E2E8F0"        # Neutral border
}


# -----------------------------------------------------------------------------
# Data Ingestion and Validation Layer
# -----------------------------------------------------------------------------

def find_data_file() -> Path:
    """
    Locate the workbook path dynamically.
    Works seamlessly on Windows (local) and Linux (Streamlit Community Cloud).
    """
    base_dir = Path(__file__).parent
    candidate_paths = [
        base_dir / "Data" / "Provisional_Natality_2025_CDC.xlsx",
        base_dir / "data" / "Provisional_Natality_2025_CDC.xlsx",
        base_dir.parent / "Data" / "Provisional_Natality_2025_CDC.xlsx",
        base_dir.parent / "data" / "Provisional_Natality_2025_CDC.xlsx",
    ]
    for p in candidate_paths:
        if p.exists():
            return p
    raise FileNotFoundError(
        "Could not find 'Provisional_Natality_2025_CDC.xlsx' in 'Data/' or 'data/' folder. "
        "Please ensure the workbook is uploaded to the data directory."
    )


@st.cache_data(show_spinner="Loading and validating CDC Natality data...")
def load_and_validate_data() -> Tuple[pd.DataFrame, Dict[str, any]]:
    """
    Load the Excel workbook, validate benchmarks, and enrich columns.
    Returns:
        df (pd.DataFrame): The cleaned and typed dataset.
        audit_info (dict): Results of data quality verification checks.
    """
    file_path = find_data_file()

    # Discover the sheet dynamically (handles title character-limit truncation)
    wb = openpyxl.load_workbook(file_path, read_only=True)
    first_sheet = wb.sheetnames[0]
    wb.close()

    df = pd.read_excel(file_path, sheet_name=first_sheet)

    # Perform quality audit checks against expected CDC benchmarks
    row_count = len(df)
    missing_count = int(df.isnull().sum().sum())
    duplicate_count = int(df.duplicated().sum())
    total_births = int(df["Births"].sum())
    geo_count = int(df["State of Residence"].nunique())
    month_count = int(df["Month Code"].nunique())
    sex_count = int(df["Sex of Infant"].nunique())

    audit_info = {
        "sheet_name": first_sheet,
        "rows": row_count,
        "cols": len(df.columns),
        "missing": missing_count,
        "duplicates": duplicate_count,
        "total_births": total_births,
        "geographies": geo_count,
        "months": month_count,
        "sex_categories": sex_count,
        "valid": (
            row_count == 1224 and
            missing_count == 0 and
            duplicate_count == 0 and
            total_births == 3604640 and
            geo_count == 51 and
            month_count == 12 and
            sex_count == 2
        )
    }

    # Data enrichment: add postal abbreviation and categorical month ordering
    df["State_Abbrev"] = df["State of Residence"].map(STATE_ABBREV_MAP)
    df["Month"] = pd.Categorical(df["Month"], categories=MONTH_ORDER, ordered=True)

    return df, audit_info


# -----------------------------------------------------------------------------
# Header Component
# -----------------------------------------------------------------------------

def render_header(audit_info: Dict[str, any]):
    """Renders the dashboard title, executive explanation, and policy notices."""
    st.title("👶 CDC Provisional Natality Dashboard (2025)")
    st.markdown(
        """
        An interactive analytical workspace designed for **business analytics students** to explore
        geographic concentration, seasonal fluctuations, and demographic patterns in provisional US live births.
        """
    )

    # Source, Provisional Notice, and Count-vs-Rate Callouts
    col_notice1, col_notice2 = st.columns([1, 1])

    with col_notice1:
        st.info(
            "🏛️ **Source & Provisional Status:** Data provided by the **CDC National Center for Health "
            "Statistics (NCHS)**. Figures represent **provisional 2025 counts** subject to ongoing vital "
            "registration updates and final annual closure.",
            icon="ℹ️"
        )

    with col_notice2:
        st.warning(
            "⚠️ **Analytical Rule — Birth Counts vs. Birth Rates:** All figures represent **absolute counts of live births ($N$)**, "
            "**not birth rates**. Computing birth or fertility rates requires external census population denominators, "
            "which are not part of this dataset.",
            icon="📊"
        )

    # Verification status indicator
    if audit_info["valid"]:
        st.caption("✅ **Data Integrity Verified:** 1,224 records, 51 geographies, 12 months, 0 missing values, 3,604,640 total births.")
    else:
        st.caption("⚠️ **Data Notice:** Dataset loaded with minor variances compared to standard baseline.")


# -----------------------------------------------------------------------------
# Sidebar Filtering and State Management
# -----------------------------------------------------------------------------

def initialize_session_state(all_states: List[str]):
    """Ensure session state variables exist with proper defaults."""
    if "ms_states" not in st.session_state:
        st.session_state["ms_states"] = all_states
    if "ms_months" not in st.session_state:
        st.session_state["ms_months"] = MONTH_ORDER
    if "radio_sex" not in st.session_state:
        st.session_state["radio_sex"] = "All"


def select_all_states(states: List[str]):
    st.session_state["ms_states"] = states


def clear_all_states():
    st.session_state["ms_states"] = []


def select_all_months():
    st.session_state["ms_months"] = MONTH_ORDER


def clear_all_months():
    st.session_state["ms_months"] = []


def reset_all_filters(states: List[str]):
    st.session_state["ms_states"] = states
    st.session_state["ms_months"] = MONTH_ORDER
    st.session_state["radio_sex"] = "All"


def render_sidebar(df: pd.DataFrame) -> pd.DataFrame:
    """Renders sidebar controls, filter helpers, reset button, and returns filtered dataframe."""
    all_states = sorted(df["State of Residence"].unique().tolist())
    initialize_session_state(all_states)

    st.sidebar.header("🔍 Filter Controls")
    st.sidebar.markdown("Tailor your analytical sample across geography, calendar season, and infant sex.")

    # 1. State / Geography Filter
    st.sidebar.subheader("1. Geography Selection")
    col_st1, col_st2 = st.sidebar.columns(2)
    col_st1.button("Select All States", on_click=select_all_states, args=(all_states,), use_container_width=True)
    col_st2.button("Clear States", on_click=clear_all_states, use_container_width=True)

    selected_states = st.sidebar.multiselect(
        "Select State(s) / Geographies:",
        options=all_states,
        key="ms_states"
    )

    # 2. Month Filter
    st.sidebar.subheader("2. Calendar Month")
    col_m1, col_m2 = st.sidebar.columns(2)
    col_m1.button("Select All Months", on_click=select_all_months, use_container_width=True)
    col_m2.button("Clear Months", on_click=clear_all_months, use_container_width=True)

    selected_months = st.sidebar.multiselect(
        "Select Month(s):",
        options=MONTH_ORDER,
        key="ms_months"
    )

    # 3. Infant Sex Filter
    st.sidebar.subheader("3. Infant Sex")
    sex_options = ["All", "Female", "Male"]
    selected_sex = st.sidebar.radio(
        "Filter by Infant Sex:",
        options=sex_options,
        horizontal=True,
        key="radio_sex"
    )

    # 4. Global Reset Button
    st.sidebar.markdown("---")
    st.sidebar.button("🔄 Reset All Filters", on_click=reset_all_filters, args=(all_states,), type="primary", use_container_width=True)

    # 5. Active Filter Summary
    st.sidebar.markdown("---")
    st.sidebar.subheader("Active Filter Summary")
    st.sidebar.markdown(f"• **Geographies:** {len(selected_states)} / {len(all_states)} selected")
    st.sidebar.markdown(f"• **Months:** {len(selected_months)} / 12 selected")
    st.sidebar.markdown(f"• **Sex Filter:** {selected_sex}")

    # Apply filters to produce the analytical subset
    filtered_df = df.copy()
    filtered_df = filtered_df[filtered_df["State of Residence"].isin(selected_states)]
    filtered_df = filtered_df[filtered_df["Month"].isin(selected_months)]
    if selected_sex != "All":
        filtered_df = filtered_df[filtered_df["Sex of Infant"] == selected_sex]

    return filtered_df


# -----------------------------------------------------------------------------
# KPI Summary Cards
# -----------------------------------------------------------------------------

def render_kpi_cards(filtered_df: pd.DataFrame, total_geos_selected: int, total_months_selected: int):
    """Renders 5 top-level KPI metrics reflecting the current filter selection."""
    st.markdown("### 📈 Key Performance Indicators")

    if filtered_df.empty:
        st.warning(
            "⚠️ **No observations match your active filter combination.** "
            "Please select at least one geography and one month, or click **Reset All Filters** in the sidebar."
        )
        return

    # 1. Total Births
    total_births = int(filtered_df["Births"].sum())

    # 2. Selected Geographies
    geo_count = filtered_df["State of Residence"].nunique()

    # 3. Average births per selected month
    month_count = filtered_df["Month"].nunique()
    avg_per_month = total_births / month_count if month_count > 0 else 0

    # 4. Highest geography
    geo_grouped = filtered_df.groupby("State of Residence")["Births"].sum()
    peak_geo_name = geo_grouped.idxmax() if not geo_grouped.empty else "N/A"
    peak_geo_val = int(geo_grouped.max()) if not geo_grouped.empty else 0

    # 5. Highest month
    month_grouped = filtered_df.groupby("Month")["Births"].sum()
    peak_month_name = str(month_grouped.idxmax()) if not month_grouped.empty else "N/A"
    peak_month_val = int(month_grouped.max()) if not month_grouped.empty else 0

    # Display in 5 responsive columns
    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:
        st.metric(
            label="Total Births",
            value=f"{total_births:,}",
            help="Total number of live births recorded in the active filter selection."
        )
    with c2:
        st.metric(
            label="Geographies Active",
            value=f"{geo_count} / 51",
            help="Number of US States and DC included in the current selection."
        )
    with c3:
        st.metric(
            label="Avg Births / Month",
            value=f"{avg_per_month:,.0f}",
            help="Average birth volume per selected calendar month across active geographies."
        )
    with c4:
        st.metric(
            label="Peak Geography",
            value=peak_geo_name,
            delta=f"{peak_geo_val:,} births",
            delta_color="off",
            help="The geography with the highest cumulative birth count in this selection."
        )
    with c5:
        st.metric(
            label="Peak Month",
            value=peak_month_name,
            delta=f"{peak_month_val:,} births",
            delta_color="off",
            help="The calendar month with the highest cumulative birth count in this selection."
        )


# -----------------------------------------------------------------------------
# Tab 1: Overview Tab
# -----------------------------------------------------------------------------

def render_overview_tab(filtered_df: pd.DataFrame):
    """Monthly birth trend, top/bottom geography comparison, and executive takeaways."""
    st.subheader("Monthly Seasonality & Volume Overview")
    st.markdown(
        "Examine high-level seasonal movement in national birth counts and observe the dispersion "
        "between the largest and smallest volume states."
    )

    col1, col2 = st.columns([3, 2])

    with col1:
        # Monthly Birth Trend Line Chart
        monthly_trend = (
            filtered_df.groupby(["Month Code", "Month"], as_index=False)["Births"]
            .sum()
            .sort_values("Month Code")
        )
        # Filter out months with 0 births if deselected
        monthly_trend = monthly_trend[monthly_trend["Births"] > 0]

        fig_trend = px.line(
            monthly_trend,
            x="Month",
            y="Births",
            markers=True,
            title="Total Live Births by Calendar Month (2025)",
            labels={"Births": "Live Births (Count)", "Month": "Calendar Month"},
        )
        fig_trend.update_traces(
            line=dict(color=COLORS["primary"], width=3),
            marker=dict(size=8, color=COLORS["accent"]),
            hovertemplate="<b>%{x}</b><br>Births: %{y:,.0f}<extra></extra>"
        )
        fig_trend.update_layout(
            template="plotly_white",
            hovermode="x unified",
            yaxis=dict(rangemode="tozero", tickformat=",d"),
            margin=dict(l=40, r=20, t=50, b=40)
        )
        st.plotly_chart(fig_trend, use_container_width=True)

    with col2:
        # Top 5 vs Bottom 5 Geographies
        geo_totals = (
            filtered_df.groupby("State of Residence", as_index=False)["Births"]
            .sum()
            .sort_values("Births", ascending=False)
        )

        if len(geo_totals) >= 10:
            top5 = geo_totals.head(5).copy()
            top5["Category"] = "Top 5 Volume"
            bot5 = geo_totals.tail(5).copy().sort_values("Births", ascending=True)
            bot5["Category"] = "Bottom 5 Volume"
            top_bot = pd.concat([top5, bot5])

            fig_tb = px.bar(
                top_bot,
                x="Births",
                y="State of Residence",
                color="Category",
                orientation="h",
                title="Top 5 vs. Bottom 5 Geographies by Birth Volume",
                labels={"Births": "Live Births", "State of Residence": "State"},
                color_discrete_map={"Top 5 Volume": COLORS["highlight_top"], "Bottom 5 Volume": COLORS["highlight_bot"]}
            )
            fig_tb.update_layout(
                template="plotly_white",
                yaxis=dict(categoryorder="total ascending"),
                xaxis=dict(rangemode="tozero", tickformat=",d"),
                margin=dict(l=40, r=20, t=50, b=40),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            fig_tb.update_traces(hovertemplate="<b>%{y}</b><br>Total Births: %{x:,.0f}<extra></extra>")
            st.plotly_chart(fig_tb, use_container_width=True)
        else:
            # When fewer than 10 states are selected
            fig_small = px.bar(
                geo_totals,
                x="Births",
                y="State of Residence",
                orientation="h",
                title="Selected Geographies by Birth Volume",
                labels={"Births": "Live Births", "State of Residence": "State"},
                color_discrete_sequence=[COLORS["primary"]]
            )
            fig_small.update_layout(
                template="plotly_white",
                yaxis=dict(categoryorder="total ascending"),
                xaxis=dict(rangemode="tozero", tickformat=",d"),
                margin=dict(l=40, r=20, t=50, b=40)
            )
            st.plotly_chart(fig_small, use_container_width=True)

    # Analytical Takeaways for Students
    st.markdown("#### 💡 Executive Insights for Business Analysts")
    st.markdown(
        """
        1. **Seasonal Peak:** In the full US dataset, births consistently rise through late summer (**July, August, and September**), with July recording the peak monthly volume (321,538 births) followed closely by August (319,245 births).
        2. **Winter Trough:** February records the fewest births (274,834), partly driven by having only 28 days, though daily birth velocities are also traditionally lower in early winter.
        3. **Geographic Concentration:** Birth volume follows general population scale. California, Texas, Florida, and New York represent a large proportion of total national births.
        """
    )


# -----------------------------------------------------------------------------
# Tab 2: Geographic Analysis Tab
# -----------------------------------------------------------------------------

def render_geographic_tab(filtered_df: pd.DataFrame):
    """Interactive US Choropleth map, full state ranking bar chart, and state-by-month heatmap."""
    st.subheader("Geographic Distribution & State Comparisons")
    st.markdown(
        "Analyze spatial distribution across the United States. Hover over states to inspect exact "
        "birth counts for the selected criteria."
    )

    # 1. State Choropleth Map
    geo_agg = (
        filtered_df.groupby(["State of Residence", "State_Abbrev"], as_index=False)["Births"]
        .sum()
    )

    fig_map = px.choropleth(
        geo_agg,
        locations="State_Abbrev",
        locationmode="USA-states",
        color="Births",
        scope="usa",
        color_continuous_scale="Tealgrn",
        hover_name="State of Residence",
        hover_data={"State_Abbrev": False, "Births": ":,"},
        title="Interactive US Live Birth Volume Map (Provisional 2025)"
    )
    fig_map.update_layout(
        geo=dict(bgcolor="rgba(0,0,0,0)"),
        margin=dict(l=0, r=0, t=40, b=0),
        coloraxis_colorbar=dict(title="Births", tickformat=",d")
    )
    st.plotly_chart(fig_map, use_container_width=True)

    st.markdown("---")

    col_rank, col_heat = st.columns([1, 1])

    with col_rank:
        st.markdown("#### State Volume Rankings")
        ranking_slice = st.radio(
            "Show ranking range:",
            options=["Top 10", "Top 25", "All Selected"],
            horizontal=True,
            key="ranking_slice"
        )
        ranked_df = geo_agg.sort_values("Births", ascending=False)
        if ranking_slice == "Top 10":
            plot_ranked = ranked_df.head(10)
        elif ranking_slice == "Top 25":
            plot_ranked = ranked_df.head(25)
        else:
            plot_ranked = ranked_df

        fig_rank = px.bar(
            plot_ranked,
            x="Births",
            y="State of Residence",
            orientation="h",
            labels={"Births": "Live Births", "State of Residence": "State"},
            color="Births",
            color_continuous_scale="Blues"
        )
        fig_rank.update_layout(
            template="plotly_white",
            yaxis=dict(categoryorder="total ascending"),
            xaxis=dict(rangemode="tozero", tickformat=",d"),
            coloraxis_showscale=False,
            height=480,
            margin=dict(l=40, r=20, t=20, b=40)
        )
        fig_rank.update_traces(hovertemplate="<b>%{y}</b><br>Total Births: %{x:,.0f}<extra></extra>")
        st.plotly_chart(fig_rank, use_container_width=True)

    with col_heat:
        st.markdown("#### State-by-Month Heatmap")
        st.caption("Displays seasonal intensity across selected states. Darker shading indicates higher birth counts.")

        # Limit heatmap display if too many states selected to keep it clean and legible
        heatmap_states = ranked_df.head(15)["State of Residence"].tolist()
        heat_data = filtered_df[filtered_df["State of Residence"].isin(heatmap_states)]

        pivot_heat = heat_data.pivot_table(
            index="State of Residence",
            columns="Month",
            values="Births",
            aggfunc="sum"
        ).fillna(0)

        # Preserve chronological month order on columns
        active_cols = [m for m in MONTH_ORDER if m in pivot_heat.columns]
        pivot_heat = pivot_heat[active_cols]
        # Sort states by descending volume
        pivot_heat = pivot_heat.loc[pivot_heat.sum(axis=1).sort_values(ascending=True).index]

        if not pivot_heat.empty:
            fig_heat = px.imshow(
                pivot_heat,
                labels=dict(x="Month", y="State", color="Births"),
                color_continuous_scale="Viridis",
                aspect="auto",
                height=480
            )
            fig_heat.update_layout(
                template="plotly_white",
                margin=dict(l=40, r=20, t=20, b=40),
                coloraxis_colorbar=dict(tickformat=",d")
            )
            st.plotly_chart(fig_heat, use_container_width=True)
            st.caption("Showing the top 15 volume states for visual clarity.")
        else:
            st.info("Select at least one state and month to render heatmap.")


# -----------------------------------------------------------------------------
# Tab 3: Monthly & Sex Analysis Tab
# -----------------------------------------------------------------------------

def render_monthly_sex_tab(filtered_df: pd.DataFrame):
    """Female vs. Male comparisons, monthly trends by sex, and demographic distribution."""
    st.subheader("Monthly Seasonality & Infant-Sex Analysis")
    st.markdown(
        "Evaluate biological sex distribution across calendar months. In human demographics, "
        "male births typically slightly exceed female births with a biological sex ratio of approximately "
        "**105 males per 100 females** (~51.2% male)."
    )

    # 1. Total Sex Breakdown Cards
    sex_agg = filtered_df.groupby("Sex of Infant", as_index=False)["Births"].sum()
    total_in_subset = sex_agg["Births"].sum()

    col_m1, col_m2, col_m3 = st.columns(3)

    female_births = int(sex_agg[sex_agg["Sex of Infant"] == "Female"]["Births"].sum())
    male_births = int(sex_agg[sex_agg["Sex of Infant"] == "Male"]["Births"].sum())

    female_pct = (female_births / total_in_subset * 100) if total_in_subset > 0 else 0
    male_pct = (male_births / total_in_subset * 100) if total_in_subset > 0 else 0
    sex_ratio = (male_births / female_births * 100) if female_births > 0 else 0

    with col_m1:
        st.metric(
            label="Female Live Births",
            value=f"{female_births:,}",
            delta=f"{female_pct:.1f}% of cohort",
            delta_color="off"
        )
    with col_m2:
        st.metric(
            label="Male Live Births",
            value=f"{male_births:,}",
            delta=f"{male_pct:.1f}% of cohort",
            delta_color="off"
        )
    with col_m3:
        st.metric(
            label="Secondary Sex Ratio",
            value=f"{sex_ratio:.1f} M : 100 F" if female_births > 0 else "N/A",
            delta="Benchmark: ~105.0",
            delta_color="normal" if 104 <= sex_ratio <= 106 else "off",
            help="Number of male births per 100 female births. Natural biological baseline is ~105."
        )

    st.markdown("---")

    # 2. Monthly Trend Split by Infant Sex
    col_chart1, col_chart2 = st.columns([3, 2])

    with col_chart1:
        trend_sex = (
            filtered_df.groupby(["Month Code", "Month", "Sex of Infant"], as_index=False)["Births"]
            .sum()
            .sort_values("Month Code")
        )
        trend_sex = trend_sex[trend_sex["Births"] > 0]

        fig_sex_trend = px.bar(
            trend_sex,
            x="Month",
            y="Births",
            color="Sex of Infant",
            barmode="group",
            title="Monthly Live Births Grouped by Infant Sex",
            labels={"Births": "Live Births", "Month": "Month", "Sex of Infant": "Sex"},
            color_discrete_map={"Female": COLORS["female"], "Male": COLORS["male"]}
        )
        fig_sex_trend.update_layout(
            template="plotly_white",
            yaxis=dict(rangemode="tozero", tickformat=",d"),
            margin=dict(l=40, r=20, t=50, b=40),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        fig_sex_trend.update_traces(hovertemplate="<b>%{x}</b> (%{data.name})<br>Births: %{y:,.0f}<extra></extra>")
        st.plotly_chart(fig_sex_trend, use_container_width=True)

    with col_chart2:
        # Donut Chart for Overall Sex Proportion
        fig_donut = px.pie(
            sex_agg,
            names="Sex of Infant",
            values="Births",
            hole=0.55,
            title="Infant Sex Distribution (Active Selection)",
            color="Sex of Infant",
            color_discrete_map={"Female": COLORS["female"], "Male": COLORS["male"]}
        )
        fig_donut.update_traces(
            textinfo="percent+label",
            hovertemplate="<b>%{label}</b><br>Births: %{value:,.0f}<br>Share: %{percent}<extra></extra>"
        )
        fig_donut.update_layout(
            template="plotly_white",
            margin=dict(l=20, r=20, t=50, b=20),
            showlegend=False
        )
        st.plotly_chart(fig_donut, use_container_width=True)


# -----------------------------------------------------------------------------
# Tab 4: Data Table & Download Tab
# -----------------------------------------------------------------------------

def render_table_tab(filtered_df: pd.DataFrame):
    """Searchable interactive table, summary statistics, and CSV export."""
    st.subheader("Filtered Dataset & Data Export")
    st.markdown(
        "Examine the granular records corresponding to your active filter criteria. "
        "You can search, sort columns, or download the filtered dataset as a CSV file."
    )

    # 1. Download Button
    csv_data = filtered_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Download Filtered Data (CSV)",
        data=csv_data,
        file_name="cdc_provisional_natality_2025_filtered.csv",
        mime="text/csv",
        type="primary"
    )

    st.markdown(f"**Showing {len(filtered_df):,} matching observations**")

    # 2. Interactive Dataframe
    display_df = filtered_df[
        ["State of Residence", "Month", "Month Code", "Year Code", "Sex of Infant", "Births"]
    ].copy()

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "State of Residence": st.column_config.TextColumn("State of Residence", width="medium"),
            "Month": st.column_config.TextColumn("Month", width="small"),
            "Month Code": st.column_config.NumberColumn("Month Code", format="%d", width="small"),
            "Year Code": st.column_config.NumberColumn("Year Code", format="%d", width="small"),
            "Sex of Infant": st.column_config.TextColumn("Sex of Infant", width="small"),
            "Births": st.column_config.NumberColumn("Live Births (Count)", format="%d", width="medium"),
        }
    )

    # 3. Descriptive Summary Statistics
    st.markdown("#### 📊 Summary Statistics for Filtered Birth Counts")
    stats_df = filtered_df["Births"].describe().to_frame()
    stats_df.columns = ["Births (Summary)"]
    stats_df["Formatted Value"] = stats_df["Births (Summary)"].apply(lambda x: f"{x:,.2f}" if x < 1000 else f"{x:,.0f}")
    st.table(stats_df[["Formatted Value"]])


# -----------------------------------------------------------------------------
# Tab 5: About the Data Tab
# -----------------------------------------------------------------------------

def render_about_tab(audit_info: Dict[str, any]):
    """Pedagogical context, counts vs. rates explanation, and dataset metadata."""
    st.subheader("📚 About the Dataset & Business Analytics Primer")

    st.markdown(
        """
        ### 1. Understanding Birth Counts vs. Birth Rates
        A central tenet in business analytics and epidemiology is the distinction between **event counts** and **event rates**:
        
        * **Birth Count ($N$):** The raw number of live births recorded in a specific area and time window.
          $$\\text{Count} = \\sum \\text{Live Births}$$
          *Example:* California recorded 410,211 births in 2025, while Wyoming recorded 5,830 births.
          Comparing these two raw counts directly without context would be misleading because California's population is approximately 67 times larger than Wyoming's!
        
        * **Crude Birth Rate (CBR):** Measures birth events relative to the total living population ($P$):
          $$\\text{CBR} = \\frac{\\text{Total Live Births}}{\\text{Total Mid-Year Population}} \\times 1,000$$
        
        * **General Fertility Rate (GFR):** A more refined demographic metric measuring births relative only to the population capable of giving birth (typically females aged 15–44):
          $$\\text{GFR} = \\frac{\\text{Total Live Births}}{\\text{Female Population Aged 15–44}} \\times 1,000$$

        > **Analytics Rule:** Because this CDC dataset contains only counts and no population denominators,
        > we cannot calculate birth rates without pulling in external census tables.

        ---

        ### 2. CDC Provisional Vital Statistics Lifecycle
        * **What does 'Provisional' mean?** Provisional natality records are based on birth certificates registered in state vital statistics offices and transmitted to the CDC's National Center for Health Statistics (NCHS).
        * **Reporting Lag:** While records are received continuously, there can be a 1–3 month reporting lag.
        * **Final Natality Files:** Final annual natality data are typically released 9 to 12 months after the close of the calendar year, following extensive duplicate resolution and demographic auditing.

        ---

        ### 3. Balanced Orthogonal Panel Structure
        This dataset represents a perfectly balanced factorial panel:
        $$\\text{Total Rows} = 51 \\text{ Geographies} \\times 12 \\text{ Months} \\times 2 \\text{ Sexes} = 1,224 \\text{ Observations}$$
        * No data suppression: In many CDC WONDER releases, cell counts below 10 are suppressed ($<10$). In this provisional release, all cells are reported (minimum cell count is 177 in Vermont).
        * Total national births in this 2025 provisional file: **3,604,640**.

        ---

        ### 4. Data Dictionary
        """
    )

    data_dict = pd.DataFrame([
        {"Variable": "State of Residence", "Data Type": "String (Categorical)", "Definition": "US State of mother's residence (50 states plus District of Columbia)."},
        {"Variable": "Month", "Data Type": "String (Categorical)", "Definition": "Calendar month name (January through December)."},
        {"Variable": "Month Code", "Data Type": "Integer (1–12)", "Definition": "Numeric representation of the calendar month."},
        {"Variable": "Year Code", "Data Type": "Integer (2025)", "Definition": "Calendar year of occurrence (provisional 2025)."},
        {"Variable": "Sex of Infant", "Data Type": "String (Binary)", "Definition": "Recorded sex of the newborn ('Female' or 'Male')."},
        {"Variable": "Births", "Data Type": "Integer (Count)", "Definition": "Raw count of registered live births in that cell."},
    ])
    st.table(data_dict)


# -----------------------------------------------------------------------------
# Main Application Flow
# -----------------------------------------------------------------------------

def main():
    # Configure wide page layout and modern branding
    st.set_page_config(
        page_title=PAGE_TITLE,
        page_icon=PAGE_ICON,
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # Custom styling for KPI cards and layout polish
    st.markdown(
        """
        <style>
        .stMetric {
            background-color: #f8fafc;
            border: 1px solid #e2e8f0;
            padding: 12px 16px;
            border-radius: 8px;
            box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);
        }
        div[data-testid="stMetricValue"] {
            font-size: 1.7rem;
            font-weight: 700;
            color: #1e3a8a;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    # 1. Load data with cached validation
    try:
        df, audit_info = load_and_validate_data()
    except Exception as e:
        st.error(f"❌ Error loading dataset: {e}")
        st.stop()

    # 2. Render Header
    render_header(audit_info)
    st.markdown("---")

    # 3. Sidebar Filtering
    filtered_df = render_sidebar(df)

    # 4. KPI Metrics Cards
    total_geos_selected = len(st.session_state.get("ms_states", []))
    total_months_selected = len(st.session_state.get("ms_months", []))
    render_kpi_cards(filtered_df, total_geos_selected, total_months_selected)

    # If filters return an empty dataframe, stop rendering tabs gracefully
    if filtered_df.empty:
        st.stop()

    st.markdown("---")

    # 5. Tabbed Analytical Navigation
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Overview",
        "🗺️ Geographic Analysis",
        "📈 Monthly & Sex Analysis",
        "📋 Data Table & Download",
        "📚 About the Data"
    ])

    with tab1:
        render_overview_tab(filtered_df)

    with tab2:
        render_geographic_tab(filtered_df)

    with tab3:
        render_monthly_sex_tab(filtered_df)

    with tab4:
        render_table_tab(filtered_df)

    with tab5:
        render_about_tab(audit_info)


if __name__ == "__main__":
    main()
