# CDC Provisional Natality 2025 Dashboard

An interactive, pedagogical Streamlit dashboard designed for undergraduate business analytics students to explore geographic distribution, seasonal fluctuations, and demographic patterns in provisional 2025 US live births.

---

## 🏛️ Project Overview & Rules
* **Data Source:** Centers for Disease Control and Prevention (CDC) National Center for Health Statistics (NCHS), Provisional Natality 2025.
* **Analytical Rule:** All metrics represent **live birth counts ($N$)**, not birth rates or fertility rates. Because the dataset does not contain mid-year census populations or demographic denominators, rates cannot be computed.
* **Provisional Status:** 2025 figures are provisional and subject to ongoing vital statistics registration updates.
* **Panel Design:** Perfectly balanced factorial dataset containing $51 \text{ geographies} \times 12 \text{ months} \times 2 \text{ infant sexes} = 1,224 \text{ observations}$, totaling **3,604,640 live births**.

---

## 🚀 Quickstart: Running Locally

1. **Prerequisites:** Python 3.10+ installed.
2. **Install Dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
3. **Run the Streamlit Dashboard:**
   ```bash
   streamlit run app.py
   ```
4. Open your web browser to `http://localhost:8501`.

---

## ☁️ Deploying to Streamlit Community Cloud

1. Create a GitHub repository named `cdc-births-2025`.
2. Push `app.py`, `requirements.txt`, `README.md`, and the `Data/` folder containing `Provisional_Natality_2025_CDC.xlsx`.
3. In [Streamlit Community Cloud](https://streamlit.io/cloud), create a new app pointing to your repository.
4. Set the **Main file path** to `app.py` and click **Deploy**.

---

## 📊 Dashboard Structure

1. **Header & Context:**
   - Dashboard title, educational scope, CDC source attribution.
   - Distinct notice of provisional status and the "counts vs. rates" analytical boundary.
2. **Sidebar Controls:**
   - Multi-select filters for 51 Geographies (with Select All / Clear All helpers).
   - Multi-select filters for 12 Calendar Months (chronological ordering preserved).
   - Radio selector for Infant Sex (`All`, `Female`, `Male`).
   - One-click `Reset All Filters` button.
   - Dynamic summary of active filters.
3. **KPI Metrics:**
   - Total live births in current filter selection.
   - Active geographies count ($N / 51$).
   - Average monthly birth volume across selected months.
   - Geography with highest cumulative birth count.
   - Calendar month with highest cumulative birth count.
4. **Interactive Tabs:**
   - **Overview:** Monthly birth seasonality trend line chart and Top 5 vs. Bottom 5 volume states.
   - **Geographic Analysis:** Interactive Plotly US state choropleth map, ranked volume bar chart, and state-by-month heatmap.
   - **Monthly & Sex Analysis:** Male vs. Female monthly grouped bar comparison, overall sex ratio ($M:F \approx 1.05$), and cohort distribution donut.
   - **Data Table & Download:** Searchable, formatted data table, summary statistics, and one-click CSV export button.
   - **About the Data:** Pedagogical guide explaining counts vs. rates, fertility formulas, vital statistics registration lifecycle, and complete data dictionary.

---

## 🧪 Browser-Based Testing Checklist
- [x] Default unfiltered state loads 3,604,640 total births.
- [x] Filter by single state (e.g. California) updates all KPIs and charts consistently.
- [x] Filter by single month (e.g. August) updates monthly averages and totals.
- [x] Filter by sex (Female: 1,760,948; Male: 1,843,692).
- [x] Reset filters button restores all 51 states, 12 months, and All sex.
- [x] Deselecting all states/months shows graceful empty state warning without Python errors.
- [x] CSV download exports filtered records.
- [x] Map renders correctly with hover details.
