# Plant Daily Operational Dashboard — Streamlit

Streamlit replica of the Power BI plant dashboard built earlier in this project.

## Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Usage

1. Open the app in your browser (Streamlit opens it automatically, usually at http://localhost:8501).
2. In the left sidebar, upload one or more daily report files named like:
   `DAILY PLANT STATUS FOR 09 AUGUST 2026.xlsx`
3. The dashboard builds automatically:
   - **KPI cards**: Energy Sent Out, Max Load, Avg Generation, Avg Inflow, Total Energy Deficit
   - **Bar chart**: cumulative Expected vs Actual generation by unit through the selected date
   - **Gauge**: Capacity utilization for the selected day
   - **Trend line**: Expected vs Actual generation across every uploaded day (not filtered by the date picker)
   - **Unit status table**: per-unit actual generation and energy deficit, color-coded (amber/red), with offline units (0 MW available) shown in gray and excluded from the deficit color scale
   - **Offline alert box**: automatically appears for any unit with 0 MW available capacity, showing its logged comment

4. Use the **"Cumulative data as of"** dropdown (top right) to switch the cumulative end date for the KPI cards, bar chart, and table. The gauge shows capacity utilization for that selected day, while the trend line always shows the full history of every file you've uploaded.

## Notes

- The file parser expects the exact report layout used in the source `DAILY PLANT STATUS FOR ...` reports (same structure the Power BI version reads). If your organization's template changes row/column positions, the `parse_daily_file()` cell references in `app.py` will need updating to match — same as the Power Query function did.
- Filenames must contain `FOR <DD> <MONTH> <YYYY>` (e.g. `FOR 09 AUGUST 2026`) for the date to be extracted correctly.
- Files starting with `~$` (Excel lock files) are automatically skipped.
