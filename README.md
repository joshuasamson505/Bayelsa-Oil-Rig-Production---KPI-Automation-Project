# Bayelsa Oil Rig Production - KPI Automation Project

### Table of Content

### Project Overview
Automated data analysis for oil rig production performance in Bayelsa State. This Python script automates repetitive KPI reporting tasks that would normally be done manually in Excel.

File: art.py | Dataset: week_1.xlsx

### Objectives & KPIs
The project evaluates 4 Rig Locations against 4 business targets:

| Objective | KPI Formula | Target | Logic |
| --- | --- | --- | --- |
| 1. Improve Productivity | UPH = Units Produced / Hours Worked | 11 | PASS if >= 11 |
| 2. Reduce Cost Per Unit | CPU = Production Cost (₦) / Units Produced | 750 | PASS if <= 750 |
| 3. Enhance Quality Control | UPI = Units Produced / Quality Issues | 250 | PASS if >= 250 |
| 4. Training Effectiveness | Avg Training Hours per Rig | 1 hr/week | PASS if >= 1 |

### Tech Stack
- pandas - Data cleaning, grouping, aggregation
- openpyxl - Excel report generation + native BarCharts
- plotly - Interactive Dashboard (3x3 subplot)

### Workflow in art.py

#### STEP 1: Load, Clean & Understand
```python
import pandas as pd
from plotly.subplots import make_subplots
import plotly.graph_objects as go
from openpyxl.chart import BarChart, Reference
import openpyxl

df = pd.read_excel("week_1.xlsx")
df.columns = df.columns.str.strip()
df["Rig Location"] = df["Rig Location"].str.strip()
print("Shape:", df.shape)
print("Missing values: \n", df.isnull().sum())
print("Duplicates:",df.duplicated().sum())
# No missing, or duplicates in this dataset

STEP 2: OBJECTIVE 1 - Productivity (UPH)

df["UPH"] = df["Units Produced"] / df["Hours Worked"]
uph_by_rig = df.groupby("Rig Location")["UPH"].mean().round(2)
print(uph_by_rig)
print("FAIL (<11):", uph_by_rig[uph_by_rig < 11].index.tolist())

STEP 3: OBJECTIVE 2 - Cost Per Unit (CPU)

df["CPU"] = df["Production Cost (₦)"] / df["Units Produced"]
cpu_by_rig = df.groupby("Rig Location")["CPU"].mean().round(2)
print(cpu_by_rig)
print("FAIL (>750):", cpu_by_rig[cpu_by_rig > 750].index.tolist())

STEP 4: OBJECTIVE 3 - Quality Control (UPI)

df["UPI"] = df["Units Produced"] / df["Quality Issues"].replace(0,1)
upi_by_rig = df.groupby("Rig Location")["UPI"].mean().round(1)
print("PASS (>250)", upi_by_rig[upi_by_rig >=250].index.tolist())
print("FAIL (<250):", upi_by_rig[upi_by_rig < 250].index.tolist())

STEP 5: Training Effectiveness & Correlation

avg_training = df.groupby("Rig Location")["Training Hours"].mean().round(2)
corr = df[["Training Hours", "UPH", "CPU", "UPI"]].corr()
print(corr.loc["Training Hours"])

avg_units_by_training = df.groupby("Training Hours")["Units Produced"].mean().reset_index().sort_values("Training Hours")
avg_units_by_hours = df.groupby("Hours Worked")["Units Produced"].mean().reset_index().sort_values("Hours Worked")

STEP 6: Summary Table & Exports

summary = df.groupby("Rig Location").agg({
 "UPH": "mean",
 "CPU": "mean",
 "UPI": "mean",
 "Training Hours": "mean"
}).reset_index()

summary = summary.rename(columns={
 "UPH": "avg_UPH",
 "CPU": "avg_CPU",
 "UPI": "avg_UPI",
 "Training Hours": "avg_Training"
})

summary["UPH_Status"] = summary["avg_UPH"].apply(lambda x: "PASS" if x >= 11 else "FAIL")
summary["CPU_Status"] = summary["avg_CPU"].apply(lambda x: "PASS" if x <= 750 else "FAIL")
summary["UPI_Status"] = summary["avg_UPI"].apply(lambda x: "PASS" if x >= 250 else "FAIL")
summary["Training_Status"] = summary["avg_Training"].apply(lambda x: "PASS" if x >= 1 else "FAIL")

# FINAL EXPORTS
summary.to_excel("Bayelsa_KPI_Report.xlsx", index=False)
df.to_excel("Bayelsa_Rig_Full_With_KPIs.xlsx", index=False)

STEP 7: Excel Dashboard with Charts

summary.to_excel("Bayelsa_KPI_Report_With_Charts.xlsx", index=False)
wb = openpyxl.load_workbook("Bayelsa_KPI_Report_With_Charts.xlsx")
ws = wb.active

chart1 = BarChart()
chart1.title = "UPH vs Target 11"
chart1.y_axis.title = "UPH"
data = Reference(ws, min_col=2, min_row=1, max_row=5, max_col=2)
cats = Reference(ws, min_col=1, min_row=2, max_row=5)
chart1.add_data(data, titles_from_data=True)
chart1.set_categories(cats)
ws.add_chart(chart1, "F2")

chart2 = BarChart()
chart2.title = "CPU vs Target 750"
data2 = Reference(ws, min_col=3, min_row=1, max_row=5, max_col=3)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats)
ws.add_chart(chart2, "F18")

wb.save("Bayelsa_KPI_Report_With_Charts.xlsx")

<img width="1366" height="681" alt="Bayelsa KPI Report With Chart" src="https://github.com/user-attachments/assets/358bf876-a66f-4e6d-ac14-4922c7ccaf98" />


STEP 8: Plotly Interactive Dashboard

fig = make_subplots(rows=3, cols=3, subplot_titles=("Avg Units", "Avg Cost", "Avg Quality Issues", "Avg Hours Worked", "Avg Training Hours", "", "Training Hours vs Units", "Hours Worked vs Units"))

fig.add_trace(go.Bar(y=avg_units.index, x=avg_units.values, orientation="h"), row=1, col=1)

<img width="450" height="217" alt="AvgUnits" src="https://github.com/user-attachments/assets/63e9680a-9030-464e-a472-dedccc5a8df0" />

fig.add_trace(go.Bar(y=avg_cost.index, x=avg_cost.values, orientation="h"), row=1, col=2)

<img width="440" height="226" alt="AvgCost" src="https://github.com/user-attachments/assets/6d346d2f-e21e-42a2-ad8d-0518ff23e5ee" />

fig.add_trace(go.Bar(y=avg_quality.index, x=avg_quality.values, orientation="h"), row=1, col=3)

<img width="439" height="221" alt="AvgQualityIssues" src="https://github.com/user-attachments/assets/10cb6982-027d-48cc-81dd-1394c8522571" />

fig.add_trace(go.Bar(y=avg_hours.index, x=avg_hours.values, orientation="h"), row=2, col=1)
fig.add_trace(go.Bar(y=avg_training.index, x=avg_training.values, orientation="h"), row=2, col=2)

<img width="469" height="234" alt="AvgTrainingHours" src="https://github.com/user-attachments/assets/9627047f-ae81-4169-bc62-f40c2131e047" />

fig.add_trace(go.Scatter(x=avg_units_by_training["Training Hours"], y=avg_units_by_training["Units Produced"], mode="lines+markers"), row=3, col=1)
fig.add_trace(go.Scatter(x=avg_units_by_hours["Hours Worked"], y=avg_units_by_hours["Units Produced"], mode="lines+markers"), row=3, col=2)

fig.update_layout(title_text="Oil Rig Production Dashboard - Bayelsa", height=900, width=1400, template="plotly_white", showlegend=False)
fig.show()
fig.write_html("dashboard.html")

Results & Outputs
After running python art.py:

1. *Console Output:* Shape, missing values, duplicates, UPH/CPU/UPI/Training by Rig Location, FAIL lists, and correlation matrix.
2. *Bayelsa_KPI_Report.xlsx:* Summary PASS/FAIL table.
3. *Bayelsa_Rig_Full_With_KPIs.xlsx:* Full dataset with calculated UPH, CPU, UPI columns.
4. *Bayelsa_KPI_Report_With_Charts.xlsx:* Executive report with embedded BarCharts.
5. *dashboard.html:* Interactive Plotly dashboard - 8 charts showing Avg Units, Cost, Quality, Hours, Training, and trend lines for Training vs Production.

Key Insights (Example)
- Rigs failing UPH < 11 need productivity review
- Rigs failing CPU > 750 need cost optimization
- Rigs failing UPI < 250 need quality control intervention
- Correlation analysis shows impact of Training Hours on UPH/CPU/UPI

Future Improvement
Add target lines on Plotly charts, automate weekly email reports, add conditional formatting in openpyxl.
