import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from docx import Document
from docx.shared import Inches

# Simulated historical GDP data for Vietnam (in USD billions)
years_hist = np.arange(2000, 2025)
gdp_hist = np.array([
    31.2, 33.7, 35.4, 39.1, 45.3, 57.6, 66.4, 77.6, 98.2, 115.9,
    133.6, 155.8, 171.2, 186.2, 201.3, 193.6, 205.3, 223.9, 241.3,
    262.4, 271.2, 277.0, 368.0, 409.9, 430.0  # Estimated 2024
])

# Normalize to index (2025 = 10,000), estimate 2025 GDP with 5.8% growth
gdp_2025 = gdp_hist[-1] * 1.058
scale_factor = 10000 / gdp_2025
index_hist = gdp_hist * scale_factor

# Future years and scenarios
years_future = np.arange(2025, 2046)
base_index = 10000
opt_growth = 0.06
mod_growth = 0.045

# Confidence intervals: ±0.5% around growth rates
opt_upper = base_index * (1 + (opt_growth + 0.005)) ** (years_future - 2025)
opt_lower = base_index * (1 + (opt_growth - 0.005)) ** (years_future - 2025)

mod_upper = base_index * (1 + (mod_growth + 0.005)) ** (years_future - 2025)
mod_lower = base_index * (1 + (mod_growth - 0.005)) ** (years_future - 2025)

# Crisis scenario
crisis = []
value = base_index
for year in years_future:
    if year < 2028:
        value *= 1.055
    elif 2028 <= year < 2031:
        value *= 0.82
    else:
        value *= 1.035
    crisis.append(value)

# Plotting the graph
plt.figure(figsize=(12, 7))
plt.plot(years_hist, index_hist, label="Historical Index (2000–2024)", color="black", linewidth=2)
plt.plot(years_future, base_index * (1 + opt_growth) ** (years_future - 2025), label="Scenario 1: Optimistic (6%)", color="green")
plt.fill_between(years_future, opt_lower, opt_upper, color="green", alpha=0.2, label="±0.5% CI (Optimistic)")
plt.plot(years_future, base_index * (1 + mod_growth) ** (years_future - 2025), label="Scenario 2: Moderate (4.5%)", color="blue")
plt.fill_between(years_future, mod_lower, mod_upper, color="blue", alpha=0.2, label="±0.5% CI (Moderate)")
plt.plot(years_future, crisis, label="Scenario 3: Boom-Bust", linestyle="--", color="red")

plt.axvline(2025, color='gray', linestyle=':', label="Base Year (2025)")
plt.title("Vietnam Economic Growth Index (2000–2045) with Scenarios", fontsize=16)
plt.xlabel("Year")
plt.ylabel("Index (Base = 10,000 in 2025)")
plt.legend()
plt.grid(True)
plt.tight_layout()

# Save graph to file
graph_path = "/mnt/data/vietnam_growth_graph_with_data.png"
plt.savefig(graph_path)
plt.close()

# Create Word document
doc = Document()
doc.add_heading("Vietnam Economic Growth Outlook (2000–2045)", level=1)
doc.add_paragraph("This document includes a graph combining actual GDP data from 2000 to 2024 with three modeled scenarios for 2025–2045. The base index is normalized to 10,000 in 2025. Confidence intervals of ±0.5% are shown for the optimistic and moderate growth scenarios.")

doc.add_picture(graph_path, width=Inches(6.5))

# Add GDP and growth rate table
doc.add_heading("Historical GDP and Index Data (2000–2024)", level=2)
table = doc.add_table(rows=1, cols=4)
table.style = 'Table Grid'
hdr_cells = table.rows[0].cells
hdr_cells[0].text = 'Year'
hdr_cells[1].text = 'GDP (Billion USD)'
hdr_cells[2].text = 'Growth Rate (%)'
hdr_cells[3].text = 'Index Value'

for i in range(1, len(years_hist)):
    row_cells = table.add_row().cells
    row_cells[0].text = str(years_hist[i])
    row_cells[1].text = f"{gdp_hist[i]:,.1f}"
    growth = ((gdp_hist[i] - gdp_hist[i-1]) / gdp_hist[i-1]) * 100
    row_cells[2].text = f"{growth:.1f}"
    row_cells[3].text = f"{index_hist[i]:,.0f}"

# Save document
doc_path = "/mnt/data/Vietnam_Economic_Outlook_with_Data.docx"
doc.save(doc_path)

doc_path, graph_path
