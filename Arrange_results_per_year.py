# -*- coding: utf-8 -*-
"""
Created on May 5, 2025
@author: Mhdella
Combined code: Extracts tables from Excel, saves to TXT, then to Excel with sheets.
"""

# -*- coding: utf-8 -*-
"""
Created on Wed May 21 02:07:27 2025

@author: Mhdella
"""

import pandas as pd
import os
import getpass

# === PARAMETERS ===

mix_fg="70/30 HB&HP"
# mix_fg="50/50 HB&HP"

output_folder = 'CCS_0_6Yrs_EZATR_Opt_HB_HP_70_30_Price3'

# output_folder = '6Yrs_EZATR_Opt_HB_HP_50_50_Price3'
# output_folder = '6Yrs_EZATR_Opt_HB_HP_70_30_Price3'
# output_folder = 'Socio_TEA_Endeg_H2Cost_HB_HP_70_30_Price3_EZATR_0'
# output_folder = 'Socio_TEA_HB_HP_70_30_Price4_EZATR_0'
# output_folder ='Socio_TEA_HB_HP_70_30_Price4_EZATR_0.5_CCS_0'
# output_folder = 'Socio_TEA_C02_0.05_HB_HP_70_30_Price4_EZATR_0.5'


# output_folder = 'Socio_TEA_50'

# price_fg="price_5"
# price_fg="price_4"
price_fg="price_3"
# price_fg="price_2"
# price_fg="price_1"

# CCS_fg="With_CCS"
CCS_fg="Without_CCS"


# year = "2035"
year = "2050"


# Automatically detect current username
username = getpass.getuser()

# Automatically detect the folder where the script is running
# (assumes script is in: C:\Users\{username}\Desktop\{project_folder})
script_dir = os.path.abspath(os.getcwd())
project_base = script_dir  # This is your project folder, like TEA_THD_DH_CCS

# Construct base path from detected folder
base_path = os.path.join(project_base, "Output", output_folder, "MCP", "1_iter")

# Scenario and ratio config
scenarios = ["Base_MC_Interv", "HP_HB_Interv", "HP_Interv", "HB_Interv"]
h2_ratios = {"A": ("H_0.0", 0), "B": ("H_0.1", 10), "C": ("H_0.2", 20), "D": ("H_1.0", 100)}

# === COLLECT RESULTS ===
results = {scenario: [] for scenario in scenarios}

for scenario in scenarios:
    for case, (h2_folder, h2_percentage) in h2_ratios.items():
        file_path = os.path.join(
            base_path, h2_folder, CCS_fg, "Wind_limit_0", "ext_grid_0", price_fg,
            scenario, "Social_MC", "FS", "Results", "tech_econ_analysis_FS.xlsx"
        )

        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            continue

        try:
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            target_sheet = next((s for s in sheet_names if str(year) in s or "Year" in s), sheet_names[0])

            df = pd.read_excel(file_path, sheet_name=target_sheet)
            if not all(col in df.columns for col in ["H2 Proportion", "Cost_OPGF", "Emissions_OPGF"]):
                print(f"Missing required columns in {file_path}, sheet {target_sheet}")
                continue

            h2_prop = df["H2 Proportion"].iloc[0] * 100
            cost = df["Cost_OPGF"].iloc[0] / 1_000_000
            emissions = round(df["Emissions_OPGF"].iloc[0], 0)

            results[scenario].append({
                "Case": case,
                "H₂ Ratio (%)": h2_percentage,
                "Emissions (tonnes)": emissions,
                "Cost (m£)": round(cost, 2)
            })

        except Exception as e:
            print(f"Error reading {file_path}: {e}")
            continue

# === WRITE TO TXT FILE ===
txt_output_file = os.path.join(project_base, "Output", output_folder, f"tables_output_{year}.txt")
with open(txt_output_file, "w", encoding="utf-8") as f:
    for scenario in scenarios:
        if results[scenario]:
            f.write(f"\n{scenario}\n")
            f.write("Case\tH₂ Ratio (%)\tEmissions (tonnes)\tCost (m£)\n")
            for row in results[scenario]:
                f.write(f"{row['Case']}\t{row['H₂ Ratio (%)']}\t\t{row['Emissions (tonnes)']}\t\t\t{row['Cost (m£)']}\n")

print(f"\nText output saved to: {txt_output_file}")


# === WRITE TO EXCEL ===
columns = ["Case", "H₂ Ratio (%)", "Emissions (tonnes)", "Cost (m£)"]
excel_output_file = os.path.join(project_base, "Output", output_folder, f"tables_output_{year}.xlsx")

# Collect 100% H₂ rows (assumed last row in each scenario)
summary_100 = []

with pd.ExcelWriter(excel_output_file) as writer:
    for scenario, rows in results.items():
        if rows:
            df = pd.DataFrame(rows, columns=columns)
            df.to_excel(writer, sheet_name=scenario[:31], index=False)

            # Append last row (assumed 100% H₂) for summary
            summary_row = df.iloc[-1].copy()
            summary_row["Scenario"] = scenario  # Add scenario label
            summary_100.append(summary_row)

    # Write the summary sheet if any 100% entries found
    if summary_100:
        df_summary = pd.DataFrame(summary_100)
        # Reorder columns with 'Scenario' first
        cols_order = ["Scenario"] + [col for col in df_summary.columns if col != "Scenario"]
        df_summary = df_summary[cols_order]
        df_summary.to_excel(writer, sheet_name="H2_100_Summary", index=False)

print(f"Excel file saved successfully: {excel_output_file}")

#################################################################

import matplotlib.pyplot as plt
import numpy as np

from matplotlib.lines import Line2D
from matplotlib.patches import Patch

# === PLOT CONFIGURATION ===

# Color map: match H2 ratio to consistent color
h2_colors = {
    0: 'red',
    10: 'blue',
    20: 'orange',
    100: 'green'
}
h2_levels = [0, 10, 20, 100]
scenario_labels = scenarios  # All 4 scenarios
bar_width = 0.2
x = np.arange(len(scenario_labels))  # Scenario positions

# Ensure Figures folder exists
figures_dir = os.path.join(project_base, "Output", output_folder, "Scenario_H2_Figures")
os.makedirs(figures_dir, exist_ok=True)


# Map scenario names to standard labels for x-axis
scenario_name_map = {
    "Base_MC_Interv": "STI",
    "HP_HB_Interv": mix_fg,
    "HP_Interv": "100% HP",
    "HB_Interv": "100% HB"
}


x_labels = [scenario_name_map[s] for s in scenario_labels]
# Set up the figure
fig, ax1 = plt.subplots(figsize=(12, 6))
ax2 = ax1.twinx()

# === PLOT DATA ===
for i, h2 in enumerate(h2_levels):
    costs = []
    emissions = []

    for scenario in scenario_labels:
        match = next((row for row in results[scenario] if row["H₂ Ratio (%)"] == h2), None)
        if match:
            costs.append(match["Cost (m£)"])
            emissions.append(match["Emissions (tonnes)"])
        else:
            costs.append(0)
            emissions.append(0)

    # Bar offset position
    offset_x = x + (i - 1.5) * bar_width

    # Plot cost bars
    ax1.bar(offset_x, costs, bar_width, color=h2_colors[h2])

    # Plot emissions lines
    ax2.plot(x, emissions, marker='o', linestyle='--', color=h2_colors[h2])

# === AXIS SETTINGS ===
ax1.set_xlabel("Scenarios", fontsize=15)
ax1.set_ylabel("Cost (m£)", color="blue", fontsize=14)
# ax2.set_ylabel("Emissions (tonnes)", color="red", fontsize=14)
ax2.set_ylabel("Emissions (tonnes)", color="red", fontsize=14)
ax2.yaxis.set_label_coords(1.075, 0.45)
ax1.set_title(f"Cost and Emissions - Year {year}", fontsize=16)

# Ticks
ax1.set_xticks(x)
# ax1.set_xticklabels(scenario_labels, rotation=15, fontsize=14)
ax1.set_xticklabels(x_labels, rotation=15, fontsize=14)
ax1.tick_params(axis='y', labelsize=14)
ax2.tick_params(axis='y', labelsize=14)

all_emissions = []
for scenario in scenario_labels:
    for h2 in h2_levels:
        match = next((row for row in results[scenario] if row["H₂ Ratio (%)"] == h2), None)
        
        if match:
            all_emissions.append(match["Emissions (tonnes)"])
            max_emissions = max(all_emissions) if all_emissions else 1
            
            upper_limit = max(5, max_emissions * 1.1)  # 10% more than max or at least 5
            ax2.set_ylim(-0.3, upper_limit)

            if max_emissions>100:
                ax2.set_ylim(-0.03*max_emissions, upper_limit)


# Y-axis tick formatting
ax1.tick_params(axis='y', labelsize=14, colors='blue')   # Cost axis (left) in blue
ax2.tick_params(axis='y', labelsize=14, colors='red')    # Emissions axis (right) in red

# Gridlines
ax1.grid(True, which='major', axis='y', linestyle='--', linewidth=0.6, alpha=0.7)
ax1.grid(True, which='major', axis='x', linestyle='--', linewidth=0.6, alpha=0.5)

# === LEGEND: Only H2 Ratios ===

# Base legend for H2 ratios
color_patches = [Patch(facecolor=h2_colors[h], label=f'H₂ {h}%') for h in h2_levels]

# Add bar and line legend keys
bar_proxy = Patch(facecolor="gray", label="Cost (bar)")
line_proxy = Line2D([0], [0], color="gray", linestyle='--', marker='o', label="Emissions (line)")

# Combine all elements
legend_elements = [bar_proxy, line_proxy] + color_patches

# Add to plot
ax1.legend(handles=legend_elements, title="Legend", fontsize=13, title_fontsize=13,
           loc="upper left", bbox_to_anchor=(1.088, 1))



# === SAVE AND SHOW ===
plt.tight_layout()
plot_output_file = os.path.join(figures_dir, f"Fig_Scenario_BarLine_{year}.png")
plt.savefig(plot_output_file, dpi=300)
plt.show()

print(f"High-resolution figure saved to: {plot_output_file}")



# === SECOND FIGURE: X-axis is H₂ Ratios ===

fig2, ax1b = plt.subplots(figsize=(12, 6))
ax2b = ax1b.twinx()

bar_width = 0.2
x_h2 = np.arange(len(h2_levels))  # [0, 1, 2, 3]

# Colors for scenarios
scenario_colors = {
    "Base_MC_Interv": "red",
    "HP_HB_Interv": "blue",
    "HP_Interv": "orange",
    "HB_Interv": "green"
}

# === PLOT DATA ===
for i, scenario in enumerate(scenario_labels):
    costs = []
    emissions = []
    for h2 in h2_levels:
        match = next((row for row in results[scenario] if row["H₂ Ratio (%)"] == h2), None)
        if match:
            costs.append(match["Cost (m£)"])
            emissions.append(match["Emissions (tonnes)"])
        else:
            costs.append(0)
            emissions.append(0)

    offset_x = x_h2 + (i - 1.5) * bar_width

    # Plot bars and lines
    ax1b.bar(offset_x, costs, bar_width, color=scenario_colors[scenario])
    ax2b.plot(x_h2, emissions, marker='o', linestyle='--', color=scenario_colors[scenario])

# === AXIS SETTINGS ===
ax1b.set_xlabel("H₂ Ratio (%)", fontsize=15)
ax1b.set_ylabel("Cost (m£)", color="blue", fontsize=14)
ax2b.set_ylabel("Emissions (tonnes)", color="red", fontsize=14)
ax2b.yaxis.set_label_coords(1.075, 0.45)
ax1b.set_title(f"Cost and Emissions vs H₂ Ratio - Year {year}", fontsize=16)

# Set ticks and labels
ax1b.set_xticks(x_h2)
ax1b.set_xticklabels([f'{h}%' for h in h2_levels], fontsize=14)
ax1b.tick_params(axis='y', labelsize=14, colors='blue')
ax2b.tick_params(axis='y', labelsize=14, colors='red')

# === Y-axis scaling ===
all_emissions2 = [row["Emissions (tonnes)"] for s in scenario_labels for row in results[s] if "Emissions (tonnes)" in row]
max_emissions2 = max(all_emissions2) if all_emissions2 else 1
upper_limit2 = max(5, max_emissions2 * 1.1)
if max_emissions2 > 100:
    ax2b.set_ylim(-0.03 * max_emissions2, upper_limit2)
else:
    ax2b.set_ylim(-0.3, upper_limit2)

# === Grid and formatting ===
ax1b.grid(True, which='major', axis='y', linestyle='--', linewidth=0.6, alpha=0.7)
ax1b.grid(True, which='major', axis='x', linestyle='--', linewidth=0.6, alpha=0.5)

# === LEGEND: Now for scenarios ===
scenario_patches = [Patch(facecolor=scenario_colors[s], label=scenario_name_map[s]) for s in scenario_labels]
bar_proxy2 = Patch(facecolor="gray", label="Cost (bar)")
line_proxy2 = Line2D([0], [0], color="gray", linestyle='--', marker='o', label="Emissions (line)")
legend_elements2 = [bar_proxy2, line_proxy2] + scenario_patches

ax1b.legend(handles=legend_elements2, title="Legend", fontsize=13, title_fontsize=13,
            loc="upper left", bbox_to_anchor=(1.088, 1))

# === SAVE ===
plot_output_file2 = os.path.join(figures_dir, f"Fig_H2Ratio_BarLine_{year}.png")
plt.tight_layout()
plt.savefig(plot_output_file2, dpi=300)
plt.show()

print(f"Second figure saved to: {plot_output_file2}")


# === COMPREHENSIVE SUMMARY (NEW SEPARATE FILE) ===
summary_all = []

# Collect all rows across all scenarios
for scenario, rows in results.items():
    for row in rows:
        row_copy = row.copy()
        row_copy["Scenario"] = scenario
        summary_all.append(row_copy)

if summary_all:
    df_summary_all = pd.DataFrame(summary_all)

    # Map scenario names to nicer labels
    scenario_name_map = {
        "Base_MC_Interv": "(1) STI: Social-Technical Interventions",
        "HP_HB_Interv": f"(2) {mix_fg}",
        "HP_Interv": "(3) 100% Heat Pumps",
        "HB_Interv": "(4) 100% Hydrogen Boilers"
    }
    df_summary_all["Scenario"] = df_summary_all["Scenario"].map(scenario_name_map)

    # Pivot COSTS
    df_costs = df_summary_all.pivot_table(
        index="Scenario",
        columns="H₂ Ratio (%)",
        values="Cost (m£)",
        aggfunc="first"
    ).reset_index()
    df_costs = df_costs[["Scenario", 0, 10, 20, 100]]
    df_costs.rename(columns={0: "0%", 10: "10%", 20: "20%", 100: "100%"}, inplace=True)

    # Pivot EMISSIONS
    df_emissions = df_summary_all.pivot_table(
        index="Scenario",
        columns="H₂ Ratio (%)",
        values="Emissions (tonnes)",
        aggfunc="first"
    ).reset_index()
    df_emissions = df_emissions[["Scenario", 0, 10, 20, 100]]
    df_emissions.rename(columns={0: "0%", 10: "10%", 20: "20%", 100: "100%"}, inplace=True)


    # Save to a NEW separate Excel file
    # summary_excel_file = os.path.join(project_base, "Output", output_folder, f"tables_output_summary_{year}.xlsx")
    summary_excel_file = os.path.join(project_base, "Output", figures_dir, f"tables_output_summary_{year}.xlsx")
    with pd.ExcelWriter(summary_excel_file) as writer:
        df_costs.to_excel(writer, sheet_name="Costs", index=False)
        df_emissions.to_excel(writer, sheet_name="Emissions", index=False)

    print(f"Comprehensive summary Excel file saved: {summary_excel_file}")


# === WRITE ENERGY MIX FILE ===
# This part collects the energy mix values from OPGF_H2_*_FS.xlsx for each case
# and saves them into a separate Excel file, grouped by blending case.

# # === WRITE ENERGY MIX FILE ===

# energy_mix_output = os.path.join(project_base, "Output", output_folder, f"energy_mix_{year}.xlsx")
energy_mix_output = os.path.join(project_base, "Output", figures_dir, f"energy_mix_{year}.xlsx")

# Map scenario names to short labels (same as before)
scenario_name_map = {
    "Base_MC_Interv": "STI",
    "HP_HB_Interv": mix_fg,
    "HP_Interv": "100% HP",
    "HB_Interv": "100% HB"
}

# Blending ratios (folder mapping already in h2_ratios)
blend_labels = {0: "H2_0pct", 10: "H2_10pct", 20: "H2_20pct", 100: "H2_100pct"}

with pd.ExcelWriter(energy_mix_output) as writer:
    for h2_case, (h2_folder, h2_percentage) in h2_ratios.items():
        scenario_dfs = []
        for scenario in scenarios:
            # Build dynamic file path depending on blending ratio
            file_path = os.path.join(
                base_path, h2_folder, CCS_fg, "Wind_limit_0", "ext_grid_0", price_fg,
                scenario, "Social_MC", "FS", "Results", f"OPGF_H2_{h2_percentage/100:.1f}_FS.xlsx"
            )

            if not os.path.exists(file_path):
                print(f"Energy mix file not found: {file_path}")
                continue

            try:
                df_emix = pd.read_excel(file_path)
                df_emix.insert(0, "Scenario", scenario_name_map.get(scenario, scenario))
                scenario_dfs.append(df_emix)
            except Exception as e:
                print(f"Error reading {file_path}: {e}")
                continue

        if scenario_dfs:
            sheet_df = pd.concat(scenario_dfs, axis=0)
            sheet_name = blend_labels[h2_percentage]
            sheet_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

print(f"Energy mix Excel file saved: {energy_mix_output}")

