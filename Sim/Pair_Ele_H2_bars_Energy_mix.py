# -*- coding: utf-8 -*-
"""
Created on Sat Oct  4 20:39:57 2025

@author: Mhdella
"""

import os
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

# === Custom y-axis tick formatter for GWh ===
def scaled_formatter(x, pos):
    return f"{x / 1e6:.1f} GWh"  # Input values are in MWh, convert to GWh

# === Emissions threshold adjustment for right y-axis ===
def plot_energy_mix_from_file(opgf_file, figures_folder, scenario_label, h2_prop):
    import pandas as pd
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter
    import os

    # Load OPGF data
    df = pd.read_excel(opgf_file)
    df.set_index("Player Type", inplace=True)

    # Remove External Gas and P2G
    df = df.drop([tech for tech in ["meth", "P2G"] if tech in df.index], errors="ignore")

    # Define tech groups
    electric_techs = ["GT", "CHP", "biomass", "wind", "solar", "FC"]
    hydrogen_techs = ["EZ", "G2H"]

    # Shiny colors
    colors_map = {
        "GT": "#1f77b4",      # Blue
        "CHP": "#9400D3",     # Shiny violet
        "biomass": "#8c564b", # Brown
        "wind": "#2ca02c",    # Green
        "solar": "#FFD700",   # Shiny gold
        "FC": "#00CED1",      # Shiny turquoise
        "EZ": "#FF69B4",      # Shiny pink
        "G2H": "#bfbfbf"      # Light grey
    }

    legend_labels = {
        "GT": "GT",
        "CHP": "CHP",
        "biomass": "Biomass",
        "wind": "Wind",
        "solar": "Solar",
        "FC": "Fuel Cell",
        "EZ": "Electrolyser",
        "G2H": "H2 Reformer"
    }

    years = [col for col in df.columns if col.isdigit()]
    x = range(len(years))
    bar_width = 0.35

    # Font sizes (default font family, like your template)
    axis_title_fs = 20
    tick_label_fs = 18
    legend_fs = 16

    fig, ax = plt.subplots(figsize=(14, 7), dpi=200)

    # --- Plot electricity bars ---
    bottom_elec = [0] * len(years)
    for tech in electric_techs:
        if tech in df.index:
            vals = df.loc[tech, years].values
            ax.bar([i - bar_width/2 for i in x], vals, bar_width,
                   bottom=bottom_elec, color=colors_map[tech], label=legend_labels[tech])
            bottom_elec = [i + j for i, j in zip(bottom_elec, vals)]

    # --- Plot hydrogen bars ---
    bottom_h2 = [0] * len(years)
    for tech in hydrogen_techs:
        if tech in df.index:
            vals = df.loc[tech, years].values
            ax.bar([i + bar_width/2 for i in x], vals, bar_width,
                   bottom=bottom_h2, color=colors_map[tech], label=legend_labels[tech])
            bottom_h2 = [i + j for i, j in zip(bottom_h2, vals)]

    # --- Formatting ---
    ax.set_xticks(x)
    ax.set_xticklabels(years, fontsize=22)  # Bigger + bold years
    ax.set_ylabel("Energy (GWh)", fontsize=axis_title_fs)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, pos: f"{int(v/1e3)}"))

    # y-axis ticks stay normal
    ax.tick_params(axis="y", labelsize=tick_label_fs)

    # x-axis ticks explicitly larger
    ax.tick_params(axis="x", labelsize=22)
    ax.grid(visible=True, axis="y", linestyle="--", alpha=0.6)



    ## Specicied or Unified title
    # For reference with scenario and H2 proportion:
    # ax.set_title(f"Energy Mix and CO₂ Emissions – {scenario_label}, H₂: {h2_prop}",
    #               fontsize=20)
    
    ax.set_title("Energy Mix and CO₂ Emissions", fontsize=22)
    
    # --- Load emissions ---
    results_folder = os.path.dirname(opgf_file)
    tech_file = os.path.join(results_folder, "tech_econ_analysis_FS.xlsx")
    emissions = None
    if os.path.exists(tech_file):
        df_em = pd.read_excel(tech_file, sheet_name=None)
        emissions = []
        for year in years:
            if year in df_em:
                sheet = df_em[year]
                if "Emissions_OPGF" in sheet.columns:
                    emissions.append(sheet["Emissions_OPGF"].values[0])
                else:
                    emissions.append(0)
            else:
                emissions.append(0)

        # Plot emissions with threshold adjustment
        ax2 = ax.twinx()
        max_emission = max(emissions)
        if max_emission < 0.1:
            ax2.set_ylim(-0.15, 5)
        else:
            ax2.set_ylim(0, max_emission * 1.1)
        
        ax2.plot(x, emissions, color="red", linestyle="--", marker="o",
                 label="CO₂ Emissions", linewidth=3, markersize=9)
        ax2.set_ylabel("CO₂ Emissions \n(Tonnes)", fontsize=axis_title_fs, color="red")
        ax2.tick_params(axis="y", colors="red", labelsize=tick_label_fs)
        ax2.yaxis.set_label_coords(1.1, 0.32)
        ax2.yaxis.label.set_color("red")

    # --- Legend outside ---
    handles, labels = ax.get_legend_handles_labels()
    by_label = dict(zip(labels, handles))
    ax.legend(by_label.values(), by_label.keys(),
              bbox_to_anchor=(1.06, 1), loc="upper left", fontsize=legend_fs,
              # title="Technologies", title_fontsize=legend_fs)
              title="Type", title_fontsize=legend_fs)


    plt.tight_layout()
    os.makedirs(figures_folder, exist_ok=True)
    
    # Adjust scenario label for filename
    file_scenario_label = scenario_label
    if scenario_label == "Base_MC_Interv":
        file_scenario_label = "STI"
    
    fig_path = os.path.join(figures_folder, f"Fig_EnergyMix_{file_scenario_label}_H2_{h2_prop}.png")
    plt.savefig(fig_path, dpi=300)
    plt.show()
    print(f"✅ Saved styled plot: {fig_path}")


# === Loop to process all files ===
if __name__ == "__main__":
    parent_folder = r"C:\Users\Mhdella\Desktop\Res_EZATRFCCHP_Optimal_HB_HP_50_50_Price3"
    figures_folder = os.path.join(parent_folder, "Energy_mix_bar_plots")
    HPHB_sw = ['HB_Interv', 'HP_Interv', 'HP_HB_Interv', 'Base_MC_Interv']
    h2_props = ["0.0", "0.1", "0.2", "1.0"]
    investment_scenario = "FS"

    for root, dirs, files in os.walk(parent_folder):
        for file in files:
            if file.startswith("OPGF_H2_") and file.endswith(".xlsx"):
                parts = file.replace(".xlsx", "").split("_")
                h2_prop = parts[2]
                inv_scenario = parts[3]

                if h2_prop in h2_props and inv_scenario == investment_scenario:
                    for heat_scenario in HPHB_sw:
                        if heat_scenario in root:
                            opgf_file = os.path.join(root, file)
                            plot_energy_mix_from_file(opgf_file, figures_folder, heat_scenario, h2_prop)
