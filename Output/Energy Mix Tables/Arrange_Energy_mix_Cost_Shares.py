# -*- coding: utf-8 -*-
"""
Created on Sat Oct 18 17:10:32 2025

@author: Mhdella
"""
# -*- coding: utf-8 -*-
"""
Energy Mix & Cost Analysis for 4 Heating Scenarios × 4 H₂ Blending Cases
Collects OPGF results, compiles energy mix and cost breakdowns into Excel files,
and generates combined plots for electricity and hydrogen.
"""

import os
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

# ==================== USER SETTINGS ====================
price_fg = "price_3"
CCS_fg = "With_CCS"         
# CCS_fg = "Without_CCS"
investment_scenario = "FS" 

# year = "2035"
year = "2050"


# =======================================================


main_folder = "0Res_EZATRFCCHP_Optimal_HB_HP_50_50_Price3"  
# main_folder = "Res_EZATRFCCHP_Opt_HB50_HP50"  
# main_folder = "Res_Allyears_CCS1_HB_HP_50_50_Price3_EZATRFCCHP_Optimal"  
# main_folder = "Res_Allyears_noCCS0_HB_HP_50_50_Price3_EZATRFCCHP_Optimal"  


    
# Automatically detect working directory
script_dir = os.path.abspath(os.getcwd())
project_base = script_dir

base_output_path = os.path.join(project_base, "Output", main_folder)
base_path = os.path.join(base_output_path, "MCP", "1_iter")


# # === New scenarios and hydrogen blending ===
scenarios = ["Base_MC_Interv", "HP_HB_Interv", "HP_Interv", "HB_Interv"]


h2_ratios = {
    "A": ("H_0.0", 0),
    "B": ("H_0.1", 10),
    "C": ("H_0.2", 20),
    "D": ("H_1.0", 100)
}

# ---- Helper: Read energy mix file ----
def read_energy_mix(opgf_file):
    df = pd.read_excel(opgf_file)
    df.set_index("Player Type", inplace=True)
    df = df.drop([tech for tech in ["meth", "P2G"] if tech in df.index], errors="ignore")
    return df


# # === Technology selection and legend labels ===
# selected_techs = ["GT", "CHP", "biomass", "wind", "solar", "FC", "EZ", "G2H"]
selected_techs = ["GT", "CHP", "biomass", "wind", "solar", "FC", "EZ", "G2H", "EmissionCost"]


legend_labels = {
    "GT": "GT",
    "CHP": "CHP",
    "biomass": "Biomass",
    "wind": "Wind",
    "solar": "Solar",
    "FC": "Fuel Cell",
    "EZ": "Electrolyser",
    "G2H": "H₂ Reformer"
}




def create_and_save_operational_cost_breakdown(
    opgf_file,
    output_dir,
    price_fg,
    CCS_fg,
    project_base,
    FES_scenario="FS",
    verbose=True
):
    """
    Create and save 'OperationalCost_Breakdown_FS.xlsx' for all years in the OPGF file.
    Reads carbon price dynamically from Excel file inside the project folder.
    Saves two sheets: 'Costs' (M£/day) and 'Shares' (%).
    """

    import pandas as pd
    import os

    # ================== EMISSION FACTORS (tCO₂/MWh) ==================
    emission_factors = {
        "GT": 0.448,
        "CHP": 0.1839,
        "biomass": 0.0892,
        "wind": 0.0,
        "solar": 0.0,
        "EZ": 0.0,
        "G2H": 0.1,
        "FC": 0.0
    }

    # ================== BASE AGGREGATED COSTS (£/MWh) ==================
    base_AggType_costs = pd.DataFrame({
        "type": ["GT", "biomass", "CHP", "wind", "solar", "EZ", "G2H", "FC"],
        # "costs": [68, 14, 66, 34, 32, None, None, None]
        "costs": [85, 202, 135, 46, 44, None, None, None]

    })

    # ----------------- PRICE SETTINGS -----------------
    if isinstance(price_fg, str) and price_fg.lower().startswith("price_"):
        price_flag_num = price_fg.lower().replace("price_", "")
    else:
        price_flag_num = str(price_fg)

    if price_flag_num == "0":
        ele_opex_price, gas_opex_price, hyd_opex_price = 33.69, 9.73, 9
    elif price_flag_num == "1":
        ele_opex_price, gas_opex_price, hyd_opex_price = 125, 37, 100
    elif price_flag_num == "2":
        ele_opex_price, gas_opex_price, hyd_opex_price = 125, 37, 120
    elif price_flag_num == "3":
        ele_opex_price, gas_opex_price, hyd_opex_price = 125, 37, 0
    elif price_flag_num == "4":
        ele_opex_price, gas_opex_price, hyd_opex_price = 125, 37, 75
    elif price_flag_num == "5":
        ele_opex_price, gas_opex_price, hyd_opex_price = 125, 37, None
    else:
        raise ValueError(f"Unknown price flag: {price_fg}")

    # ----------------- READ ENERGY MIX -----------------
    df = pd.read_excel(opgf_file)
    if "Player Type" not in df.columns:
        raise ValueError("Energy mix file must contain 'Player Type' column.")

    df.set_index("Player Type", inplace=True)
    years = df.columns.tolist()

    # ----------------- READ CARBON PRICE FROM EXCEL -----------------
    carbon_price_file = os.path.join(
        project_base,
        "Refined FES scenario inputs",
        "Carbon price",
        "carbon_price.xlsx"
    )

    if not os.path.exists(carbon_price_file):
        raise FileNotFoundError(f"❌ Carbon price file not found:\n{carbon_price_file}")

    df_carbon = pd.read_excel(carbon_price_file)
    df_carbon.columns = df_carbon.columns.astype(str).str.strip()
    df_carbon["Year"] = df_carbon["Year"].astype(int)

    if FES_scenario not in df_carbon.columns:
        raise ValueError(
            f"Scenario '{FES_scenario}' not found in carbon price file columns: {list(df_carbon.columns)}"
        )

    carbon_price_map = dict(zip(df_carbon["Year"], df_carbon[FES_scenario]))

    # Prepare dicts for pivoting
    costs_dict = {}
    shares_dict = {}

    # ----------------- LOOP OVER YEARS -----------------
    for year in years:
        energy_mix = df[year].to_dict()

        # Hydrogen cost from mix
        EZ_output = energy_mix.get("EZ", 0)
        G2H_output = energy_mix.get("G2H", 0)
        H2_total = EZ_output + G2H_output
        Cost_H2 = ((EZ_output * ele_opex_price) + (G2H_output * gas_opex_price)) / H2_total if H2_total > 0 else (hyd_opex_price or 0)

        # ----------------- ADJUST COSTS FOR CCS -----------------
        fossil_units = ["GT", "CHP", "biomass", "G2H"]

        if int(year) >= 2040 and "With_CCS" in CCS_fg:
            # CCS applied: emissions = 0, fossil costs +3%
            base_AggType_costs.loc[base_AggType_costs["type"].isin(fossil_units), "costs"] *= 1.03
            total_emission = 0
        else:
            total_emission = sum(
                energy_mix.get(tech, 0) * emission_factors.get(tech, 0)
                for tech in emission_factors)
                        
            

        # Fill dynamic prices
        base_AggType_costs.loc[base_AggType_costs["type"] == "EZ", "costs"] = ele_opex_price
        base_AggType_costs.loc[base_AggType_costs["type"] == "G2H", "costs"] = gas_opex_price
        base_AggType_costs.loc[base_AggType_costs["type"] == "FC", "costs"] = Cost_H2
        AggType_costs = base_AggType_costs.copy()

        # ----------------- OPERATIONAL COSTS -----------------
        efficiency_map = {"GT": 0.55, "CHP": 0.40, "FC": 0.60, "G2H": 0.75}
        zero_cost_types = {"EZ", "P2G", "meth"}
        player_costs = {}

        for player, output_mwh in energy_mix.items():
            if player in zero_cost_types:
                cost = 0.0
            elif player in efficiency_map:
                if player == "FC":
                    cost = output_mwh * (Cost_H2 / efficiency_map["FC"])
                elif player == "G2H":
                    gas_price = AggType_costs.loc[AggType_costs["type"] == "GT", "costs"].iloc[0]
                    cost = output_mwh * (gas_price / efficiency_map["G2H"])
                else:
                    marginal_cost = AggType_costs.loc[AggType_costs["type"] == player, "costs"].iloc[0]
                    cost = output_mwh * marginal_cost
            else:
                if player in AggType_costs["type"].values:
                    marginal_cost = AggType_costs.loc[AggType_costs["type"] == player, "costs"].iloc[0]
                    cost = output_mwh * marginal_cost
                else:
                    cost = 0.0
            player_costs[player] = cost

        # ----------------- EMISSION COST -----------------
        carbon_price_value = carbon_price_map.get(int(year), 0)  # £/tCO₂
        emission_cost = total_emission * carbon_price_value
        player_costs["EmissionCost"] = emission_cost
        player_costs["TotalEmission (tCO2)"]=total_emission
        
        # print('base_AggType_costs', base_AggType_costs)
        # print('carbon_price_value', carbon_price_value)
        # print('total_emission', total_emission)
        # print('emission_cost', emission_cost)
        
        # st=stop


        # ----------------- STORE RESULTS -----------------
        total_cost = sum(player_costs.values())
        costs_dict[year] = {k: (v / 1e6 if k != "TotalEmission (tCO2)" else v)
                            for k, v in player_costs.items()}  
        shares_dict[year] = {k: v / total_cost * 100 if total_cost != 0 else 0 for k, v in player_costs.items()}

    # ----------------- CREATE DATAFRAMES -----------------
    costs_df = pd.DataFrame(costs_dict).fillna(0)
    shares_df = pd.DataFrame(shares_dict).fillna(0)

    costs_df.index.name = "Type"
    shares_df.index.name = "Type"

    # ----------------- SAVE OUTPUT -----------------
    os.makedirs(output_dir, exist_ok=True)
    output_file = os.path.join(output_dir, "OperationalCost_Breakdown_FS.xlsx")

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        costs_df.to_excel(writer, sheet_name="Costs")
        shares_df.to_excel(writer, sheet_name="Shares")

    if verbose:
        print(f"✅ Operational cost breakdown saved:\n   {output_file}")
        

    # Return last year's results
    last_year = years[-1]
    breakdown_df = pd.DataFrame({
        "Values (M£/day)": list(costs_df[last_year]),
        "Shares (%)": list(shares_df[last_year])
    }, index=costs_df.index)


    # return breakdown_df, total_cost, Cost_H2
    return breakdown_df, total_cost, Cost_H2, costs_df.loc["TotalEmission (tCO2)"].to_dict()





# ===================== MAIN LOOP =====================
for key, (h_folder, h2_value) in h2_ratios.items():
    print(f"\n=== Processing H₂ blending case: {h_folder} ({h2_value}%) ===")

    # Output folders
    energy_mix_base = os.path.join(
        project_base,
        "Output",
         main_folder,
        "Energy_Mix_Costs"
        )
    
    
    # ===================== OUTPUT FOLDERS =====================    
    # Base output folder for the selected year
    energy_mix_year_folder = os.path.join(
        project_base,
        "Output",
        main_folder,
        f"Energy_Mix_{year}")
    os.makedirs(energy_mix_year_folder, exist_ok=True)
    
    # Inside it, create "Energy_Mix_Costs"
    energy_mix_base = os.path.join(energy_mix_year_folder, "")
    os.makedirs(energy_mix_base, exist_ok=True)
    
    # Folder for each hydrogen blending case
    h2_folder = os.path.join(energy_mix_base, f"H2_{h2_value}")
    os.makedirs(h2_folder, exist_ok=True)
    
    # Folder for figures (plots)
    figures_folder = os.path.join(h2_folder, "Bar_plots")
    os.makedirs(figures_folder, exist_ok=True)
    
    print(f"📂 Energy mix results will be saved in:\n{h2_folder}\n")



    all_data = {}

    # --- Collect Energy Mix data ---
    for scenario in scenarios:
        opgf_file = os.path.join(
            base_path,
            h_folder,
            CCS_fg,
            "Wind_limit_0",
            "ext_grid_0",
            price_fg,
            scenario,
            "Social_MC",              
            investment_scenario,
            "Results",
            f"OPGF_H2_{h2_value/100:.1f}_{investment_scenario}.xlsx"
        )

        if os.path.exists(opgf_file):
            df = read_energy_mix(opgf_file)
            df_flat = df[[year]].rename(columns={year: f"{scenario}_H2_{h2_value}"})
            all_data.update(df_flat.to_dict(orient="series"))
        else:
            print(f"⚠ File not found: {opgf_file}")
            
    # Define output folder for breakdown file
        op_cost_dir = os.path.join(
            base_path,
            h_folder,
            CCS_fg,
            "Wind_limit_0",
            "ext_grid_0",
            price_fg,
            scenario,
            "Social_MC",
            investment_scenario,
            "Results")
        

        
        

        breakdown_df, total_cost, Cost_H2, emission_dict = create_and_save_operational_cost_breakdown(
            opgf_file=opgf_file,
            output_dir=op_cost_dir,
            price_fg=int(price_fg.split("_")[-1]),
            CCS_fg=CCS_fg,
            project_base=project_base,
            FES_scenario=investment_scenario)


        
        # Store emission for the selected year (e.g. "2035" or "2050")
        if year in emission_dict:
            emission_value = emission_dict[year]
        else:
            emission_value = 0
        
        # Save per-scenario emission (we’ll add it to the Energy Mix later)
        if "emission_values" not in locals():
            emission_values = {}
        emission_values[scenario] = emission_value

        

    # --- Combine and Save Energy Mix ---
    combined_df = pd.DataFrame(all_data)
    combined_df = combined_df.loc[combined_df.index.intersection(selected_techs)]
    combined_df.rename(index=legend_labels, inplace=True)
    combined_df = combined_df.round(2)
    combined_df.index.name = "Type"
    
    # --- Collect Cost Breakdown ---
    cost_values_dict = {}
    cost_shares_dict = {}
    
    for scenario in scenarios:
        op_cost_file = os.path.join(
            base_path,
            h_folder,
            CCS_fg,
            "Wind_limit_0",
            "ext_grid_0",
            price_fg,
            scenario,
            "Social_MC",
            investment_scenario,
            "Results",
            "OperationalCost_Breakdown_FS.xlsx"
        )
    
        if os.path.exists(op_cost_file):
            df_cost = pd.read_excel(op_cost_file, sheet_name="Costs", index_col=0)
            df_share = pd.read_excel(op_cost_file, sheet_name="Shares", index_col=0)
    
            # Keep only selected techs and rename rows
            df_cost = df_cost.loc[df_cost.index.intersection(selected_techs)]
            df_share = df_share.loc[df_share.index.intersection(selected_techs)]
            df_cost.rename(index=legend_labels, inplace=True)
            df_share.rename(index=legend_labels, inplace=True)
    
            # Add scenario suffix to columns
            if year in df_cost.columns:
                df_cost = df_cost[[year]]
            else:
                print(f"⚠ Year {year} not found in Operational Cost file for scenario {scenario}")
                continue
            
            if year in df_share.columns:
                df_share = df_share[[year]]
            else:
                print(f"⚠ Year {year} not found in Cost Shares file for scenario {scenario}")
                continue
            
            # Rename column to scenario name
            df_cost.columns = [f"{scenario}"]
            df_share.columns = [f"{scenario}"]
            cost_values_dict[scenario] = df_cost
            cost_shares_dict[scenario] = df_share
        else:
            print(f"⚠ Operational cost breakdown not found: {op_cost_file}")
    
    # Merge all scenarios horizontally
    if cost_values_dict:
        cost_values_df = pd.concat(cost_values_dict.values(), axis=1)
        cost_shares_df = pd.concat(cost_shares_dict.values(), axis=1)
    
        # Add total row
        cost_values_df.loc["Total"] = cost_values_df.sum(numeric_only=True)
        cost_shares_df.loc["Total"] = cost_shares_df.sum(numeric_only=True)
    
        cost_values_df.index.name = "Type"
        cost_shares_df.index.name = "Type"
    else:
        cost_values_df = pd.DataFrame()
        cost_shares_df = pd.DataFrame()
        
        # --- Add Emissions (tCO₂) Row after Wind ---
    if "Wind" in combined_df.index and "emission_values" in locals():
    
        # 1️⃣ Align emission columns with the same naming convention as EnergyMix
        aligned_emission = {}
        for col in combined_df.columns:
            base_col = col.replace(f"_H2_{h2_value}", "")
            if base_col in emission_values:
                val = emission_values[base_col]
                # Convert to MtCO₂ if very large
                if val > 1e6:
                    val = val / 1e6
                    unit = "Emissions (MtCO₂)"
                else:
                    unit = "Emissions (tCO₂)"
                aligned_emission[col] = val
            else:
                aligned_emission[col] = 0
    
        # 2️⃣ Create the aligned series
        emission_series = pd.Series(aligned_emission, name=unit)
    
        # 3️⃣ Insert right after Wind row
        wind_idx = combined_df.index.get_loc("Wind") + 1
        combined_df = pd.concat([
            combined_df.iloc[:wind_idx],
            emission_series.to_frame().T,
            combined_df.iloc[wind_idx:]
        ])
    
        print(f"✅ Added '{unit}' row for H₂={h2_value}% (aligned columns)")


    
    # --- Save Energy Mix + Costs into one file ---
    excel_file = os.path.join(h2_folder, f"EnergyMix_H2_{h2_value}.xlsx")
    with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
        combined_df.to_excel(writer, sheet_name="Energy Mix (MWh)")
        if not cost_values_df.empty:
            cost_values_df.to_excel(writer, sheet_name="Operational Cost (M£)")
            cost_shares_df.to_excel(writer, sheet_name="Cost Shares (%)")
    
    print(f"📁 Combined Excel file saved: {excel_file}")



    # ===================== PLOTS =====================
    combined_df_plot = combined_df.copy()

    electric_techs = ["GT", "CHP", "Biomass", "Wind", "Solar", "Fuel Cell"]
    hydrogen_techs = ["Electrolyser", "H₂ Reformer"]

    colors_map = {
        "GT": "#1f77b4",
        "CHP": "#9400D3",
        "Biomass": "#8c564b",
        "Wind": "#2ca02c",
        "Solar": "#FFD700",
        "Fuel Cell": "#00CED1",
        "Electrolyser": "#FF69B4",
        "H₂ Reformer": "#bfbfbf"
    }

    def plot_combined(df_subset, techs, title, filename, reverse_order=False):
        # available_columns = [c for c in df_subset.columns if c in scenarios]
        available_columns = df_subset.columns  # use all columns
        df_plot = df_subset.loc[df_subset.index.intersection(techs), available_columns].T

        fig, ax = plt.subplots(figsize=(12, 6), dpi=200)
        bottom = [0] * len(df_plot)
        plotting_order = reversed(techs) if reverse_order else techs

        for tech in plotting_order:
            if tech in df_plot.columns:
                ax.bar(
                    df_plot.index,
                    df_plot[tech].values,
                    bottom=bottom,
                    color=colors_map.get(tech, "#999999"),
                    label=tech
                )
                bottom = [b + v for b, v in zip(bottom, df_plot[tech].values)]

        ax.set_xlabel("Heating Scenarios", fontsize=14)
        ax.set_ylabel("Energy (GWh)", fontsize=14)
        ax.set_title(title, fontsize=18)
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, pos: f"{int(v/1e3)}"))
        ax.legend(fontsize=10, loc="upper left")
        ax.grid(True, axis="y", linestyle="--", alpha=0.6)
        plt.tight_layout()
        plt.savefig(os.path.join(figures_folder, filename), dpi=300)
        plt.close()
        print(f"  Saved plot: {filename}")

    plot_combined(combined_df_plot, electric_techs, f"Electricity Energy Mix (H₂={h2_value}%)", f"ElectricityMix_H2_{h2_value}.png")
    plot_combined(combined_df_plot, hydrogen_techs, f"Hydrogen Energy Mix (H₂={h2_value}%)", f"HydrogenMix_H2_{h2_value}.png", reverse_order=True)

print("\n🎯 All H₂ cases processed successfully.")


# ==============================================================
# 📘 FUNCTION: CREATE SUMMARY TABLE ACROSS ALL H₂ CASES
# ==============================================================

def create_summary_excel(year):
    """
    Combines Energy Mix, Operational Cost, and Cost Shares
    from all H₂ blending cases into one Excel summary file.
    """
    base_summary_path = os.path.join(
        project_base,
        "Output",
        main_folder,
        f"Energy_Mix_{year}"
    )

    summary_file = os.path.join(base_summary_path, f"All_summary_cost_{year}.xlsx")

    print(f"\n📊 Creating summary file for {year}...")
    print(f"   Searching in: {base_summary_path}\n")

    h2_folders = [f for f in os.listdir(base_summary_path) if f.startswith("H2_")]
    all_energy_mix = {}
    all_op_cost = {}
    all_cost_share = {}

    for h2_folder in h2_folders:
        h2_path = os.path.join(base_summary_path, h2_folder)
        excel_file = os.path.join(h2_path, f"EnergyMix_{h2_folder}.xlsx")

        if not os.path.exists(excel_file):
            print(f"⚠ Skipped (missing file): {excel_file}")
            continue

        print(f"✅ Reading from {excel_file}")

        try:
            # Read all three sheets
            df_mix = pd.read_excel(excel_file, sheet_name="Energy Mix (MWh)", index_col=0)
            df_cost = pd.read_excel(excel_file, sheet_name="Operational Cost (M£)", index_col=0)
            df_share = pd.read_excel(excel_file, sheet_name="Cost Shares (%)", index_col=0)

            # Add H2 ratio as suffix to columns
            df_mix.columns = [f"{col}_H2_{h2_folder.split('_')[1]}" for col in df_mix.columns]
            df_cost.columns = [f"{col}_H2_{h2_folder.split('_')[1]}" for col in df_cost.columns]
            df_share.columns = [f"{col}_H2_{h2_folder.split('_')[1]}" for col in df_share.columns]

            all_energy_mix[h2_folder] = df_mix
            all_op_cost[h2_folder] = df_cost
            all_cost_share[h2_folder] = df_share

        except Exception as e:
            print(f"❌ Error reading {excel_file}: {e}")

    # --- Combine Data ---
    if not all_energy_mix:
        print("⚠ No valid EnergyMix files found. Aborting summary creation.")
        return

    combined_mix = pd.concat(all_energy_mix.values(), axis=1)
    combined_cost = pd.concat(all_op_cost.values(), axis=1)
    combined_share = pd.concat(all_cost_share.values(), axis=1)

    # --- Save to summary Excel ---
    with pd.ExcelWriter(summary_file, engine="openpyxl") as writer:
        combined_mix.to_excel(writer, sheet_name="Energy Mix (MWh)")
        combined_cost.to_excel(writer, sheet_name="Operational Cost (M£)")
        combined_share.to_excel(writer, sheet_name="Cost Shares (%)")

    print(f"\n✅ Summary file created successfully:\n{summary_file}\n")

# ==============================================================
# 📢 EXECUTE SUMMARY CREATION
# ==============================================================

create_summary_excel(year)


# ==============================================================
# 📘 FUNCTION: CREATE SUMMARY OF TOTAL OPERATIONAL COSTS ONLY
# ==============================================================

def create_total_cost_summary(year):
    """
    Collects only the 'Total' row from the 'Operational Cost (M£)' sheets
    of all H₂ blending cases and all heating scenarios, and saves in
    Costs_Summary_<year>.xlsx
    """
    base_path_year = os.path.join(
        project_base,
        "Output",
        main_folder,
        f"Energy_Mix_{year}"
    )

    summary_file = os.path.join(base_path_year, f"Costs_Summary_{year}.xlsx")

    print(f"\n📊 Creating total operational cost summary for {year}...")
    print(f"   Searching in: {base_path_year}\n")

    # Sort H2 folders numerically by the percentage
    h2_folders = sorted(
        [f for f in os.listdir(base_path_year) if f.startswith("H2_")],
        key=lambda x: int(x.split("_")[1]))

    total_cost_summary = pd.DataFrame()

    for h2_folder in h2_folders:
        h2_path = os.path.join(base_path_year, h2_folder)
        excel_file = os.path.join(h2_path, f"EnergyMix_{h2_folder}.xlsx")

        if not os.path.exists(excel_file):
            print(f"⚠ Skipped (missing file): {excel_file}")
            continue

        df_cost = pd.read_excel(excel_file, sheet_name="Operational Cost (M£)", index_col=0)

        if "Total" not in df_cost.index:
            print(f"⚠ 'Total' row not found in {excel_file}")
            continue

        total_row = df_cost.loc["Total"]
        # Rename the column to the H2 ratio (0, 10, 20, 100)
        h2_percent = h2_folder.split("_")[1]
        total_row.name = h2_percent
        total_cost_summary = pd.concat([total_cost_summary, total_row.to_frame().T])

    # Optional: rename heating scenario columns for clarity
    total_cost_summary.index.name = "H2_ratio (%)"
    total_cost_summary.columns.name = "Scenario"

    # Reorder columns if needed
    total_cost_summary = total_cost_summary[scenarios]  # same order as your heating scenarios

    # Save to Excel
    total_cost_summary.to_excel(summary_file)
    print(f"\n✅ Total cost summary saved successfully:\n{summary_file}\n")

# ==============================================================
# 📢 EXECUTE TOTAL COST SUMMARY
# ==============================================================

create_total_cost_summary(year)


# ==============================================================
# 📘 FUNCTION: CREATE SUMMARY OF TOTAL EMISSIONS ONLY
# ==============================================================

def create_total_emissions_summary(year):
    """
    Collects only the 'Emissions (tCO₂)' row from the 'Energy Mix (MWh)' sheets
    of all H₂ blending cases and all heating scenarios, and saves in
    Emissions_Summary_<year>.xlsx
    """
    base_path_year = os.path.join(
        project_base,
        "Output",
        main_folder,
        f"Energy_Mix_{year}"
    )

    summary_file = os.path.join(base_path_year, f"Emissions_Summary_{year}.xlsx")

    print(f"\n📊 Creating total emissions summary for {year}...")
    print(f"   Searching in: {base_path_year}\n")

    # Sort H2 folders numerically by the percentage
    h2_folders = sorted(
        [f for f in os.listdir(base_path_year) if f.startswith("H2_")],
        key=lambda x: int(x.split("_")[1])
    )

    total_emissions_summary = pd.DataFrame()

    for h2_folder in h2_folders:
        h2_path = os.path.join(base_path_year, h2_folder)
        excel_file = os.path.join(h2_path, f"EnergyMix_{h2_folder}.xlsx")

        if not os.path.exists(excel_file):
            print(f"⚠ Skipped (missing file): {excel_file}")
            continue

        # Read 'Energy Mix (MWh)' sheet
        df_mix = pd.read_excel(excel_file, sheet_name="Energy Mix (MWh)", index_col=0)

        # Identify the emissions row: usually named 'Emissions (tCO₂)' or 'Emissions (MtCO₂)'
        emissions_row = [r for r in df_mix.index if "Emissions" in r]
        if not emissions_row:
            print(f"⚠ Emissions row not found in {excel_file}")
            continue

        emissions_values = df_mix.loc[emissions_row[0]]

        # Strip H2 suffix from column names so they match scenarios
        emissions_values.index = [col.split('_H2_')[0] for col in emissions_values.index]

        # Keep only the scenarios (ignore any extra columns)
        scenarios = ["Base_MC_Interv", "HP_HB_Interv", "HP_Interv", "HB_Interv"]
        emissions_values = emissions_values.reindex(scenarios, fill_value=0)

        # Rename the row to H2 ratio (0, 10, 20, 100)
        h2_percent = h2_folder.split("_")[1]
        emissions_values.name = h2_percent

        # Append to total summary
        total_emissions_summary = pd.concat([total_emissions_summary, emissions_values.to_frame().T])

    # Optional: rename columns and index
    total_emissions_summary.index.name = "H2_ratio (%)"
    total_emissions_summary.columns.name = "Scenario"

    # Save to Excel
    total_emissions_summary.to_excel(summary_file)
    print(f"\n✅ Total emissions summary saved successfully:\n{summary_file}\n")



create_total_emissions_summary(year)


# ==============================================================
# 📘 FUNCTION: CREATE SUMMARY OF TOTAL ENERGY SUPPLY AND SYSTEM EFFICIENCY
# ==============================================================

def create_total_supply_summary(year):
    """
    Collects the sum of energy supply rows from the 'Energy Mix (MWh)' sheets
    of all H₂ blending cases and all heating scenarios, and saves in:
    1️⃣ Supply_Summary_<year>.xlsx
    2️⃣ System_Efficiency_<year>.xlsx (reciprocal normalization by top-left cell)
    """
    import pandas as pd
    import os

    base_path_year = os.path.join(
        project_base,
        "Output",
        main_folder,
        f"Energy_Mix_{year}"
    )

    supply_file = os.path.join(base_path_year, f"Supply_Summary_{year}.xlsx")
    efficiency_file = os.path.join(base_path_year, f"System_Efficiency_{year}.xlsx")

    print(f"\n📊 Creating total supply summary for {year}...")
    print(f"   Searching in: {base_path_year}\n")

    # Define which rows to sum
    supply_rows = ["CHP", "Electrolyser", "Fuel Cell", "H₂ Reformer",
                   "GT", "Biomass", "Solar", "Wind"]

    # Sort H2 folders numerically by the percentage
    h2_folders = sorted(
        [f for f in os.listdir(base_path_year) if f.startswith("H2_")],
        key=lambda x: int(x.split("_")[1])
    )

    total_supply_summary = pd.DataFrame()

    for h2_folder in h2_folders:
        h2_path = os.path.join(base_path_year, h2_folder)
        excel_file = os.path.join(h2_path, f"EnergyMix_{h2_folder}.xlsx")

        if not os.path.exists(excel_file):
            print(f"⚠ Skipped (missing file): {excel_file}")
            continue

        df_mix = pd.read_excel(excel_file, sheet_name="Energy Mix (MWh)", index_col=0)

        # Keep only supply rows
        df_supply = df_mix.loc[df_mix.index.intersection(supply_rows)]

        # Sum across rows to get total supply per scenario
        supply_values = df_supply.sum(axis=0)

        # Strip H2 suffix from column names to match scenarios
        supply_values.index = [col.split('_H2_')[0] for col in supply_values.index]

        # Keep only the heating scenarios
        scenarios = ["Base_MC_Interv", "HP_HB_Interv", "HP_Interv", "HB_Interv"]
        supply_values = supply_values.reindex(scenarios, fill_value=0)

        # Rename the row to H2 ratio (0, 10, 20, 100)
        h2_percent = h2_folder.split("_")[1]
        supply_values.name = h2_percent

        # Append to total summary
        total_supply_summary = pd.concat([total_supply_summary, supply_values.to_frame().T])

    # Optional: rename columns and index
    total_supply_summary.index.name = "H2_ratio (%)"
    total_supply_summary.columns.name = "Scenario"

    # ------------------ SAVE TOTAL SUPPLY ------------------
    total_supply_summary.to_excel(supply_file)
    print(f"\n✅ Total supply summary saved successfully:\n{supply_file}\n")

    # ------------------ CREATE SYSTEM-WIDE EFFICIENCY ------------------
    if not total_supply_summary.empty:
        # Reciprocal normalization: base value divided by entire table
        ref_value = total_supply_summary.iloc[0, 0]  # top-left cell
        df_efficiency = ref_value / total_supply_summary
        df_efficiency.to_excel(efficiency_file)
        print(f"\n✅ System-wide efficiency summary saved successfully:\n{efficiency_file}\n")

# ==============================================================
# 📢 EXECUTE TOTAL SUPPLY + SYSTEM EFFICIENCY
# ==============================================================


create_total_supply_summary(year)


# ==============================================================
# 📘 FUNCTION: CREATE SUMMARY OF HYDROGEN MARGINAL COST
# ==============================================================

def create_h2_cost_summary(year):
    """
    Collects the hydrogen marginal cost (Cost_H2) from all H₂ blending
    cases and all heating scenarios, and saves in
    Cost_H2_Summary_<year>.xlsx
    """
    base_path_year = os.path.join(
        project_base,
        "Output",
        main_folder,
        f"Energy_Mix_{year}"
    )

    summary_file = os.path.join(base_path_year, f"Cost_H2_Summary_{year}.xlsx")

    print(f"\n📊 Creating hydrogen marginal cost summary for {year}...")
    print(f"   Searching in: {base_path_year}\n")

    # Initialize DataFrame to store hydrogen marginal cost
    Cost_H2_Summary = pd.DataFrame(index=[v[1] for v in h2_ratios.values()], columns=scenarios)

    for key, (h_folder, h2_value) in h2_ratios.items():
        for scenario in scenarios:
            opgf_file = os.path.join(
                base_path,
                h_folder,
                CCS_fg,
                "Wind_limit_0",
                "ext_grid_0",
                price_fg,
                scenario,
                "Social_MC",              
                investment_scenario,
                "Results",
                f"OPGF_H2_{h2_value/100:.1f}_{investment_scenario}.xlsx"
            )

            if os.path.exists(opgf_file):
                _, _, Cost_H2, _ = create_and_save_operational_cost_breakdown(
                    opgf_file=opgf_file,
                    output_dir=os.path.dirname(opgf_file),
                    price_fg=int(price_fg.split("_")[-1]),
                    CCS_fg=CCS_fg,
                    project_base=project_base,
                    FES_scenario=investment_scenario,
                    verbose=False
                )
                # Store Cost_H2 in the summary table
                Cost_H2_Summary.at[h2_value, scenario] = Cost_H2
            else:
                print(f"⚠ File not found: {opgf_file}")
                Cost_H2_Summary.at[h2_value, scenario] = None

    # Sort by H2 ratio
    Cost_H2_Summary.sort_index(inplace=True)
    Cost_H2_Summary.index.name = "H2_ratio (%)"
    Cost_H2_Summary.columns.name = "Scenario"

    # Save to Excel
    Cost_H2_Summary.to_excel(summary_file)
    print(f"\n✅ Hydrogen marginal cost summary saved successfully:\n{summary_file}\n")

    return Cost_H2_Summary

# ==============================================================
# 📢 EXECUTE H₂ COST SUMMARY
# ==============================================================

Cost_H2_Summary_2050 = create_h2_cost_summary(year)
print(Cost_H2_Summary_2050)


