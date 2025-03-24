import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load Emission Factors
def load_emission_factors(file_path):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()  # Remove spaces from column names
    df["Unit"] = df["Unit"].str.strip().str.lower()  # Standardize unit formatting
    df["Base Unit"] = df["Unit"].apply(lambda x: x.split("/")[-1])  # Extract base unit
    return df

# Load Activity Data
def load_activity_data(file_path):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()  # Remove spaces from column names
    df["Unit"] = df["Unit"].str.strip().str.lower()  # Standardize unit formatting
    return df

# Calculate Emissions
def calculate_emissions(activity_file, emission_file):
    activity_df = load_activity_data(activity_file)
    emission_df = load_emission_factors(emission_file)

    merged_df = activity_df.merge(
        emission_df,  
        left_on=["Year", "Scope", "Category", "Activity", "Unit"],
        right_on=["Year", "Scope", "Category", "Activity", "Base Unit"],
        how="left"
    ).drop(columns=["Base Unit"])  # Remove extra column after merging

    # Calculate Emissions
    merged_df["Emissions (kg CO2e)"] = merged_df["Amount"] * merged_df["Emission factor"]

    return merged_df

# Generate Summary Report
def generate_summary(result_df):
    scope_summary = result_df.groupby("Scope")["Emissions (kg CO2e)"].sum().reset_index()

    # Add total GHG emissions row
    total_emissions = scope_summary["Emissions (kg CO2e)"].sum()
    total_row = pd.DataFrame({"Scope": ["Total GHG Emissions"], "Emissions (kg CO2e)": [total_emissions]})
    scope_summary = pd.concat([scope_summary, total_row], ignore_index=True)

    if "Entity" in result_df.columns:
        entity_scope_summary = result_df.groupby(["Entity", "Scope"])["Emissions (kg CO2e)"].sum().reset_index()

        # Pivot for better visualization
        entity_pivot = entity_scope_summary.pivot(index="Entity", columns="Scope", values="Emissions (kg CO2e)").fillna(0)

        # Add total emissions column
        entity_pivot["Total GHG Emissions"] = entity_pivot.sum(axis=1)
        entity_scope_summary = entity_pivot.reset_index()
    else:
        entity_scope_summary = None

    return scope_summary, entity_scope_summary

# Generate and Save Plots
def generate_visualizations(scope_summary, entity_scope_summary):
    # Bar chart - Total emissions by scope
    plt.figure(figsize=(8, 5))
    sns.barplot(x="Scope", y="Emissions (kg CO2e)", data=scope_summary[:-1], palette="coolwarm")  # Exclude 'Total'
    plt.title("Total Emissions by Scope")
    plt.xticks(rotation=45)
    plt.savefig("emissions_by_scope.png")
    plt.close()

    # Pie chart - Share of emissions by scope
    plt.figure(figsize=(6, 6))
    plt.pie(scope_summary[:-1]["Emissions (kg CO2e)"], labels=scope_summary[:-1]["Scope"], autopct="%1.1f%%", colors=sns.color_palette("coolwarm"))
    plt.title("Share of Emissions by Scope")
    plt.savefig("emissions_pie_chart.png")
    plt.close()

    # Stacked bar chart - Emissions by entity and scope
    if entity_scope_summary is not None:
        entity_scope_summary.set_index("Entity").plot(kind="bar", stacked=True, figsize=(10, 6), colormap="coolwarm")
        plt.title("Emissions by Entity and Scope")
        plt.xlabel("Entity")
        plt.ylabel("Emissions (kg CO2e)")
        plt.xticks(rotation=45)
        plt.legend(title="Scope")
        plt.savefig("emissions_by_entity.png")
        plt.close()

# Main Execution
activity_file = "activity_data.xlsx"
emission_file = "emission_factors.xlsx"

result_df = calculate_emissions(activity_file, emission_file)
scope_summary, entity_scope_summary = generate_summary(result_df)

# Save results
with pd.ExcelWriter("calculated_emissions.xlsx", engine="openpyxl") as writer:
    result_df.to_excel(writer, sheet_name="Emissions Data", index=False)
    scope_summary.to_excel(writer, sheet_name="Scope Summary", index=False)
    if entity_scope_summary is not None:
        entity_scope_summary.to_excel(writer, sheet_name="Entity Scope Summary", index=False)

# Generate and save visualizations
generate_visualizations(scope_summary, entity_scope_summary)

print("Emissions calculated and summary report added to 'calculated_emissions.xlsx'")
print("Graphs saved as PNG images in the project folder.")