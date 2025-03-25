import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import io

sns.set_theme(style="whitegrid", palette="coolwarm")

st.set_page_config(page_title="CO₂e Emissions Dashboard", layout="wide")
st.title("CO₂e Emissions Dashboard")

# --- Downloadable Templates ---
with st.expander("📂 Download Excel Templates"):
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**📘 Activity Data Template**")
        activity_template = pd.DataFrame({
            "Year": [2024],
            "Entity": ["Company A"],
            "Scope": ["Scope 1"],
            "Category": ["Fuel"],
            "Activity": ["Natural gas"],
            "Amount": [1000],
            "Unit": ["kWh"]
        })
        buffer1 = io.BytesIO()
        activity_template.to_excel(buffer1, index=False)
        st.download_button("⬇️ Download Activity Template", buffer1.getvalue(), "activity_data_template.xlsx")

    with col2:
        st.markdown("**📘 Emission Factors Template**")
        factors_template = pd.DataFrame({
            "Year": [2024],
            "Scope": ["Scope 1"],
            "Category": ["Fuel"],
            "Activity": ["Natural gas"],
            "Emission factor": [0.183],
            "Unit": ["kgCO2e/kWh"]
        })
        buffer2 = io.BytesIO()
        factors_template.to_excel(buffer2, index=False)
        st.download_button("⬇️ Download Emission Factors Template", buffer2.getvalue(), "emission_factors_template.xlsx")


# --- Upload Section ---
st.sidebar.header("Upload Excel Files")

activity_file = st.sidebar.file_uploader("📥 Upload activity_data.xlsx", type=["xlsx"])
emission_file = st.sidebar.file_uploader("📥 Upload emission_factors.xlsx", type=["xlsx"])

if activity_file and emission_file:
    try:
        # --- Load Emission Factors ---
        emission_df = pd.read_excel(emission_file)
        emission_df.columns = emission_df.columns.str.strip().str.lower()
        emission_df["unit"] = emission_df["unit"].str.strip().str.lower()
        emission_df["base unit"] = emission_df["unit"].apply(lambda x: x.split("/")[-1])

        # --- Load Activity Data ---
        activity_df = pd.read_excel(activity_file)
        activity_df.columns = activity_df.columns.str.strip().str.lower()
        activity_df["unit"] = activity_df["unit"].str.strip().str.lower()

        # --- Validate Required Columns ---
        required_activity_cols = {"year", "scope", "category", "activity", "amount", "unit"}
        required_factors_cols = {"year", "scope", "category", "activity", "emission factor", "unit"}

        missing_act_cols = required_activity_cols - set(activity_df.columns)
        missing_fac_cols = required_factors_cols - set(emission_df.columns)

        if missing_act_cols:
            st.error(f"❌ Missing columns in activity data: {', '.join(missing_act_cols)}")
            st.stop()

        if missing_fac_cols:
            st.error(f"❌ Missing columns in emission factors: {', '.join(missing_fac_cols)}")
            st.stop()

        # --- Merge ---
        merged_df = activity_df.merge(
            emission_df,
            left_on=["year", "scope", "category", "activity", "unit"],
            right_on=["year", "scope", "category", "activity", "base unit"],
            how="left"
        ).drop(columns=["base unit"])

        # --- Calculate Emissions ---
        merged_df["emissions (kg co2e)"] = merged_df["amount"] * merged_df["emission factor"]

        # --- Summarize by Scope ---
        scope_summary = merged_df.groupby("scope")["emissions (kg co2e)"].sum().reset_index()
        total_row = pd.DataFrame({
            "scope": ["Total GHG Emissions"],
            "emissions (kg co2e)": [scope_summary["emissions (kg co2e)"].sum()]
        })
        scope_summary = pd.concat([scope_summary, total_row], ignore_index=True)

        # --- Summarize by Entity + Scope ---
        if "entity" in merged_df.columns:
            entity_summary_raw = merged_df.groupby(["entity", "scope"])["emissions (kg co2e)"].sum().reset_index()
            entity_summary_pivot = entity_summary_raw.pivot(index="entity", columns="scope", values="emissions (kg co2e)").fillna(0)
            entity_summary_pivot["total ghg emissions"] = entity_summary_pivot.sum(axis=1)
            entity_summary = entity_summary_pivot.reset_index()
        else:
            entity_summary = None

        # --- Display Results ---
        st.success("✅ Emissions calculated successfully!")

        st.subheader("📄 Emissions Data")
        st.dataframe(merged_df)

        st.subheader("📋 Scope Summary")
        st.dataframe(scope_summary)

        if entity_summary is not None:
            st.subheader("🏢 Entity Scope Summary")
            st.dataframe(entity_summary)

        # --- Visualizations ---
        st.subheader("📊 Emissions by Scope")
        fig, ax = plt.subplots(figsize=(8, 5))
        sns.barplot(data=scope_summary[:-1], x="scope", y="emissions (kg co2e)", ax=ax)
        ax.set_title("Total Emissions by Scope", fontsize=16, fontweight="bold", color="#0A3A5C")
        ax.set_xlabel("Scope", fontsize=12)
        ax.set_ylabel("Emissions (kg CO₂e)", fontsize=12)
        ax.tick_params(axis='x', labelrotation=45)
        sns.despine()

        st.subheader("🎂 Share of Emissions by Scope")
        fig2, ax2 = plt.subplots()
        ax2.pie(scope_summary[:-1]["emissions (kg co2e)"], labels=scope_summary[:-1]["scope"], autopct="%1.1f%%", colors=sns.color_palette("coolwarm"))
        st.pyplot(fig2)

        if entity_summary is not None:
            st.subheader("🏢 Emissions by Entity and Scope")
            fig3, ax3 = plt.subplots(figsize=(10, 6))
            entity_summary.set_index("entity").drop(columns=["total ghg emissions"]).plot(kind="bar", stacked=True, ax=ax3, colormap="coolwarm")
            ax3.set_ylabel("Emissions (kg CO2e)")
            ax3.set_title("Entity-Level Emissions by Scope")
            st.pyplot(fig3)

        # --- Download Button ---
        st.subheader("📥 Download Calculated Emissions")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            merged_df.to_excel(writer, sheet_name="Emissions Data", index=False)
            scope_summary.to_excel(writer, sheet_name="Scope Summary", index=False)
            if entity_summary is not None:
                entity_summary.to_excel(writer, sheet_name="Entity Scope Summary", index=False)
        st.download_button("📤 Download Excel Report", output.getvalue(), "calculated_emissions.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"❌ Something went wrong: {e}")

else:
    st.info("👈 Upload both Activity Data and Emission Factors to begin.")
