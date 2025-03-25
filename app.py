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
            "Scope": ["Scope 2"],
            "Country": ["UK"],
            "Category": ["Electricity"],
            "Activity": ["Purchased electricity"],
            "Amount": [10000],
            "Unit": ["kWh"]
        })
        buffer1 = io.BytesIO()
        activity_template.to_excel(buffer1, index=False)
        st.download_button("⬇️ Download Activity Template", buffer1.getvalue(), "activity_data_template.xlsx")

    with col2:
        st.markdown("**📘 Emission Factors Template**")
        factors_template = pd.DataFrame({
            "Year": [2024],
            "Scope": ["Scope 2"],
            "Country": ["UK"],
            "Category": ["Electricity"],
            "Activity": ["Purchased electricity"],
            "Emission Factor (location-based ef)": [0.233],
            "Market-Based EF": [0.160],
            "Unit": ["kgCO2e/kWh"]
        })
        buffer2 = io.BytesIO()
        factors_template.to_excel(buffer2, index=False)
        st.download_button("⬇️ Download Emission Factors Template", buffer2.getvalue(), "emission_factors_template.xlsx")

# --- Upload Section ---
st.sidebar.header("Upload Excel Files")

activity_file = st.sidebar.file_uploader("📅 Upload activity_data.xlsx", type=["xlsx"])
emission_file = st.sidebar.file_uploader("📅 Upload emission_factors.xlsx", type=["xlsx"])

if activity_file and emission_file:
    try:
        # --- Load Emission Factors ---
        emission_df = pd.read_excel(emission_file)
        emission_df.columns = emission_df.columns.str.strip().str.lower()
        emission_df.rename(columns={"emission factor (location-based ef)": "location-based ef"}, inplace=True)

        for col in ["unit", "country", "scope", "category", "activity"]:
            if col in emission_df.columns:
                emission_df[col] = emission_df[col].astype(str).str.strip().str.lower()

        emission_df["location-based ef"] = pd.to_numeric(emission_df["location-based ef"], errors="coerce")
        if "market-based ef" in emission_df.columns:
            emission_df["market-based ef"] = pd.to_numeric(emission_df["market-based ef"], errors="coerce")

        # --- Load Activity Data ---
        activity_df = pd.read_excel(activity_file)
        activity_df.columns = activity_df.columns.str.strip().str.lower()

        for col in ["unit", "country", "scope", "category", "activity"]:
            if col in activity_df.columns:
                activity_df[col] = activity_df[col].astype(str).str.strip().str.lower()

        # --- Validate Required Columns ---
        required_activity_cols = {"year", "scope", "category", "activity", "amount", "unit", "entity"}
        required_factors_cols = {"year", "scope", "category", "activity"}

        missing_act_cols = required_activity_cols - set(activity_df.columns)
        missing_fac_cols = required_factors_cols - set(emission_df.columns)

        if missing_act_cols:
            st.error(f"❌ Missing columns in activity data: {', '.join(missing_act_cols)}")
            st.stop()

        if missing_fac_cols:
            st.error(f"❌ Missing columns in emission factors: {', '.join(missing_fac_cols)}")
            st.stop()

        # --- Separate Scope 2 and Other Scopes ---
        scope2_df = activity_df[activity_df["scope"] == "scope 2"]
        other_scopes_df = activity_df[activity_df["scope"] != "scope 2"]

        # --- Merge Scope 2 with location & market-based EF ---
        scope2_merged = scope2_df.merge(
            emission_df,
            on=["year", "scope", "category", "activity", "country"],
            how="left"
        )
        scope2_merged["location-based"] = scope2_merged["amount"] * scope2_merged["location-based ef"]
        scope2_merged["market-based"] = scope2_merged["amount"] * scope2_merged["market-based ef"]

        # --- Merge other scopes (ignoring unit) ---
        other_merged = other_scopes_df.merge(
            emission_df.drop(columns=["unit"]).drop_duplicates(),
            on=["year", "scope", "category", "activity"],
            how="left"
        )
        other_merged["location-based ef"] = pd.to_numeric(other_merged["location-based ef"], errors="coerce")
        other_merged["emissions (kg co2e)"] = other_merged["amount"] * other_merged["location-based ef"]

        # --- Combine for Location-Based and Market-Based Reporting ---
        def prepare_summary(df, scope2_col):
            df_combined = pd.concat([
                other_merged[["entity", "scope", "emissions (kg co2e)"]],
                scope2_merged[["entity", "scope", scope2_col]].rename(columns={scope2_col: "emissions (kg co2e)"})
            ], ignore_index=True)

            total_by_scope = df_combined.groupby("scope")["emissions (kg co2e)"].sum().reset_index()
            total = total_by_scope["emissions (kg co2e)"].sum()
            total_by_scope = pd.concat([
                total_by_scope,
                pd.DataFrame({"scope": [f"Total GHG Emissions ({scope2_col.replace('-', ' ').title()})"], "emissions (kg co2e)": [total]})
            ], ignore_index=True)

            entity_scope = df_combined.groupby(["entity", "scope"])["emissions (kg co2e)"].sum().reset_index()
            pivot = entity_scope.pivot(index="entity", columns="scope", values="emissions (kg co2e)").fillna(0)
            pivot["Total GHG Emissions"] = pivot.sum(axis=1)

            return total_by_scope, pivot.reset_index()

        location_summary, location_entity_summary = prepare_summary(scope2_merged, "location-based")
        market_summary, market_entity_summary = prepare_summary(scope2_merged, "market-based")

        # --- Display ---
        st.success("✅ Emissions calculated successfully!")

        st.subheader("📋 Location-Based Emissions Summary")
        st.dataframe(location_summary)

        st.subheader("🏢 Location-Based Emissions by Entity")
        st.dataframe(location_entity_summary)

        st.subheader("📋 Market-Based Emissions Summary")
        st.dataframe(market_summary)

        st.subheader("🏢 Market-Based Emissions by Entity")
        st.dataframe(market_entity_summary)

        # --- Export Excel Report ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            location_summary.to_excel(writer, sheet_name="Location Summary", index=False)
            location_entity_summary.to_excel(writer, sheet_name="Location by Entity", index=False)
            market_summary.to_excel(writer, sheet_name="Market Summary", index=False)
            market_entity_summary.to_excel(writer, sheet_name="Market by Entity", index=False)
        st.download_button("📤 Download Excel Report", output.getvalue(), "emissions_summary.xlsx")

        # --- Visualization ---
        def plot_summary_chart(summary_df, title):
            fig, ax = plt.subplots(figsize=(6, 4))
            sns.barplot(data=summary_df[:-1], x="scope", y="emissions (kg co2e)", ax=ax, palette="coolwarm")
            ax.set_title(title, fontsize=14, weight="bold")
            ax.set_ylabel("Emissions (kg CO₂e)")
            ax.set_xlabel("")
            st.pyplot(fig)

        st.subheader("📊 Emissions by Scope (Location-Based)")
        plot_summary_chart(location_summary, "Total Emissions by Scope (Location-Based)")

        st.subheader("📊 Emissions by Scope (Market-Based)")
        plot_summary_chart(market_summary, "Total Emissions by Scope (Market-Based)")

    except Exception as e:
        st.error(f"❌ Something went wrong: {e}")

else:
    st.info("📈 Upload both Activity Data and Emission Factors to begin.")
