import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load Data
@st.cache_data
def load_data():
    file_path = "calculated_emissions.xlsx"
    with pd.ExcelFile(file_path) as xls:
        emissions_data = pd.read_excel(xls, sheet_name="Emissions Data")
        scope_summary = pd.read_excel(xls, sheet_name="Scope Summary")
        entity_scope_summary = pd.read_excel(xls, sheet_name="Entity Scope Summary") if "Entity Scope Summary" in xls.sheet_names else None
    return emissions_data, scope_summary, entity_scope_summary

# Load Data
emissions_data, scope_summary, entity_scope_summary = load_data()

# Streamlit Layout
st.title("📊 CO2e Emissions Dashboard")

st.sidebar.header("Filter Data")
selected_year = st.sidebar.multiselect("Select Year", options=emissions_data["Year"].unique(), default=emissions_data["Year"].unique())
selected_scope = st.sidebar.multiselect("Select Scope", options=emissions_data["Scope"].unique(), default=emissions_data["Scope"].unique())

# Filter Data Based on Selection
filtered_data = emissions_data[(emissions_data["Year"].isin(selected_year)) & (emissions_data["Scope"].isin(selected_scope))]

st.subheader("📄 Emissions Data")
st.dataframe(filtered_data)

# Bar Chart: Emissions by Scope
st.subheader("📊 Total Emissions by Scope")
fig, ax = plt.subplots()
sns.barplot(x="Scope", y="Emissions (kg CO2e)", data=scope_summary[:-1], palette="coolwarm", ax=ax)
ax.set_xticklabels(ax.get_xticklabels(), rotation=45)
st.pyplot(fig)

# Pie Chart: Share of Emissions by Scope
st.subheader("🎂 Share of Emissions by Scope")
fig, ax = plt.subplots()
ax.pie(scope_summary[:-1]["Emissions (kg CO2e)"], labels=scope_summary[:-1]["Scope"], autopct="%1.1f%%", colors=sns.color_palette("coolwarm"))
st.pyplot(fig)

# Stacked Bar Chart: Emissions by Entity and Scope
if entity_scope_summary is not None:
    st.subheader("🏢 Emissions by Entity and Scope")
    entity_scope_summary.set_index("Entity").plot(kind="bar", stacked=True, figsize=(10, 6), colormap="coolwarm")
    st.pyplot(plt)

# Download Button
st.subheader("📥 Download Processed Data")
st.download_button("Download as Excel", open("calculated_emissions.xlsx", "rb").read(), "calculated_emissions.xlsx")

st.write("Built with ❤️ using Streamlit")