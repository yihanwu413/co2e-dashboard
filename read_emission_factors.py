import pandas as pd

def load_emission_factors(file_path):
    """Reads an Excel file and returns a DataFrame of emission factors."""
    df = pd.read_excel(file_path)
    return df

# Test the function
file_path = "emission_factors.xlsx"
emission_factors_df = load_emission_factors(file_path)

print(emission_factors_df)
