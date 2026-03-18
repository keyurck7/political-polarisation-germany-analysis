import pandas as pd

# === File paths ===
votes_path = "Partyvoteshare.xlsx"
template_path = "Germany_Election_Master_Dataset_Template (1).xlsx"
output_path = "Germany_Election_Master_Filled_2021.xlsx"

# === Read both files ===
votes_df = pd.read_excel(votes_path)
template_df = pd.read_excel(template_path, sheet_name="Master_Template")

# === Add Year column ===
votes_df["Year"] = 2021

# === Reorder & fill missing columns to match master ===
final_cols = [
    "State", "Year",
    "CDU/CSU%", "SPD%", "Greens%", "FDP%", "AfD%", "Left%", "Others%",
    "Turnout%", "GDP per capita (€)", "Unemployment%", "Median Income (€)",
    "Foreign-born%", "Education% (Tertiary)", "Urbanisation%",
    "Age <30%", "Age 30-60%", "Age >60%"
]

for col in final_cols:
    if col not in votes_df.columns:
        votes_df[col] = None

votes_df = votes_df[final_cols]

# === Replace the master template rows ===
votes_df.to_excel(output_path, sheet_name="Master_Template", index=False)

print(f"✅ Done! Merged 2021 data saved as '{output_path}'")
