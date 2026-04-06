'''
Use this program to calculate quantity discrepancies from D365 and SCI(MIF)
Steps:
1. Run & Download Carton reports from SCI(MIF) including all "P" and "0" cartons
2. Run & Download Carton Discrepancy(Open in Excel) from D365
3. Install the following packages
4. Rename any "⚠️" lines with proper file path & name
Requirements:
1. pip install pandas
2. pip install pandas openpyxl
'''

import pandas as pd

# Gather CSV files from SCI filtered by 'P' and 0

# ⚠️ Import sci report with P package id
sci_p = pd.read_csv( r"C:\Users\FirstInitialLastName\OneDrive - Pacific Sunwear of California, Inc\Desktop\Copy_SCI_P_Name.csv", encoding='cp1252', low_memory=False)

# ⚠️ Import sci report with 00 package id
sci_zero = pd.read_csv( r"C:\Users\FirstInitialLastName\OneDrive - Pacific Sunwear of California, Inc\Desktop\Copy_SCI_0_Name.csv", encoding='cp1252', low_memory=False)

# Sort by package id descending
sci_p = sci_p.sort_values(by='Package ID', ascending=False)
# Sort by package id descending
sci_zero = sci_zero.sort_values(by='Package ID', ascending=False)

# Combine both sci reports all carton types
sci = pd.concat([sci_zero, sci_p], ignore_index=True)

# Convert Package ID to string
sci['Package ID'] = sci['Package ID']
sci_report = sci.rename(columns={'Package ID': 'Carton Number'})

# ⚠️ Import Carton Discrepancy Report
carton_disc = pd.read_excel( r"C:\Users\FirstInitialLastName\OneDrive - Pacific Sunwear of California, Inc\Desktop\CartonDiscrepancy.xlsx", engine="openpyxl")

carton_disc['Carton number'] = carton_disc['Carton number']
carton_disc = carton_disc.rename(columns={'Carton number': 'Carton Number'})

# Find matching cartons in each file

# Convert package id to string for easy comparision
sci_report['Carton Number'] = sci_report['Carton Number'].astype(str).str.lstrip('0')
carton_disc['Carton Number'] = carton_disc['Carton Number'].astype(str).str.lstrip('0')

# Convert to sets for comparison
sci_ids = set(sci_report['Carton Number'])
carton_ids = set(carton_disc['Carton Number'])

# Find matches and differences
both = sci_ids & carton_ids
only_sci = sci_ids - carton_ids
only_carton = carton_ids - sci_ids

# Print summary counts
print(f"{len(both)} Packages in both SCI and Carton Discrepancy Report")
print(f"{len(only_sci)} Packages only in SCI")
print(f"{len(only_carton)} Packages only in Carton Discrepancy Report")

# Check Quantity in Received cartons

# Filter carton_log to only include relevant Notes
carton_disc_recvd = carton_disc[carton_disc['Carton status'].isin(["Acknowledged", "Complete"])]

# Filter sci_report to only include Status 4000
sci_report_4000 = sci_report[sci_report['Package Status'] == 4000]

# Turn the set into a DataFrame
both_df = pd.DataFrame({'Carton Number': list(both)})

# Filter both original DataFrames to keep only Package IDs in 'both'
sci_filtered = sci_report_4000[sci_report_4000['Carton Number'].isin(both)]
carton_filtered = carton_disc_recvd[carton_disc_recvd['Carton Number'].isin(both)]

# Merge the filtered DataFrames on Package ID
merged_both = pd.merge(
    sci_filtered[['Carton Number', "Package Detail Recv'd Qty"]],
    carton_filtered[['Carton Number', 'Quantity shipped']],
    on='Carton Number',
    how='inner'
)

# Sort and reset index for clean display
merged_both = merged_both.sort_values('Carton Number').reset_index(drop=True)

# Drop NaNs and make a new DataFrame
cleaned = merged_both.dropna(subset=["Package Detail Recv'd Qty", "Quantity shipped"]).copy()

# Convert data type to int
cleaned["Package Detail Recv'd Qty"] = cleaned["Package Detail Recv'd Qty"].astype(int)
cleaned["Quantity shipped"] = cleaned["Quantity shipped"].astype(int)

# Compare values
cleaned['Mismatch'] = cleaned["Package Detail Recv'd Qty"] != cleaned["Quantity shipped"]

# Show mismatches
mismatched_df = cleaned[cleaned['Mismatch']]
print(mismatched_df)

# ⚠️ Export to csv
export_path = fr"C:\Users\FirstInitialLastName\OneDrive - Pacific Sunwear of California, Inc\Desktop\CartonDiscReport.csv"
mismatched_df.to_csv(export_path, index=False)

