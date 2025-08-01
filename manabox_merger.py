import pandas as pd
import tkinter as tk
from tkinter import filedialog


# Create a root window and hide it
root = tk.Tk()
root.withdraw()

# Open file dialog to select CSV file
csv_file = filedialog.askopenfilename(
    title="Select CSV file to merge",
    filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
)

# Check if user cancelled the dialog
if not csv_file:
    print("No file selected. Exiting...")
    exit()

# Load the CSV file, using the first row as the header
try:
    df = pd.read_csv(csv_file, header=0)

    # --- Data Cleaning and Preparation ---

    # Convert 'Quantity' and 'Purchase price' to numeric types.
    # Errors will be converted to NaN (Not a Number)
    df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce')
    df['Purchase price'] = pd.to_numeric(df['Purchase price'], errors='coerce')

    # Fill NaN values in 'Quantity' and 'Purchase price' with 0.
    # This prevents errors during aggregation.
    df['Quantity'] = df['Quantity'].fillna(0)
    df['Purchase price'] = df['Purchase price'].fillna(0)

    # Fill NaN values in the 'Altered' column with a placeholder string 'No'
    # to ensure they are grouped correctly.
    df['Altered'] = df['Altered'].fillna('No')

    # --- Aggregation ---

    # Define the columns that identify a unique item
    grouping_cols = [
        'Name',
        'Set code',
        'Collector number',
        'Language',
        'Foil',
        'Condition',
        'Purchase price currency',
        'Altered',
        'Scryfall ID'
    ]

    # --- Pre-merge Statistics ---
    original_total_quantity = df['Quantity'].sum()
    original_foil_quantity = df[df['Foil'] == 'foil']['Quantity'].sum()
    original_normal_quantity = df[df['Foil'] != 'foil']['Quantity'].sum()
    
    print("\n--- ORIGINAL INVENTORY STATISTICS ---")
    print(f"Total cards: {original_total_quantity}")
    print(f"Foil cards: {original_foil_quantity}")
    print(f"Normal cards: {original_normal_quantity}")
    print(f"Total entries (rows): {len(df)}")

    # Group by the identifying columns and aggregate quantity and purchase price
    merged_df = df.groupby(grouping_cols, as_index=False).agg({
        'Quantity': 'sum',
        'Purchase price': 'mean'  # Use average purchase price when merging
    })

    # --- Post-merge Statistics ---
    merged_total_quantity = merged_df['Quantity'].sum()
    merged_foil_quantity = merged_df[merged_df['Foil'] == 'foil']['Quantity'].sum()
    merged_normal_quantity = merged_df[merged_df['Foil'] != 'foil']['Quantity'].sum()

    print("\n--- MERGED INVENTORY STATISTICS ---")
    print(f"Total cards: {merged_total_quantity}")
    print(f"Foil cards: {merged_foil_quantity}")
    print(f"Normal cards: {merged_normal_quantity}")
    print(f"Total entries (rows): {len(merged_df)}")

    # --- Verification ---
    print("\n--- VERIFICATION ---")
    total_match = original_total_quantity == merged_total_quantity
    foil_match = original_foil_quantity == merged_foil_quantity
    normal_match = original_normal_quantity == merged_normal_quantity
    
    print(f"✓ Total quantities match: {total_match} ({original_total_quantity} → {merged_total_quantity})")
    print(f"✓ Foil quantities match: {foil_match} ({original_foil_quantity} → {merged_foil_quantity})")
    print(f"✓ Normal quantities match: {normal_match} ({original_normal_quantity} → {merged_normal_quantity})")
    
    if total_match and foil_match and normal_match:
        print("[✓] ALL VERIFICATIONS PASSED - Merge completed successfully!")
    else:
        print("[✕] VERIFICATION FAILED - Quantities do not match!")

    # The resulting merged_df will have the summed quantities.
    # Let's reorder columns to match the original format as closely as possible.
    output_cols = [
        'Name',
        'Set code',
        'Collector number',
        'Language',
        'Foil',
        'Condition',
        'Quantity',
        'Scryfall ID',
        'Purchase price',
        'Altered',
        'Purchase price currency'
    ]
    # Filter out any columns that might not exist if the input changes
    # This makes the code more robust
    final_cols = [col for col in output_cols if col in merged_df.columns]
    merged_df = merged_df[final_cols]

    # Save the merged data to a new CSV file
    output_filename = 'manabox_inventory_merged.csv'
    merged_df.to_csv(output_filename, index=False)

    print(f"\nSuccessfully merged the data and saved it to '{output_filename}'.")
    print("\nFirst 5 rows of the merged data:")
    print(merged_df.head())

except FileNotFoundError:
    print("Error: 'inventory (3).csv' not found. Please ensure the file is in the correct directory.")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
