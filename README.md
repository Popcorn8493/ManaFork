# ManaFork

A collection of Python scripts to help manage your Magic: The Gathering card inventory between Manabox and TCGplayer.

## Overview

ManaFork provides tools for:
- **Converting** Manabox CSV exports to TCGplayer-compatible format
- **Merging** duplicate entries in Manabox inventory files

### Disclaimer

This is a work in progress. The accuracy of the output is not guaranteed, so please verify the results after running the scripts and using the results to implement inventory updates

---

## Setup & Installation

### Installation Steps

1. **Create a virtual environment:**
   ```bash
   python -m venv venv
   ```

2. **Activate the virtual environment:**
   - **Windows:**
     ```bash
     .\venv\Scripts\activate
     ```
   - **macOS/Linux:**
     ```bash
     source venv/bin/activate
     ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

---

## Scripts

### 1. Convert Manabox to TCGplayer (`convert_manabox_tcgp.py`)

This script converts a CSV export from Manabox to a TCGplayer-compatible format.

#### Usage

1. **Get your TCGplayer reference file:**
   - From your TCGplayer seller portal, go to the **Pricing** tab (https://store.tcgplayer.com/admin/pricing)
   - Click **Export Filtered CSV**
   - Leave the options as default, but make sure **Export only from Live Inventory** is unchecked
   - Save this file as `REFERENCE.csv` in the same folder as the script. This generally only needs to be updated as new sets or releases come out. (it has been discovered sometimes it takes awhile for the new cards to properly be added, so generally not a script issue and hopefully cards get added in a timely manner.)

2. **Run the script:**
   ```bash
   python convert_manabox_tcgp.py
   ```

3. **Follow the prompts:**
   - A window will pop up asking you to select your Manabox CSV file
   - The script will try to match each card. You may be prompted to confirm matches:
     - Press **Y** to confirm a match
     - Press **N** to reject it and see the next suggestion
     - Press **G** to give up on a card and move to the next
   - The output will be saved as `tcgplayer_staged.csv`
   - Any cards you gave up on will be in `tcgplayer_given_up.csv`

### 2. Manabox Inventory Merger (`manabox_merger.py`)

A script that merges duplicate entries in Manabox inventory CSV files, consolidating quantities while preserving all card details.

#### Features

- Automatically combines duplicate card entries based on identifying characteristics
- Preserves card details including name, set, condition, foil status, and language
- Sums quantities for identical cards
- Calculates average purchase prices for merged entries
- Provides detailed statistics before and after merging
- Includes data verification to ensure no cards are lost during the process
- User-friendly file selection dialog

#### Usage

1. **Run the script:**
   ```bash
   python manabox_merger.py
   ```

2. **Select your CSV file:** A file dialog will open. Navigate to and select your Manabox inventory CSV file.

3. **Review the output:** The script will display:
   - Original inventory statistics
   - Merged inventory statistics
   - Verification results
   - Preview of merged data

4. **Find your merged file:** The output will be saved as `inventory_merged.csv` in the same directory as the script.

#### Expected CSV Format

The script expects a CSV file with the following columns:
- Name
- Set code
- Collector number
- Language
- Foil
- Condition
- Quantity
- Scryfall ID
- Purchase price
- Altered
- Purchase price currency

#### How It Works

The script identifies duplicate cards by comparing:
- Card name
- Set code
- Collector number
- Language
- Foil status (foil/normal)
- Condition
- Purchase price currency
- Altered status
- Scryfall ID

When duplicates are found, the script:
- Sums the quantities
- Calculates the average purchase price
- Keeps one entry with the combined data

---

## Output Examples

### Merger Output
```
--- ORIGINAL INVENTORY STATISTICS ---
Total cards: 1,247
Foil cards: 89
Normal cards: 1,158
Total entries (rows): 543

--- MERGED INVENTORY STATISTICS ---
Total cards: 1,247
Foil cards: 89
Normal cards: 1,158
Total entries (rows): 398

--- VERIFICATION ---
Total quantities match: True (1247 → 1247)
Foil quantities match: True (89 → 89)
Normal quantities match: True (1158 → 1158)
ALL VERIFICATIONS PASSED - Merge completed successfully!
```



### File Locations

- **Converter outputs:**
  - `tcgplayer_staged.csv` - Successfully matched cards
  - `tcgplayer_given_up.csv` - Cards that couldn't be matched

- **Merger output:**
  - `manabox_inventory_merged.csv` - Merged inventory file

To change output locations, modify the respective filenames in the scripts.

---

## Data Safety

- The scripts never modify your original files
- All changes are saved to new files
- Verification checks ensure no data is lost during processing
- Always backup your original files before processing as a precaution
