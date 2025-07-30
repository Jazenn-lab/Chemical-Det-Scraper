🧪 Fresh Chemical Scraper
A Python tool to auto-fill missing chemical details from an Excel sheet using CAS numbers. It fetches data from PubChem and Ambeed, guesses the chemical category, and saves the enriched results into a structured Excel file.


🔧 Features

✅ Auto-fetch Chemical Name, Molecular Formula, Molecular Weight, Appearance, etc.

✅ Intelligent category guessing (e.g. Aromatic, Halogenated)

✅ Multi-threaded for speed 🚀

✅ Built-in retry logic for unreliable responses

✅ Custom start point and file path set directly in code

✅ Logging to console and optional log file

✅ Clean and modular code


📁 Requirements

Python 3.7+

Dependencies listed in requirements.txt

Install them using:
pip install -r requirements.txt

Usage
1. Edit the input settings:
Open the Python file and set these two variables at the top:
INPUT_FILE = r"path_to_your_excel_file.xlsx"
CUSTOM_START_INDEX = 0 # start from this row (0-based index)

2. Ensure the Excel file has at least one column named:
CAS No
(You can optionally include a column called Chemical Name, otherwise fallback methods will be
used.)

3. Run the script:
python fresh_scraper.py
The script will generate:
- Fresh_Chemical_Database.xlsx - output file with enriched data.
- progress.json - saves current progress in case of interruption.
- Optional console log with each CAS processed.

📌 Output Fields
Product Code (e.g., S1-0001)

Chemical Name

CAS No

Synonyms

Molecular Formula

Molecular Weight

Appearance

Storage

Shipping Conditions

Applications

Category (e.g. Aromatic, Aliphatic, Impurity...)

💡 Notes
This version does not include structure image downloading — but it's easily extendable.

Automatically saves progress every few rows to avoid data loss.

You can plug in additional APIs or local databases as needed.

📜 License
MIT License
