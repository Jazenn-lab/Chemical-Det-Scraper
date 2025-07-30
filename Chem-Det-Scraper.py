import os
import json
import pandas as pd
import requests
import time
import logging
from bs4 import BeautifulSoup
from functools import wraps
from concurrent.futures import ThreadPoolExecutor, as_completed

# === CONFIGURATION ===
INPUT_FILE = "New_Master_Product_Catalogue.xlsx"
OUTPUT_FILE = "Fresh_Chemical_Database.xlsx"
PROGRESS_FILE = "progress.json"
CUSTOM_START_INDEX = 2800  # Set to an integer to override resume progress
SAVE_EVERY_N_ROWS = 5

# === Logging Setup ===
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("scraper.log"),
        logging.StreamHandler()
    ]
)

# === Retry Decorator ===
def retry(max_attempts=3, delay=2, backoff=2):
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            attempts = 0
            current_delay = delay
            while attempts < max_attempts:
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    attempts += 1
                    logging.warning(f"[{func.__name__}] Retry {attempts}/{max_attempts} due to: {e}")
                    time.sleep(current_delay)
                    current_delay *= backoff
            logging.error(f"[{func.__name__}] ❌ All retries failed")
            return None
        return wrapper
    return decorator

# === Utility Functions ===
def guess_general_category(chemical_name):
    if not chemical_name or pd.isna(chemical_name):
        return "Impurity"

    name = str(chemical_name).lower()

    category_map = {
        "Heterocyclic": ["triazole", "pyrazole", "imidazole", "thiazole", "oxazole",
                          "pyridine", "quinoline", "isoquinoline", "indole", "pyrimidine",
                          "piperidine", "piperazine", "morpholine", "azetidine"],
        "Aromatic": ["benzene", "phenyl", "toluene", "naphthalene", "aromatic", "indole"],
        "Aliphatic": ["methyl", "ethyl", "propyl", "butyl", "pentyl", "hexyl", "cycloalkyl"],
        "Halogenated": ["chloro", "fluoro", "bromo", "iodo", "halide"],
        "Nitrogen-Containing": ["amine", "amino", "urea", "azide", "azo", "nitrile", "hydrazine"],
        "Oxygen-Containing": ["alcohol", "ether", "aldehyde", "ketone", "acid", "ester", "anhydride"],
        "Sulfur-Containing": ["thiol", "thio", "sulfonamide", "thiazole", "sulfide"],
        "Phosphorus-Containing": ["phosphate", "phosphonate", "phosphine"],
        "Steroidal": ["steroid", "estradiol", "testosterone", "corticosteroid"]
    }

    for category, keywords in category_map.items():
        for kw in keywords:
            if kw in name:
                return category
    return "Impurity"

@retry(max_attempts=3)
def fetch_pubchem(cas):
    url = f"https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/{cas}/JSON"
    res = requests.get(url, timeout=10)
    if res.status_code != 200:
        raise Exception(f"PubChem status {res.status_code}")
    data = res.json()
    compound = data.get("PC_Compounds", [])[0]

    ids = compound.get("id", {}).get("id", {})
    name = ids.get("name", cas)

    formula, weight, appearance = "", "", ""
    for prop in compound.get("props", []):
        label = prop.get("urn", {}).get("label", "")
        val = prop.get("value", {})
        sval, fval = val.get("sval"), val.get("fval")

        if label == "Molecular Formula" and sval:
            formula = sval
        elif label == "Molecular Weight":
            weight = sval or str(fval)
        elif label == "Appearance" and sval:
            appearance = sval

    return {
        "Chemical Name": name,
        "Molecular Formula": formula,
        "Molecular Weight": weight,
        "Appearance": appearance
    }

@retry(max_attempts=3)
def fetch_category(cas):
    try:
        url = f"https://www.ambeed.com/products/{cas}.html"
        res = requests.get(url, timeout=10)
        soup = BeautifulSoup(res.text, "html.parser")
        all_links = soup.select("a.moreicon")
        valid_links = [a.text.strip().split("(")[0].strip() for a in all_links if a.text.strip() and len(a.text.strip()) < 30]
        return valid_links[-1] if valid_links else "Impurity"
    except Exception as e:
        raise Exception(f"Failed to fetch Ambeed category for {cas}: {e}")

def generate_applications(name):
    return "Used in chemical synthesis, pharmaceutical or industrial research."

def process_row(index, cas, chemical_name):
    logging.info(f"[{index+1}] Processing {cas}")
    pubchem_data = fetch_pubchem(cas) or {"Chemical Name": cas, "Molecular Formula": "", "Molecular Weight": "", "Appearance": ""}
    category = guess_general_category(chemical_name)
    return {
        "Product Code": f"S1-{str(index+1).zfill(4)}",
        "Chemical Name": chemical_name,
        "CAS No": cas,
        "Synonyms": chemical_name,
        "Molecular Formula": pubchem_data["Molecular Formula"],
        "Molecular Weight": pubchem_data["Molecular Weight"],
        "Appearance": pubchem_data["Appearance"],
        "Storage": "2-8°C Refrigerator",
        "Shipping Conditions": "Ambient",
        "Applications": generate_applications(chemical_name),
        "Category": category
    }

def main():
    input_df = pd.read_excel(INPUT_FILE)
    cas_list = input_df["CAS No"].dropna().astype(str).tolist()

    start_index = CUSTOM_START_INDEX if CUSTOM_START_INDEX is not None else 0
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE) as f:
            progress = json.load(f)
            start_index = progress.get("last_row", start_index)

    output_data = []
    if os.path.exists(OUTPUT_FILE):
        output_data = pd.read_excel(OUTPUT_FILE).to_dict(orient='records')

    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = []
        for i in range(start_index, len(cas_list)):
            cas = cas_list[i].strip()
            chemical_name_row = input_df.loc[input_df["CAS No"].astype(str).str.strip() == cas, "Chemical Name"]
            chemical_name = chemical_name_row.values[0] if not chemical_name_row.empty else cas
            futures.append(executor.submit(process_row, i, cas, chemical_name))

        for count, future in enumerate(as_completed(futures), start=start_index + 1):
            result = future.result()
            output_data.append(result)
            if count % SAVE_EVERY_N_ROWS == 0:
                pd.DataFrame(output_data).to_excel(OUTPUT_FILE, index=False)
                with open(PROGRESS_FILE, "w") as f:
                    json.dump({"last_row": count}, f)
                logging.info(f"💾 Auto-saved at row {count}")

    pd.DataFrame(output_data).to_excel(OUTPUT_FILE, index=False)
    with open(PROGRESS_FILE, "w") as f:
        json.dump({"last_row": len(cas_list)}, f)
    logging.info(f"✅ Done! Final file saved to: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()