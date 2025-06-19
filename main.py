from relevanceai import RelevanceAI
from excel_utils import read_excel, append_to_excel
from config import API_KEY, PROJECT, REGION
import pandas as pd

# Set file paths
SOURCE = "source.xlsx"
TARGET = "target.xlsx"

# Auth
client = RelevanceAI(api_key=API_KEY, project=PROJECT, region=REGION)

# Step 1: Load Excel data
df = read_excel(SOURCE)

# Step 2: Convert rows to dict format
documents = df.to_dict(orient="records")

# Step 3: Use filtering logic (e.g., only Status == Pending)
filtered = [doc for doc in documents if doc.get("Status") == "Pending"]

# Step 4: Save to Excel
filtered_df = pd.DataFrame(filtered)
if append_to_excel(filtered_df, TARGET):
    print(f"✅ Copied {len(filtered)} rows where Status = 'Pending' to {TARGET}")
else:
    print("❌ Failed to copy rows to target file. Please check the error message above.")
