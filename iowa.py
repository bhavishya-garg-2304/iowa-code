import pdfplumber
import pandas as pd
import re

pdf_path = "June2025.pdf"

with pdfplumber.open(pdf_path) as pdf:
    text = pdf.pages[0].extract_text()

# Step 2: Remove report title and split lines
lines = text.split("\n")
lines = [line.strip() for line in lines if "GAMING REVENUE REPORT" not in line and line.strip() != ""]

# Step 3: Extract casino names (appear before "ADJUSTED GROSS REVENUE")
casino_names = []
i = 0
while i < len(lines):
    if "ADJUSTED GROSS REVENUE" in lines[i]:
        break
    # Combine two lines to form full name
    name = lines[i]
    if i + 1 < len(lines) and not any(c.isdigit() for c in lines[i + 1]):
        name += " " + lines[i + 1]
        i += 1
    casino_names.append(name.strip(" -"))
    i += 1

# Step 4: Parse data rows starting from "ADJUSTED GROSS REVENUE"
metrics = []
data = []
while i < len(lines):
    label = lines[i]
    i += 1
    if i >= len(lines):
        break
    values = re.split(r'\s{2,}', lines[i])
    if len(values) == len(casino_names):
        metrics.append(label)
        data.append(values)
    i += 1

# Step 5: Build DataFrame
df = pd.DataFrame(data, columns=casino_names)
df.insert(0, "METRIC", metrics)

# Step 6: Export to Excel
df.to_excel("gaming_revenue_june2025_final.xlsx", index=False)
print("✅ Saved as 'gaming_revenue_june2025_final.xlsx'")
