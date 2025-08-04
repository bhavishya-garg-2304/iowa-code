import pdfplumber
import pandas as pd
import re

pdf_path = "June2025.pdf"

with pdfplumber.open(pdf_path) as pdf:
    text = ""
    for page in pdf.pages:
        t = page.extract_text()
        if "GAMING REVENUE REPORT -- JUNE 2025" in t:
            text += t
            break  # assuming first page contains the full table

# Step 1: Extract lines
lines = text.split("\n")

# Step 2: Build casino column headers
casino_names = []
i = 0
while i < len(lines):
    if "ADJUSTED GROSS REVENUE" in lines[i]:
        break
    line1 = lines[i].strip()
    line2 = lines[i+1].strip() if i+1 < len(lines) else ""
    # Merge multi-line casino names
    if not any(char.isdigit() for char in line1 + line2):
        full_name = f"{line1} {line2}".strip(" -")
        casino_names.append(full_name)
        i += 2
    else:
        i += 1

# Step 3: Extract metric rows
metrics = []
data = []

while i < len(lines):
    line = lines[i].strip()
    if re.match(r'^[A-Z ]+$', line):  # all caps: it's a metric name
        metric_name = line
        i += 1
        if i < len(lines):
            values = re.split(r'\s{2,}', lines[i].strip())
            if len(values) == len(casino_names):
                metrics.append(metric_name)
                data.append(values)
    i += 1

# Step 4: Create DataFrame
df = pd.DataFrame(data, columns=casino_names)
df.insert(0, "METRIC", metrics)

# Step 5: Export
df.to_excel("gaming_revenue_june2025_final.xlsx", index=False)
print("✅ Exported to 'gaming_revenue_june2025_final.xlsx'")
