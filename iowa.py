import pdfplumber
import pandas as pd
from collections import defaultdict
import math

# Settings
pdf_path = "June2025.pdf"
metric_col_x0 = 41
metric_col_x1 = 155
casino_col_width = 86
first_casino_x0 = 155
num_casinos = 7
pages_to_extract = [0, 1]  # Page 1 and Page 2 (0-indexed)

def get_column_index(x0):
    return int((x0 - first_casino_x0) // casino_col_width)

# Function to extract casino names based on x0 positions
def extract_casino_names(page):
    words = page.extract_words()
    name_blocks = defaultdict(list)

    for w in words:
        if w['x0'] >= first_casino_x0 and not any(char.isdigit() for char in w['text']):
            col_idx = get_column_index(w['x0'])
            if 0 <= col_idx < num_casinos:
                name_blocks[col_idx].append((w['top'], w['text']))

    # Join words in order of appearance (top-to-bottom)
    casino_names = []
    for i in range(num_casinos):
        lines = sorted(name_blocks[i], key=lambda x: x[0])
        name = ' '.join(text for _, text in lines).replace(' -', '').strip()
        casino_names.append(name)
    return casino_names

# Main extraction loop
all_rows = []

with pdfplumber.open(pdf_path) as pdf:
    for page_num in pages_to_extract:
        page = pdf.pages[page_num]
        words = page.extract_words(use_text_flow=True)

        # Get casino names dynamically
        casino_names = extract_casino_names(page)

        # Group words by Y position
        line_map = defaultdict(list)
        for word in words:
            y = round(word["top"], 1)
            line_map[y].append(word)

        for y in sorted(line_map):
            row = [''] * (num_casinos + 1)  # one for metric, rest for casinos
            for word in sorted(line_map[y], key=lambda w: w['x0']):
                x0 = word['x0']
                text = word['text']
                if metric_col_x0 <= x0 < metric_col_x1:
                    row[0] += text + ' '
                elif x0 >= first_casino_x0:
                    col_idx = get_column_index(x0)
                    if 0 <= col_idx < num_casinos:
                        row[col_idx + 1] += text + ' '
            row = [cell.strip() for cell in row]
            if any(row[1:]):  # skip empty data lines
                all_rows.append((tuple(casino_names), row))

# Separate tables by casino name sets
tables = defaultdict(list)
for header, row in all_rows:
    tables[header].append(row)

# Save each distinct header set (i.e., table) to one DataFrame
with pd.ExcelWriter("gaming_revenue_pg1_pg2.xlsx") as writer:
    for i, (casino_headers, rows) in enumerate(tables.items()):
        df = pd.DataFrame(rows, columns=["METRIC"] + list(casino_headers))
        df.to_excel(writer, sheet_name=f"Page{i+1}", index=False)

print("✅ Exported 'gaming_revenue_pg1_pg2.xlsx' with real casino names.")
