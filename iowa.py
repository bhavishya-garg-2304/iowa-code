pdf_path = "June2025.pdf"
metric_col_x0 = 41
metric_col_x1 = 155
casino_col_width = 86
first_casino_x0 = 155
num_casinos = 7
pages_to_extract = [0, 1]

def get_column_index(x0):
    return int((x0 - first_casino_x0) // casino_col_width)

# Step 1: Extract casino headers per page
def extract_casino_names(page):
    casino_names = []

    for i in range(num_casinos):
        # Step 1: Pull all candidate words in this column
        x0 = first_casino_x0 + i * casino_col_width
        x1 = x0 + casino_col_width
        all_words = page.extract_words()
        col_words = [w for w in all_words if x0 <= w['x0'] <= x1 and not any(char.isdigit() for char in w['text'])]

        if not col_words:
            print(f"[WARN] No words found for column {i+1}")
            casino_names.append(f"Casino {i+1}")
            continue

        # Step 2: Find top-most Y position
        min_top = min(w["top"] for w in col_words)
        name_y0 = max(min_top - 2, 0)  # small margin above
        name_y1 = name_y0 + 45  # slightly expanded height

        cropped = page.within_bbox((x0, name_y0, x1, name_y1))
        name_words = cropped.extract_words()

        if not name_words:
            print(f"[WARN] No words extracted in bbox for column {i+1} (y0={name_y0:.1f}, y1={name_y1:.1f})")
            casino_names.append(f"Casino {i+1}")
            continue

        name_sorted = sorted(name_words, key=lambda w: (w["top"], w["x0"]))
        text = ' '.join([w["text"] for w in name_sorted])
        text = text.replace(" -", "").strip()
        text = re.sub(r"\s+", " ", text)

        print(f"[INFO] Casino {i+1} name: '{text}'")
        casino_names.append(text)

    return casino_names

# Step 2: Extract all data rows with values
def extract_data_rows(page, casino_names):
    words = page.extract_words(use_text_flow=True)

    line_map = defaultdict(list)
    for word in words:
        y = round(word["top"], 1)
        line_map[y].append(word)

    data_rows = []
    for y in sorted(line_map):
        row = [''] * (num_casinos + 1)
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

        # Keep only rows that have at least one value (excluding metric)
        if any(row[1:]):
            data_rows.append(row)
    return data_rows

# Step 3: Clean numeric values
def clean_numeric(val):
    val = val.replace("$", "").replace(",", "").replace("(", "-").replace(")", "")
    try:
        return float(val)
    except:
        return val  # fallback to string if not a number

# Master list of all rows
combined_rows = []
header = None

# Process both pages
with pdfplumber.open(pdf_path) as pdf:
    for page_num in pages_to_extract:
        page = pdf.pages[page_num]
        casino_names = extract_casino_names(page)
        data_rows = extract_data_rows(page, casino_names)

        # Set column header once
        if header is None:
            header = ["METRIC"] + casino_names

        combined_rows.extend(data_rows)

# Create DataFrame
df = pd.DataFrame(combined_rows, columns=header)

# Clean numeric columns (all except METRIC)
for col in df.columns[1:]:
    df[col] = df[col].apply(clean_numeric)

# Export to Excel
df.to_excel("gaming_revenue_combined_cleaned.xlsx", index=False)
print("✅ Done: Cleaned + Combined data saved as 'gaming_revenue_combined_cleaned.xlsx'")
