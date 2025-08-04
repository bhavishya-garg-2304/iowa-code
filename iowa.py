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
    # Step 1: Get all candidate words (non-numeric, in casino column area)
    words = page.extract_words()
    column_word_map = defaultdict(list)

    for w in words:
        if w['x0'] >= first_casino_x0 and not any(char.isdigit() for char in w['text']):
            col_idx = get_column_index(w['x0'])
            if 0 <= col_idx < num_casinos:
                column_word_map[col_idx].append(w)

    casino_names = []

    for i in range(num_casinos):
        col_words = column_word_map[i]
        if not col_words:
            casino_names.append(f"Casino {i+1}")
            continue

        # Step 2: Find top-most word in the column — that starts the name block
        top_word = min(col_words, key=lambda w: w["top"])
        name_y0 = top_word["top"]
        name_y1 = name_y0 + 40  # 40pt tall box

        # Step 3: Extract all words within this dynamic box and current column's x-range
        x0 = first_casino_x0 + i * casino_col_width
        x1 = x0 + casino_col_width
        cropped = page.within_bbox((x0, name_y0, x1, name_y1))
        name_words = cropped.extract_words()

        # Step 4: Sort and combine
        name_sorted = sorted(name_words, key=lambda w: (w["top"], w["x0"]))
        name = ' '.join(w["text"] for w in name_sorted).replace(' -', '').strip()
        name = re.sub(r'\s+', ' ', name)

        casino_names.append(name)

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
