# Excel-Data-Transfer

Here's a clean and informative `README.md` file for your GitHub repository:

---

````markdown
# Excel Data Transfer Script using `openpyxl`

This Python script automates the process of copying data from one Excel workbook (`source.xlsx`) to another template workbook (`template.xlsx`) using the `openpyxl` library. The final output is saved as `output.xlsx`.

## 📄 Overview

- Copies data from `source.xlsx` to `template.xlsx`
- Assumes that:
  - Both files have the same structure (i.e., same number and order of columns)
  - Headers are in the first row
  - Data begins in row 2
  - There are 50 rows of data to transfer

## 🧩 Requirements

- Python 3.x
- `openpyxl` library

Install `openpyxl` with:

```bash
pip install openpyxl
````

## 🚀 How to Use

1. Place the following files in the same directory as the script:

   * `source.xlsx` — contains the data you want to transfer
   * `template.xlsx` — the workbook to receive the data

2. Run the script:

```bash
python transfer.py
```

3. A new file named `output.xlsx` will be created with the copied data.

## 🛠 Code

```python
from openpyxl import load_workbook

wb_source = load_workbook("source.xlsx")
ws_source = wb_source.active

wb_target = load_workbook("template.xlsx")
ws_target = wb_target.active

# Copy values (assuming headers match and data starts at row 2)
for i in range(2, 52):  # 50 rows
    for j in range(1, ws_target.max_column + 1):
        ws_target.cell(row=i, column=j).value = ws_source.cell(row=i, column=j).value

wb_target.save("output.xlsx")
```

## 📝 Notes

* You can adjust the row range (`range(2, 52)`) if your dataset has more or fewer rows.
* Ensure that `source.xlsx` and `template.xlsx` use consistent headers and cell layouts.

## 📦 Output

* ✅ `output.xlsx` — contains data from `source.xlsx` placed into the structure of `template.xlsx`.
