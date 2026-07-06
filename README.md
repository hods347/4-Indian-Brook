# 4-Indian-Brook
Hub for tracking purchase, improvements, and miscellaneous projects at the Hodlin house.

## House tracker workbook

- `4-Indian-Brook-House-Tracker.xlsx` — master tracker: cost basis (purchase price + closing
  costs + capital improvements), a project tab for the second-floor flooring (options
  comparison, budget vs. actual, cost log), and a reusable project template.
  Import into Google Sheets via **File > Import > Upload > Replace spreadsheet** to share.
- `build_tracker.py` — regenerates the workbook from scratch (`python3 build_tracker.py`,
  requires `openpyxl`). Edit the config blocks in the script to change quote/cart numbers.
