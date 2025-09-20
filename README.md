###### Developed and tested on Ubuntu
# Excel Automation Tool 📊
Automatically apply discounts to Excel price data and generate charts.

## Features

- Apply configurable discounts to price columns
- Generate bar charts automatically
- Handle invalid data gracefully
- Export processed results

## Requirements

- Python 3.7+
- openpyxl

## Installation

**On Ubuntu/Linux/macOS:**
```bash
git clone https://github.com/jfgmesquita/excel-automation.git
cd excel-automation
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

**On Windows:**
```cmd
git clone https://github.com/jfgmesquita/excel-automation.git
cd excel-automation
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

## Usage

### Basic Usage

**On Ubuntu/Linux/macOS:**
```bash
python3 app.py
```

**On Windows:**
```cmd
python app.py
```

Processes `transactions.xlsx` with 10% discount.

### Advanced Usage
```python
from app import process_workbook

# Custom discount (20% off)
process_workbook("data.xlsx", discount_rate=0.2)

# Custom output name
process_workbook("data.xlsx", output_filename="results.xlsx")
```

## Excel Format

| Column A | Column B | Column C | Column D |
|----------|----------|----------|----------|
| Item Name | Description | Price | (Auto-filled) |
| Laptop | Gaming | 1000.00 | |
| Mouse | Wireless | 50.00 | |

**Requirements:**
- Column C: numeric prices
- Data starts row 2
- Headers in row 1

## Output

Creates `[filename]_corrected.xlsx` with:
- Discounted prices in column D
- Embedded bar chart
- Console progress info

## Troubleshooting

**File not found:** Check file exists and name is correct  
**Invalid data:** Ensure column C has numbers only  
**No chart:** Make sure column A has item names  

## Project Structure

```
excel-automation/
├── app.py
├── README.md
├── requirements.txt
├── transactions.xlsx
└── venv/
```
##
##