# CO2 Emission Calculator

A web application that calculates carbon dioxide (CO2) emissions from vehicle travel distance.

## ⚠️ Important: Pay Attention to Units
- Meters <> Kilometers
- Kilograms <> Tons

## 📋 Features

### Type A: Calculate from Actual Fuel Efficiency (calc_method:0)
Calculates CO2 emissions using actual fuel efficiency data.

**Input fields:**
- Actual fuel efficiency (km/L)
- Fuel type (Gasoline 2.32 / Diesel 2.58 / LPG 3.00)
- Travel distance (km)

**Formula:**
```
CO2 emissions (kg-CO2) = Travel distance (km) ÷ Actual fuel efficiency (km/L) × Fuel coefficient ÷ 1000
```

### Type B: Calculate from Type-Specific Fuel Efficiency (calc_method:1)
Calculates CO2 emissions using fuel efficiency data by vehicle type from the Ministry of Land, Infrastructure, Transport and Tourism (Japan).

**Input fields:**
- Mode (WLTC / JC08 / 10・15 / JH25 / JH15) ※Auto-search if not selected
- Type designation (e.g., 3BA-NRE210H)
- Fuel type (Gasoline 2.32 / Diesel 2.58 / LPG 3.00)
- Travel distance (km)

**Formula:**
```
CO2 emissions (kg-CO2) = Travel distance ÷ Fuel efficiency value × Fuel coefficient ÷ 1000 × Mode coefficient
```

**Mode coefficients:**
| Mode   | Coefficient |
| ------ | ----------- |
| WLTC   | 1.0         |
| JC08   | 0.9         |
| 10・15 | 0.8         |
| JH25   | 1.0         |
| JH15   | 0.9         |

**Search order:** WLTC → JC08 → 10・15 → JH25 → JH15

## 🚀 How to Use

### Basic Usage

1. Open `index.html` in a browser (double-click)
2. Enter the required information
3. Click the "Calculate" button
4. CO2 emissions will be displayed

### Type Designation Input Format

**Format:** `3 alphanumeric chars-1 or more alphanumeric chars`

**Valid examples:**
- `3BA-NRE210H`
- `DBA-ZVW50`
- `DAA-NHP10`

**Note:**
- Full-width characters are automatically converted to half-width
- Full-width hyphens (－、ー) are also automatically converted to half-width

## 📁 File Structure

```
Co2排出量計算/
├── index.html                    # Main application
├── convert_csv_to_js.py          # CSV→JS conversion script
├── README.md                     # This file (Japanese)
├── README_en.md                  # This file (English)
└── 型式一覧/
    ├── excel_consolidator.py     # Script to organize MLIT Excel files into mode-specific CSVs
    ├── output_WLTC.csv           # WLTC mode CSV data
    ├── output_WLTC.js            # WLTC mode JS data
    ├── output_JC08.csv           # JC08 mode CSV data
    ├── output_JC08.js            # JC08 mode JS data
    ├── output_10-15.csv          # 10・15 mode CSV data
    ├── output_10-15.js           # 10・15 mode JS data
    ├── output_JH25.csv           # JH25 mode CSV data
    ├── output_JH25.js            # JH25 mode JS data
    ├── output_JH15.csv           # JH15 mode CSV data
    ├── output_JH15.js            # JH15 mode JS data
    └── MLIT data folders by year
```

## 🔄 How to Update to Latest Data

1. Download the relevant files from the Ministry of Land, Infrastructure, Transport and Tourism website.
   https://www.mlit.go.jp/jidosha/jidosha_mn10_000002.html

2. Create a folder under the 型式一覧 (type list) directory with the year/month of the downloaded data, and place the downloaded files inside.

3. Run the following commands in WSL (within the 型式一覧 folder):

```bash
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python excel_consolidator.py
```

Five CSV files will be created in the same directory as excel_consolidator.py.

## 🔄 How to Update CSV Data

When CSV files are updated, they must be converted to JavaScript files.
※CSV files are converted to JS for local execution.

### How to Run

```bash
python convert_csv_to_js.py
```

### Example Output

```
============================================================
CSV→JavaScript Conversion Tool
============================================================

✓ Converted: output_WLTC.csv -> output_WLTC.js
  Records: 1,274
✓ Converted: output_JC08.csv -> output_JC08.js
  Records: 2,539
✓ Converted: output_10-15.csv -> output_10-15.js
  Records: 592
✓ Converted: output_JH25.csv -> output_JH25.js
  Records: 1,172
✓ Converted: output_JH15.csv -> output_JH15.js
  Records: 4,236

============================================================
All conversions completed!
============================================================
```

## 📊 CSV File Format

CSV files must follow this format:

```csv
Folder name,File name,Vehicle name,Nickname,Type designation,Fuel efficiency (km/L)
Reiwa 2nd year March,001337994_WLTC_ガソリン乗用車（普通・小型）.xls,Toyota,Corolla Sport,3BA-NRE210H,16.4
Reiwa 2nd year March,001337994_WLTC_ガソリン乗用車（普通・小型）.xls,Suzuki,Jimny,3BA-JB74W,15.0
...
```

**Required columns:**
- `型式` (Type designation) - Vehicle type designation number
- `燃費値（km/L）` (Fuel efficiency km/L) - Catalog fuel efficiency value

## 🧪 Usage Examples

### Example 1: Calculate from Actual Fuel Efficiency

**Input:**
- Actual fuel efficiency: 15.5 km/L
- Fuel type: Gasoline (2.32)
- Travel distance: 10,000 km

**Calculation:**
```
10000 ÷ 15.5 × 2.32 ÷ 1000 = 1.497 t-CO2
```

### Example 2: Calculate from Type-Specific Fuel Efficiency (Mode Specified)

**Input:**
- Mode: WLTC
- Type designation: 3BA-NRE210H
- Fuel type: Gasoline (2.32)
- Travel distance: 10,000 km

**Calculation:**
```
Fuel efficiency (WLTC): 15.8 km/L
10000 ÷ 15.8 × 2.32 ÷ 1000 × 1.0 = 1.468 t-CO2
```

### Example 3: Calculate from Type-Specific Fuel Efficiency (Auto Search)

**Input:**
- Mode: Auto search (not selected)
- Type designation: 3BA-NRE210H
- Fuel type: Gasoline (2.32)
- Travel distance: 10,000 km

**Behavior:**
- Search for type designation in WLTC mode → If found, use that data
- If not found, search in the next mode (JC08)
- Continue searching in order (10・15 → JH25 → JH15)

## 🔧 Technical Specifications

- **HTML5** - Structure and layout
- **CSS3** - Styling (gradients, animations)
- **JavaScript (ES6+)** - Logic and calculations
- **Python 3** - CSV conversion tool, JS conversion tool

## ⚙️ Requirements

### Runtime Environment
- Modern browser (Chrome, Firefox, Edge, Safari)
- JavaScript enabled

### Development Environment (CSV conversion only)
- Python 3.x

## 📞 Support

If you encounter issues, please check:

1. Is JavaScript enabled in your browser?
2. If you updated CSV files, did you run the conversion script?
3. Is the type designation input format correct? (e.g., 3BA-NRE210H)

---

**Created:** January 2026  
**Version:** 1.0

