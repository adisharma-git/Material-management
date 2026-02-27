# Hospital Pharmacy Inventory Manager

A desktop GUI application for hospital pharmacy inventory management.  
Runs on **Windows** and **macOS** — no installation required.

---

## Features

| Tab | Script | What it does |
|-----|--------|--------------|
| Script 1 | Global Stock Lookup | Aggregates stock from ALL hospital locations (CSV → Excel) |
| Script 2 | Main Store Stock Lookup | Aggregates batch-level stock from Main Medical Store |
| Script 4 | Inventory Calculation | Full reorder sheet: master data + stock + pending POs |

---

## Download (pre-built)

Go to the [Releases](../../releases) page and download:

- **Windows:** `PharmacyInventory.exe` – double-click, no install needed
- **macOS:** `PharmacyInventory-macOS.dmg` – drag to Applications

> **macOS note:** On first launch, right-click → Open to bypass Gatekeeper.

---

## Input Files Required

### Script 1
| File | Description |
|------|-------------|
| `Material_Global_Stock_Report.CSV` | Raw global stock report (all locations) |

### Script 2
| File | Description |
|------|-------------|
| `Stock_Report_Material_without_item_category.xlsx` | Batch-level stock from Main Medical Store (`RAW DATA` sheet) |

### Script 4 (run Scripts 1 & 2 first)
| File | Description |
|------|-------------|
| `MASTER_DATA_INPUT.xlsx` | Master item catalog with min/max levels |
| `Material_Global_Stock_Lookup.xlsx` | Output of Script 1 |
| `Material_Main_Store_Stock_Lookup.xlsx` | Output of Script 2 |
| `Expected_Items_Material.xlsx` | Pending purchase orders |

---

## Run Order

```
Script 1  →  generates  Material_Global_Stock_Lookup.xlsx
Script 2  →  generates  Material_Main_Store_Stock_Lookup.xlsx
Script 4  →  generates  INVENTORY_CALCULATION.xlsx   (needs outputs of 1 & 2)
```

---

## Build from Source

### Prerequisites
- Python 3.11+
- pip

### Setup

```bash
git clone https://github.com/YOUR_ORG/pharmacy-inventory.git
cd pharmacy-inventory
pip install -r requirements.txt
python main.py          # Run the GUI directly
```

### Build locally

```bash
pip install pyinstaller

# Windows
pyinstaller pharmacy_inventory.spec --clean --noconfirm
# → dist/PharmacyInventory.exe

# macOS
pyinstaller pharmacy_inventory.spec --clean --noconfirm
# → dist/PharmacyInventory.app
```

---

## GitHub Actions (CI/CD)

Builds are triggered automatically on every push to `main`.  
Tagged releases (`v1.0.0`, etc.) create a GitHub Release with downloadable files.

```
.github/workflows/
  build-windows.yml   →  produces PharmacyInventory.exe
  build-macos.yml     →  produces PharmacyInventory-macOS.dmg
```

To create a release:

```bash
git tag v1.0.0
git push origin v1.0.0
```

---

## Project Structure

```
pharmacy-inventory/
├── main.py                          # GUI entry point
├── src/
│   ├── __init__.py
│   ├── script1_global_stock.py      # Script 1 logic
│   ├── script2_main_store_stock.py  # Script 2 logic
│   └── script4_inventory_calc.py   # Script 4 logic
├── pharmacy_inventory.spec          # PyInstaller spec
├── requirements.txt
├── .github/
│   └── workflows/
│       ├── build-windows.yml
│       └── build-macos.yml
└── README.md
```
