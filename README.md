<img width="1920" height="1033" alt="{C50DB3AE-A8F6-4578-A97D-2D6A1DECC37F}" src="https://github.com/user-attachments/assets/37f9e149-bd60-4382-bb0a-a8b248c63f46" />


# Vehicle Labels

A desktop application for generating and printing vehicle inventory labels on Zebra printers. Built with Python and tkinter — no browser required.

Prints combo labels (1 large + 4 small VIN labels) or individual labels with Code 128-B barcodes.

## Download

Grab **[VehicleLabels.exe](dist/VehicleLabels.exe)** from the `dist/` folder and run it on any Windows PC (no Python required).

> **Note:** Windows SmartScreen may warn on first launch since the exe is unsigned. Click **More info → Run anyway**.

---

## Getting Started

There are **two ways** to load vehicle data:

### Option A — CDK Inventory Export (your existing workflow)

1. Scan vehicles with your laser scanner
2. Upload the scanned data into CDK
3. Download the inventory `.xls` report from CDK
4. Open the app → click **1. Import report** → select the `.xls` file
5. The app automatically converts the CDK report format

### Option B — Manual Entry Template (no scanner needed)

1. Download the [formatted template](templates/VehicleLabels_Template.xlsx) below
2. Open it in Excel, Google Sheets, or any spreadsheet editor
3. Delete the example row and fill in your vehicle data (see column guide below)
4. Save the file as `.xlsx`
5. Open the app → click **1. Import report** → select your `.xlsx` file

---

## Spreadsheet Column Layout

Both the CDK export and the template use these columns:

| Column | Header      | Example          | Notes                                |
|--------|-------------|------------------|--------------------------------------|
| A      | Location    | LOT A            | Lot / zone identifier                |
| B      | Location    | 1                | Sub-location (section, row, etc.)   |
| C      | VIN         | 1TESTVEHICLE00001| **17-character VIN** — must be exactly 17 alphanumeric characters |
| D      | Stock #     | EX-00-1 A        | Stock / inventory code               |
| E      | Type        | TLX              | Vehicle model type                   |
| F      | Year Make   | 25 ACURA         | Year + manufacturer                  |
| G      | Color       | LUNAR SILVER     | Exterior color                       |
| H      | Co          | 1                | Count                                |

> **Important:** The VIN **must** be in column C (column index 4 in Excel). The app locates rows by finding 17-character VINs in this column.

---

## How It Works

1. **Import** — load vehicle data via the CDK `.xls` route or a formatted `.xlsx`
2. **Filter & Select** — search by VIN/stock/location, filter by location, select vehicles with **All/None** or double-click individual rows
3. **Preview** — click any row to see a live label preview in the right panel
4. **Adjust layout** — open the **Layout Editor** to fine-tune text and barcode positions; changes are visible in real time
5. **Print** — **Download ZPL** saves a `.zpl` file, **Send to Printer** sends directly over TCP

> **Safety:** No labels are selected by default after import. You must explicitly select which vehicles to print before sending to the printer.

---

## Printer Configuration

**Printer IP and port are set in the app** — click the **Printer...** button in the bottom bar and enter the Zebra's IP address (e.g. `192.0.2.1`). The setting saves automatically.

All settings (including DPI and shift values) are stored in:

```
%LOCALAPPDATA%\VehicleLabeler\config.json
```

```json
{
  "printer": { "ip": "192.0.2.1", "port": 9100 },
  "dpi": 300,
  "adj_x": 0.0,
  "adj_y": 0.35
}
```

| Key      | What it does                                      | Default  |
|----------|---------------------------------------------------|----------|
| `ip`     | Zebra printer IP address (set via the **Printer...** button in the app) | `192.0.2.1` |
| `port`   | TCP port (usually 9100)                           | `9100`   |
| `dpi`    | Printer resolution — **300** or 203                | `300`    |
| `adj_x`  | Horizontal shift (inches) — moves all labels left/right | `0.0`   |
| `adj_y`  | Vertical shift (inches) — moves all labels up/down | `0.35`  |

Layout positions (barcode/text placement) are stored separately in `layout.json` in the same folder. Use the **Layout Editor** in the app to adjust them — changes save automatically.

---

## Build from Source

Requires Python 3.12+ with the standard library (no pip dependencies).

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name VehicleLabels vehicle_labels_app.py
```

The finished exe is in `dist/VehicleLabels.exe`.

---

## Project Structure

```
VehicleLabels/
├── vehicle_labels_app.py          # Full application (source)
├── VehicleLabels.spec             # PyInstaller build config
├── dist/
│   └── VehicleLabels.exe          # Standalone Windows exe
└── templates/
    └── VehicleLabels_Template.xlsx # Blank template for manual entry
```

---

## License

MIT
