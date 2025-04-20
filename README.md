# Python-tkinter-GUI-Data Analytics

A Python/Tkinter application to create flexible, multi‑Y‑axis plots from Excel or CSV data files.

---

<img width="1470" alt="image" src="https://github.com/user-attachments/assets/4c8fc00b-980b-4b74-9d3b-0651df6d9f10" />

## Features

- **Load Data**  
  Import one or multiple Excel (`.xlsx`, `.xls`) or CSV (`.csv`) files.

- **X‑Axis Configuration**  
  Choose a file, sheet (for Excel), and column for the X‑axis.

- **Y‑Axis Configuration**  
  - **Direct**: Plot a single column from a file.  
  - **Derived**: Compute Y values by combining two columns (addition, subtraction, multiplication, division).  
  - Add or remove an arbitrary number of Y‑axes.  
  - Automatic label defaults based on selected column (editable).

- **Dynamic Plotting**  
  - Color‑coded lines, markers, and axis spines.  
  - Endpoint annotations showing final data values.  
  - Seaborn theming for a clean, modern look.

- **Legend Placement**  
  Inside the chart at the top center with small font, out of the way of data.

- **Save Plot**  
  Export your chart to PNG.

- **Responsive Layout**  
  Adjustable sidebar with minimum width and scrollable controls.

---

## Installation

1. Clone the repository (or download the source files).  
2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

   **`requirements.txt` includes:**
   - `pandas`
   - `matplotlib`
   - `seaborn`
   - `openpyxl` (Excel backend)
   - `xlrd` (Excel backend)

3. Run the application:

   ```bash
   python main.py
   ```

---

## Usage

1. Click **Load Excel/CSV…** and select one or more data files.  
2. Under **X‑Axis Configuration**, choose the file, sheet (for Excel), and column for your X‑axis.  
3. Click **+ Add Y‑Axis** to configure each Y‑axis:  
   - Select **Direct** or **Derived**.  
   - Pick files, sheets, and columns (or define an operation between two columns).  
   - Edit the default label if desired.  
4. Set your **X‑Label** and **Title**.  
5. Click **Plot** to generate the chart.  
6. To save, click **Save Plot** and choose a location.

---

## Requirements

- Python 3.7 or higher  
- macOS, Windows, or Linux with a Tkinter‑capable environment

---

## File List

- `main.py` — The main application script.  
- `requirements.txt` — List of Python dependencies.

---


