# Hankamer Info Session Report Generator

Automatically generates the Hankamer School of Business info-session
PowerPoint from a single Excel data file. Takes ~5 seconds to run.

---

## One-Time Setup (do this once, ever)

### 1. Install Python
Go to https://python.org/downloads and install Python 3.10 or later.
During install, check **"Add Python to PATH"**.

Verify it worked — open Terminal (Mac) or Command Prompt (Windows) and type:
```
python --version
```
You should see something like `Python 3.11.4`.

### 2. Install VS Code (optional but recommended)
Download from https://code.visualstudio.com

### 3. Download this project folder
Put the entire `hankamer_report/` folder somewhere easy to find,
like your Desktop or Documents.

### 4. Install dependencies
Open Terminal / Command Prompt, navigate to the folder, and run:
```
cd path/to/hankamer_report
pip install -r requirements.txt
```
This installs the Python libraries the script needs. Do this once.

### 5. Create your data template
```
python build_data_template.py
```
This creates `data/data_template.xlsx` — your starting point each year.

---

## Every Year — Generating a New Report

### Step 1 — Duplicate and rename the data file
Copy `data/data_template.xlsx` and rename it:
```
data/hankamer_data_AY2627.xlsx   ← change the year each time
```

### Step 2 — Fill in the yellow cells
Open the file. Fill in every **yellow cell** across these sheets:

| Sheet | What to enter |
|-------|--------------|
| `config` | Academic year labels, session count, visitor estimate, states count |
| `attendance` | Monthly headcount for each semester (leave grey cells empty) |
| `grad_year` | Student count by graduation class |
| `top10_states` | Top 10 states, **largest first** (TX at top) |
| `regional` | Student count per region |
| `companion` | Frequency labels for who joins students |
| `bullets` | All text blocks — overview, recommendations, engagement insights |

> **Do not edit cells that are NOT yellow.** Green = auto-calculated formulas.
> Check the cross-check cells at the bottom of `attendance` and `grad_year`
> — they should both say **✔ MATCH** before you run the script.

### Step 3 — Run the script
Open Terminal in the project folder and run:
```
python generate_report.py --input data/hankamer_data_AY2627.xlsx
```

The finished PowerPoint will appear in `output/` named:
```
output/Hankamer_Report_AY2026-2027.pptx
```

Open it in PowerPoint — it's ready to present or share.

---

## Troubleshooting

**"python: command not found"**
→ Python isn't on your PATH. Reinstall from python.org and check "Add to PATH".

**"No module named pptx"**
→ Run `pip install -r requirements.txt` again.

**"Sheet 'attendance' not found"**
→ Make sure your Excel file has all the required sheet names exactly as listed above.

**"✗ MISMATCH" in the data file**
→ Your attendance totals and grad_year totals don't add up. Fix the numbers before running.

**Chart doesn't update in PowerPoint**
→ Right-click the chart → Edit Data → close the window. Or use File → Info → Edit Links.

**Something else broke**
→ Open VS Code, install the Claude extension, open this folder, and describe the problem.
   Claude Code can read all the files and fix it.

---

## Project Structure

```
hankamer_report/
├── generate_report.py      ← main script — run this each year
├── build_data_template.py  ← run once to create the data template
├── template.pptx           ← master PowerPoint template — do not edit
├── requirements.txt        ← Python dependencies
├── README.md               ← this file
├── data/
│   ├── data_template.xlsx          ← blank template (keep as backup)
│   └── hankamer_data_AY2627.xlsx   ← your filled-in data (rename each year)
└── output/
    └── Hankamer_Report_AY2026-2027.pptx   ← generated output
```

---

## What Updates Automatically

| Slide | What changes |
|-------|-------------|
| 1 — Cover | Academic year, all 5 stat boxes, overview bullets, recommendations |
| 2 — Attendance | Bar chart (all 9 months), pie chart (grad year breakdown), callout box, footnote |
| 3 — Geographic | States represented count, regional table, callout boxes (65%/35%/1), bullets, footnote |
| 4 — Family | Top 10 states bar chart, companion frequency table, engagement bullets, footnote, footer |

**What does NOT update automatically:**
- The hexagon map image on slide 3 (static — replace the image manually if needed)

---

## Using Claude Code for Changes

Install the Claude extension in VS Code (search "Claude" in Extensions).
Open this folder in VS Code, then you can say things like:

- *"Add a 5th slide showing year-over-year comparison"*
- *"The regional table now has 7 rows instead of 6 — update the script"*
- *"Export a PDF version as well as the PPTX"*
- *"The input file uses different column names now — fix the parser"*

Claude Code can see all the files and will edit the script directly.
