# Allsvenskan 2026 · Match-by-Match Performance Dashboard

A static, single-team dashboard tracking AIK's 2026 match-by-match KPIs, benchmarked against their 2025 season averages and the full 16-team 2025 Allsvenskan league.

Live at: **https://playful-ganache-8d504d.netlify.app**

---

## What the dashboard shows

### Offensive & Defensive tabs
Each KPI cell is colour-coded by where the match value would rank across the 2025 league:

| Colour | Rank (out of 16) |
|--------|-----------------|
| 🔵 Blue | 1–4 (best) |
| 🟢 Green | 5–8 |
| 🟡 Yellow | 9–12 |
| 🔴 Red | 13–16 (worst) |

For "lower is better" KPIs (opponent ball progression, own defence broken, etc.) the direction is inverted — a low value lands in blue.

The dashboard shows **5 matches at a time**. Use the slider or Prev/Next buttons to move the window. Each KPI row shows:
- The 5 match values for the selected window
- A season average column
- A reference column (2025 avg)
- A dot-plot and match-by-match progression chart when you click a KPI

### Styles tab
Shows AIK's match-by-match score for three playing styles:
- **Possession & Control**
- **Heavy Metal**
- **Direct & Aerial**

Click any style card to open a radar chart with the 8–9 KPIs that define that style, plus the season average overlay. Values are shown next to each dot.

---

## Repository structure

```
repo/
├── index.html          ← the entire dashboard (single file)
├── preprocess.py       ← generates all JSON from the Excel files
├── push.sh             ← one-command update script
├── requirements.txt    ← Python dependency (openpyxl)
├── .nojekyll           ← tells GitHub Pages not to use Jekyll
├── data/
│   ├── matches.json    ← per-match KPI values
│   ├── columns.json    ← all available columns for the picker
│   ├── kpis.json       ← KPI definitions and metadata
│   ├── thresholds.json ← league benchmarks from 2025
│   └── config.json     ← season config (name, window size, etc.)
├── logos/              ← team PNG logos (filename = team name)
└── source-data/
    ├── Allsvenskan_2025.xlsx  ← reference season (all 16 teams)
    └── Allsvenskan_2026.xlsx  ← current season match data
```

---

## One-time setup

### 1. Clone the repo
```bash
git clone https://github.com/tpsantos08/Match-by-match.git
cd Match-by-match
```

### 2. Install Python dependency
```bash
pip install openpyxl
```

### 3. Connect to Netlify
- Go to [netlify.com](https://netlify.com) → Add new site → Import from Git
- Select this repo → Deploy
- Netlify deploys automatically on every push

---

## Updating after each matchday

1. Add the new match row to `source-data/Allsvenskan_2026.xlsx`
2. Run the update script:

```bash
cd ~/Downloads/repo
bash push.sh
```

That's it. The script runs `preprocess.py`, commits everything, and pushes. Netlify deploys within ~1 minute.

---

## Manual push (alternative)

```bash
cd ~/Downloads/repo
python3 preprocess.py
git add .
git commit -m "add M14 data"
git push
```

---

## Adding a new KPI to the dashboard

KPIs are defined in `preprocess.py` in the `KPIS` list. Each entry looks like:

```python
("offensive", "offensive_kpis", "Offensive", "Offensive KPIs",
 "KPI Name",  col_2026,  col_2025, higher_is_better, "format", default_on),
```

- `col_2026` — column number in `Allsvenskan_2026.xlsx`
- `col_2025` — column number in `Allsvenskan_2025.xlsx`
- `higher_is_better` — `True` or `False`
- `format` — `"int"`, `"float"`, `"pct"`, `"xg"`
- `default_on` — `True` to show by default, `False` to hide until added

After editing, run `bash push.sh` to regenerate and deploy.

---

## Column shifts

If InStat updates the Excel file structure and columns shift position, update the column numbers in `preprocess.py` and the `STYLE_RADAR_METRICS` in `index.html`. Always verify against a known match value (e.g. Malmö row = row 11).
