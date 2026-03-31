# RevPar MD — Multi-Hotel Dashboard Platform

## Files
- `revpar_app.py`   — Main Streamlit app (run this)
- `hotel_config.py` — Hotel registry (edit to add hotels)
- `data_loader.py`  — All data parsing logic (no edits needed)

## Setup in 4 steps

### 1. Install dependencies
```bash
pip install streamlit pandas plotly openpyxl pyxlsb
```

### 2. Create your OneDrive folder structure
```
RevPar Hotels/
├── Hampton_Lino_Lakes/
│   ├── RevPar_MD_7_1.xlsb
│   ├── STR.xlsx
│   ├── Budget.xlsx
│   ├── year.xlsx
│   ├── Group_Wash_Report.xlsx
│   ├── Rates.xlsx
│   ├── Booking_Reports_SRP_Activity.xlsx
│   ├── 1.xlsx
│   └── 7.xlsx
├── Next_Hotel/
│   └── (same 9 files)
```

### 3. Edit hotel_config.py
- Set `ONEDRIVE_ROOT` to your OneDrive path
- Add each hotel to the `HOTELS` list (template is in the file)

### 4. Run the app
```bash
streamlit run revpar_app.py
```

## Adding hotels
Copy the template block in hotel_config.py HOTELS list.
Set `active: False` for hotels not yet configured — they show as "Coming Soon" cards.

## Dashboard tabs
Dashboard · Biweekly · Day+SRP · Groups · Rates · STR · Demand Nights · Events · Call Recap
