# Pulte VIP Traffic Sheet Dashboard

Keep these files in the same folder:

- `Tsheet.py`
- `pulte_vip.py`
- `master_template.xlsm`
- `Pulte_Adobe_Tracking_Codes.xlsm`
- `requirements.txt`

Run locally:

```bash
pip install -r requirements.txt
streamlit run Tsheet.py
```

Current status:

- All account names appear in the dashboard.
- Only **Pulte VIP** is enabled.
- The Pulte workflow clears old template data, pastes the new Prisma export,
  creates ad names, matches creatives, builds Adobe `cmp` codes and returns
  a macro-enabled `.xlsm` workbook.
