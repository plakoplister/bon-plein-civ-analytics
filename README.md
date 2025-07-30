# Cacao Export Analytics - Côte d'Ivoire

Analytics webapp for Ivory Coast cocoa exports data (Oct 2024 - July 2025).

## 🚀 Quick Deploy to Netlify

1. **Drag & Drop**: Upload the `WEBAPP_PUBLICATION/` folder to Netlify
2. **Git Deploy**: Connect this repo to Netlify for automatic deployments

## 📊 Features

- **4 Tabs**: ZES, Transformation, Destinations, Maps
- **Bilingual**: French/English toggle
- **Interactive**: Charts, filters, responsive design
- **Professional**: Arial font, grey theme, no emojis

## 📁 Structure

```
WEBAPP_PUBLICATION/          # Ready-to-deploy webapp
├── index.html              # Main webapp file
├── dynamic_data_enriched.json  # Export data (12.6k records)
├── broyage_data.json       # Processing capacity data
└── logo.PNG               # Bon Plein logo

generate_detailed_cocoa_report.py  # Word report generator
fix_scomcao.py             # Data correction script
```

## 🔄 Recent Updates

- ✅ SCOMCAO merged with S3C (293 transactions)
- ✅ Numbers display as whole tonnes (no decimals)
- ✅ Gap explanation added to reports

## 🛠️ Development

```bash
# Serve locally
cd WEBAPP_PUBLICATION
python3 -m http.server 8000

# Generate report
python3 generate_detailed_cocoa_report.py

# Fix data (if needed)
python3 fix_scomcao.py
```

## 📝 Data Sources

- Monthly export declarations (ABJ + SPY ports)
- Processing capacity survey 2024
- ZES development data
- Country mapping (ISO codes)

---
🇨🇮 **Côte d'Ivoire - Premier producteur mondial de cacao**