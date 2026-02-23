# 🚢 Shipping Schedule Organizer

A web application for organizing and managing shipping schedules with multi-format support and automatic Excel export.

## ✨ Features

- **📊 Multi-format Upload**: Excel, CSV, PDF, PNG, JPG
- **🤖 AI-Powered Parsing**: Claude AI reads PDF and image files
- **✏️ Interactive Editing**: Preview and edit schedules in your browser
- **📁 Smart Excel Export**: Auto-generates sheets by Carrier-POD-Month (e.g., "CNC - KHH - MAR")
- **🔄 Duplicate Detection**: Automatically removes duplicate entries

## 📋 Supported Formats

| Format | Processing | API Required | Cost |
|--------|------------|--------------|------|
| Excel/CSV | Column auto-detection | ❌ No | Free |
| PDF | Claude AI text extraction | ✅ Yes | ~$0.01/file |
| PNG/JPG | Claude Vision image reading | ✅ Yes | ~$0.03/image |

## 🚀 Quick Start

### Deploy to Streamlit Cloud (Recommended)

1. Fork/clone this repository
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click "New app" and select your repository
4. Set main file: `app.py`
5. (Optional) Add API key in Settings → Secrets

### Local Development

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 📊 Excel Output Format

**Automatic sheet creation:**
- `All Schedules` - Complete dataset
- `CNC - KHH - MAR` - CNC to Kaohsiung, March
- `YML - HKG - FEB` - YML to Hong Kong, February

**Columns:**
```
CARRIER | POL | POD | Vessel | Voyage | ETD | ETA | 
T/T Time | CY Cut-off | SI Cut-off
```

## 📝 License

MIT License
