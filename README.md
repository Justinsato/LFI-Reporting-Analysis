# Rent Roll Forecasting Function App

This Azure Function forecasts tenant turnover probability using move-in dates and rent pressure factors.

## 🧠 Features
- Accepts Excel rent roll uploads
- Calculates tenure and turnover probability
- Outputs updated Excel file

## 🚀 Deploy via GitHub Actions
This repo uses `.github/workflows/deploy.yml` to auto-deploy to Azure Functions.

### Setup
1. Create a Function App in Azure
2. Download its **Publish Profile**
3. Add it as a GitHub secret: `AZURE_FUNCTIONAPP_PUBLISH_PROFILE`
4. Push to `main`

## 📂 Structure
```
ForecastFunction/
├── __init__.py
├── function.json
requirements.txt
host.json
```

## 📥 Trigger
HTTP POST with file upload (`multipart/form-data`) containing an Excel file.

