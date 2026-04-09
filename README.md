# 📌 SmartPaste AI — Intelligent Copy-Paste for Google Sheets

> Part of the **WITI Google Workspace Suite** by [Sanadidari SARL](https://sanadidari.com)

AI-powered Google Sheets add-on that transforms raw pasted data into structured, clean spreadsheet content. Uses Google Gemini to transmute unstructured input (CSV, JSON, plain text, tables) into the correct format for your sheet.

## Features

- Intelligent data transmutation from any format to structured columns
- Detects headers, data types, and encoding automatically
- Handles messy copy-paste from PDFs, emails, websites
- Credit-based usage system per user

## Stack

Google Apps Script · Google Gemini API · HTML Service · PropertiesService

## Setup

1. Open the Google Sheets add-on editor and paste `Code.gs`
2. Run `setupApiKey()` from the Apps Script editor
3. Enter your Gemini API key when prompted
4. Reload the sheet — the SmartPaste AI menu will appear

## Security

API keys are stored via `PropertiesService.getScriptProperties()` — never hardcoded.

---

*Built by [Samir Chatwiti](https://sanadidari.com) · Sanadidari SARL*
