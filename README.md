# Fly91 Invoice Automation

A Flask-based web application to automate invoice PDF generation from Excel data.

## Features
- **Excel Upload**: Upload Master Excel files to generate batch invoices.
- **Interactive Preview**: Position signatures and seals on a real PDF preview.
- **Batch Processing**: Multi-threaded PDF generation for high-volume data.
- **ZIP Download**: Automatic bundling of all invoices into a structured ZIP file.
- **Vercel Ready**: Optimized for serverless deployment with `/tmp` storage.

## Local Setup
1. Clone the repository.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the app:
   ```bash
   python app.py
   ```
4. Open `http://127.0.0.1:5000`.

## Deployment
This project is configured for one-click deployment to Vercel.
- Uses `vercel.json` for Python runtime configuration.
- Filesystem operations use `/tmp` for serverless compatibility.
