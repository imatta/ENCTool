# Elector Name Comparison Web Application

A web-based interface for the Elector Name Comparison Tool that compares elector names between two Excel sheets (2025_LIST and 2002_LIST) using fuzzy matching.

## Features

- 🌐 Web-based interface - no command line needed
- 📤 File upload with drag-and-drop support
- ⚙️ Configurable similarity threshold (0-100%)
- 📊 Real-time statistics and results display
- 📥 Download results as Excel file
- 🎨 Modern, responsive UI

## Installation

1. Install the required dependencies:
```bash
pip install -r requirements_elector_comparison.txt
```

## Running the Application

1. Start the Flask web server:
```bash
python app.py
```

2. Open your web browser and navigate to:
```
http://localhost:5000
```

3. Upload an Excel file containing:
   - Sheet named `2025_LIST`
   - Sheet named `2002_LIST`
   - Each sheet must have columns: `Elector's Name` and `Elector's Name(Vernacular)`

4. Set the similarity threshold (default: 85%)

5. Click "Process Comparison" and wait for results

6. View results in the browser or download the complete Excel file

## File Requirements

- **Format**: .xlsx or .xls
- **Maximum size**: 50MB
- **Required sheets**: 
  - `2025_LIST`
  - `2002_LIST`
- **Required columns in each sheet**:
  - `Elector's Name`
  - `Elector's Name(Vernacular)`

## Configuration

The application can be configured by modifying `app.py`:

- **Port**: Change `port=5000` in the `app.run()` call
- **Host**: Change `host='0.0.0.0'` to `host='127.0.0.1'` for local-only access
- **Max file size**: Modify `MAX_CONTENT_LENGTH` (default: 50MB)
- **Debug mode**: Set `debug=True` for development (default: True)

## Production Deployment

For production deployment, consider:

1. Using a production WSGI server (e.g., Gunicorn, uWSGI)
2. Setting `debug=False`
3. Using environment variables for configuration
4. Implementing proper authentication/authorization
5. Setting up HTTPS/SSL
6. Using a reverse proxy (e.g., Nginx)

Example with Gunicorn:
```bash
pip install gunicorn
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

## Troubleshooting

- **File upload fails**: Check file size (max 50MB) and format (.xlsx or .xls)
- **Sheets not found**: Ensure sheet names are exactly `2025_LIST` and `2002_LIST`
- **Columns not found**: Verify column names match exactly (case-sensitive)
- **Processing slow**: Large files may take several minutes to process

## License

Same as the main project.

