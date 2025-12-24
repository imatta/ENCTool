#!/usr/bin/env python3
"""
Flask Web Application for Elector Name Comparison
Provides a web interface for the elector name duplicate finder tool.
"""

import os
import sys
from flask import Flask, render_template, request, send_file, jsonify, flash, redirect, url_for
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge
import tempfile
import shutil
from pathlib import Path

# Import the comparator class from the existing script
from elector_name_comparison import ElectorNameComparator

app = Flask(__name__)
app.config['SECRET_KEY'] = os.urandom(24)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp(prefix='elector_uploads_')
app.config['RESULTS_FOLDER'] = tempfile.mkdtemp(prefix='elector_results_')

# Ensure directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['RESULTS_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}


def allowed_file(filename):
    """Check if file extension is allowed."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """Render the main upload page."""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and process comparison."""
    try:
        # Check if file was uploaded
        if 'file' not in request.files:
            flash('No file uploaded. Please select an Excel file.', 'error')
            return redirect(url_for('index'))
        
        file = request.files['file']
        
        if file.filename == '':
            flash('No file selected. Please choose an Excel file.', 'error')
            return redirect(url_for('index'))
        
        if not allowed_file(file.filename):
            flash('Invalid file type. Please upload an Excel file (.xlsx or .xls).', 'error')
            return redirect(url_for('index'))
        
        # Get similarity threshold
        threshold = request.form.get('threshold', '85', type=int)
        if threshold < 0 or threshold > 100:
            threshold = 85
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Initialize comparator
        try:
            comparator = ElectorNameComparator(filepath, similarity_threshold=threshold)
        except Exception as e:
            flash(f'Error initializing comparator: {str(e)}', 'error')
            return redirect(url_for('index'))
        
        # Load Excel sheets
        if not comparator.load_excel_sheets():
            flash('Failed to load Excel sheets. Please ensure the file contains "2025_LIST" and "2002_LIST" sheets with required columns.', 'error')
            return redirect(url_for('index'))
        
        # Compare names
        duplicates = comparator.compare_names()
        
        # Export results
        timestamp = comparator.export_results()
        output_filename = os.path.basename(timestamp)
        output_path = os.path.join(app.config['RESULTS_FOLDER'], output_filename)
        shutil.move(timestamp, output_path)
        
        # Prepare results data for display
        results_data = {
            'filename': filename,
            'output_filename': output_filename,
            'stats': comparator.stats,
            'duplicates_count': len(duplicates),
            'threshold': threshold,
            'duplicates': duplicates[:100],  # Limit to first 100 for display
            'total_duplicates': len(duplicates)
        }
        
        # Clean up uploaded file
        try:
            os.remove(filepath)
        except:
            pass
        
        return render_template('results.html', **results_data)
        
    except RequestEntityTooLarge:
        flash('File too large. Maximum file size is 50MB.', 'error')
        return redirect(url_for('index'))
    except Exception as e:
        flash(f'An error occurred: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/download/<filename>')
def download_file(filename):
    """Download the results Excel file."""
    try:
        filepath = os.path.join(app.config['RESULTS_FOLDER'], secure_filename(filename))
        if os.path.exists(filepath):
            return send_file(filepath, as_attachment=True, download_name=filename)
        else:
            flash('File not found.', 'error')
            return redirect(url_for('index'))
    except Exception as e:
        flash(f'Error downloading file: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/health')
def health():
    """Health check endpoint."""
    return jsonify({'status': 'healthy'})


@app.errorhandler(413)
def too_large(e):
    """Handle file too large error."""
    flash('File too large. Maximum file size is 50MB.', 'error')
    return redirect(url_for('index'))


@app.errorhandler(500)
def internal_error(e):
    """Handle internal server errors."""
    flash('An internal error occurred. Please try again.', 'error')
    return redirect(url_for('index'))


if __name__ == '__main__':
    # Clean up old files on startup
    for folder in [app.config['UPLOAD_FOLDER'], app.config['RESULTS_FOLDER']]:
        for filename in os.listdir(folder):
            filepath = os.path.join(folder, filename)
            try:
                if os.path.isfile(filepath):
                    os.remove(filepath)
            except:
                pass
    
    print(f"Starting Elector Name Comparison Web Application...")
    print(f"Upload folder: {app.config['UPLOAD_FOLDER']}")
    print(f"Results folder: {app.config['RESULTS_FOLDER']}")
    print(f"Access the application at: http://localhost:5000")
    
    app.run(debug=True, host='0.0.0.0', port=5000)

