from fastapi import FastAPI, File, UploadFile, Form, Request, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import pandas as pd
import os
import zipfile
from pathlib import Path
import tempfile
import shutil
from werkzeug.utils import secure_filename
import uvicorn
from typing import Optional

app = FastAPI(title="Division Data Consolidator")

# Configuration
MAX_FILE_SIZE = 16 * 1024 * 1024  # 16MB max file size
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# Create necessary directories
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs('templates', exist_ok=True)

# Setup templates
templates = Jinja2Templates(directory="templates")

def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def process_excel_file(filepath: str, division_col: Optional[str] = None):
    """Process the Excel file and return division files info"""
    try:
        df = pd.read_excel(filepath)
        
        # Auto-detect division column if not provided
        if division_col is None:
            for col in df.columns:
                if col.lower().strip() == 'division':
                    division_col = col
                    break
        
        if division_col is None or division_col not in df.columns:
            return None, f"Division column not found. Available columns: {list(df.columns)}"
        
        # Clean and normalize division names
        df[division_col] = df[division_col].astype(str).str.strip().str.title()
        
        # Get unique divisions
        unique_divisions = df[division_col].unique()
        unique_divisions = [div for div in unique_divisions if not pd.isna(div) and div.lower() != 'nan']
        
        # Create temporary directory for output files
        temp_dir = tempfile.mkdtemp()
        created_files = []
        
        for division in unique_divisions:
            # Filter data for current division
            division_data = df[df[division_col] == division]
            
            # Create filename
            safe_division_name = "".join(c for c in division if c.isalnum() or c in (' ', '-', '_')).rstrip()
            safe_division_name = safe_division_name.replace(' ', '_')
            filename = f"{safe_division_name}_Div.xlsx"
            filepath = os.path.join(temp_dir, filename)
            
            # Save to Excel file
            division_data.to_excel(filepath, index=False)
            created_files.append({
                'filename': filename,
                'filepath': filepath,
                'row_count': len(division_data),
                'division': division
            })
        
        return created_files, temp_dir
        
    except Exception as e:
        return None, f"Error processing file: {str(e)}"

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/upload")
async def upload_file(
    request: Request,
    file: UploadFile = File(...),
    division_col: Optional[str] = Form("")
):
    if not file.filename:
        return templates.TemplateResponse("index.html", {
            "request": request, 
            "error": "No file selected"
        })
    
    if not allowed_file(file.filename):
        return templates.TemplateResponse("index.html", {
            "request": request, 
            "error": "Invalid file type. Please upload .xlsx or .xls files only."
        })
    
    # Check file size
    contents = await file.read()
    if len(contents) > MAX_FILE_SIZE:
        return templates.TemplateResponse("index.html", {
            "request": request, 
            "error": "File size exceeds 16MB limit."
        })
    
    # Save uploaded file
    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    
    with open(filepath, "wb") as f:
        f.write(contents)
    
    # Process the file
    division_col_value = division_col.strip() if division_col else None
    created_files, result = process_excel_file(filepath, division_col_value)
    
    # Clean up uploaded file
    os.remove(filepath)
    
    if created_files is None:
        return templates.TemplateResponse("index.html", {
            "request": request, 
            "error": result
        })
    
    # Create ZIP file with all division files
    zip_filename = f"divisions_{filename.rsplit('.', 1)[0]}.zip"
    zip_filepath = os.path.join(OUTPUT_FOLDER, zip_filename)
    
    with zipfile.ZipFile(zip_filepath, 'w') as zipf:
        for file_info in created_files:
            zipf.write(file_info['filepath'], file_info['filename'])
    
    # Clean up temporary files
    shutil.rmtree(result)
    
    return templates.TemplateResponse("success.html", {
        "request": request,
        "files": created_files,
        "zip_filename": zip_filename,
        "total_divisions": len(created_files)
    })

@app.get("/download/{filename}")
async def download_file(filename: str):
    filepath = os.path.join(OUTPUT_FOLDER, filename)
    if os.path.exists(filepath):
        return FileResponse(
            path=filepath,
            filename=filename,
            media_type='application/octet-stream'
        )
    raise HTTPException(status_code=404, detail="File not found")

def create_templates():
    """Create HTML template files"""
    
    # Create base.html
    with open('templates/base.html', 'w', encoding='utf-8') as f:
        f.write('''<!DOCTYPE html>
<html>
<head>
    <title>Division Consolidator</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        .drag-drop-area {
            border: 2px dashed #ccc;
            border-radius: 10px;
            padding: 40px;
            text-align: center;
            margin: 20px 0;
            transition: border-color 0.3s;
        }
        .drag-drop-area:hover {
            border-color: #007bff;
        }
        .file-info {
            background: #f8f9fa;
            padding: 15px;
            border-radius: 5px;
            margin: 10px 0;
        }
    </style>
</head>
<body>
    <nav class="navbar navbar-dark bg-dark">
        <div class="container">
            <span class="navbar-brand mb-0 h1">📊 Division Data Consolidator</span>
        </div>
    </nav>
    
    <div class="container mt-4">
        {% if error %}
            <div class="alert alert-warning alert-dismissible fade show" role="alert">
                {{ error }}
                <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
            </div>
        {% endif %}
        
        {% block content %}{% endblock %}
    </div>
    
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>''')
    
    # Create index.html
    with open('templates/index.html', 'w', encoding='utf-8') as f:
        f.write('''{% extends "base.html" %}

{% block content %}
<div class="row justify-content-center">
    <div class="col-md-8">
        <div class="card">
            <div class="card-header">
                <h3 class="mb-0">Upload Excel File for Division Processing</h3>
            </div>
            <div class="card-body">
                <form method="POST" action="/upload" enctype="multipart/form-data">
                    <div class="drag-drop-area">
                        <i class="fas fa-cloud-upload-alt fa-3x text-muted mb-3"></i>
                        <h5>Select or Drop your Excel file here</h5>
                        <input type="file" class="form-control mt-3" name="file" accept=".xlsx,.xls" required>
                        <small class="text-muted">Supported formats: .xlsx, .xls (Max 16MB)</small>
                    </div>
                    
                    <div class="mb-3">
                        <label class="form-label">Division Column Name (Optional)</label>
                        <input type="text" class="form-control" name="division_col" placeholder="Leave empty for auto-detection">
                        <small class="text-muted">If your division column is not named 'division', specify it here</small>
                    </div>
                    
                    <button type="submit" class="btn btn-primary btn-lg w-100">
                        Process File
                    </button>
                </form>
            </div>
        </div>
        
        <div class="card mt-4">
            <div class="card-body">
                <h5>How it works:</h5>
                <ol>
                    <li>Upload your Excel file containing division data</li>
                    <li>The tool will automatically detect divisions (case-insensitive)</li>
                    <li>Creates separate Excel files for each division</li>
                    <li>Downloads all files as a ZIP package</li>
                </ol>
                
                <div class="alert alert-info">
                    <strong>Examples:</strong> If you have divisions like "UP-1", "UP-2", "Assam", "GUJARAT", 
                    you'll get separate files: UP-1_Div.xlsx, UP-2_Div.xlsx, Assam_Div.xlsx, Gujarat_Div.xlsx
                </div>
            </div>
        </div>
    </div>
</div>
{% endblock %}''')
    
    # Create success.html
    with open('templates/success.html', 'w', encoding='utf-8') as f:
        f.write('''{% extends "base.html" %}

{% block content %}
<div class="row justify-content-center">
    <div class="col-md-8">
        <div class="alert alert-success">
            <h4 class="alert-heading">Processing Complete!</h4>
            <p>Successfully processed your file and created <strong>{{ total_divisions }}</strong> division files.</p>
        </div>
        
        <div class="card">
            <div class="card-header">
                <h5>Generated Division Files</h5>
            </div>
            <div class="card-body">
                {% for file in files %}
                <div class="file-info">
                    <strong>{{ file.filename }}</strong>
                    <span class="badge bg-secondary ms-2">{{ file.row_count }} rows</span>
                    <div class="text-muted small">Division: {{ file.division }}</div>
                </div>
                {% endfor %}
                
                <div class="mt-3">
                    <a href="/download/{{ zip_filename }}" class="btn btn-success btn-lg">
                        Download All Files (ZIP)
                    </a>
                    <a href="/" class="btn btn-outline-primary ms-2">
                        Process Another File
                    </a>
                </div>
            </div>
        </div>
    </div>
</div>
{% endblock %}''')

if __name__ == '__main__':
    # Create template files
    create_templates()
    
    print("FastAPI app setup complete!")
    print("Starting server...")
    print("Access the application at: http://localhost:8000")
    
    # Run the application
    uvicorn.run(
        "main:app",  # Change this to your filename:app if different
        host="0.0.0.0",
        port=8000,
        reload=True
    )
