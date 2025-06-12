# app.py

import os
import uuid
import pandas as pd
import xml.etree.ElementTree as ET
from flask import (Flask, request, render_template, send_from_directory,
                   redirect, url_for, flash)
from werkzeug.utils import secure_filename

# --- Configuration ---
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'twb'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max upload size
app.secret_key = 'supersecretkey' # Needed for flashing messages

# --- Helper Functions ---

def allowed_file(filename):
    """Checks if the file extension is allowed."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def parse_twb(file_path):
    """
    Parses a .twb file and extracts metadata.
    A .twb file is an XML file.
    """
    tree = ET.parse(file_path)
    root = tree.getroot()

    # --- Data Extraction Dictionaries ---
    data = {
        'Datasources': [],
        'Worksheets': [],
        'Dashboards': [],
        'Calculated_Fields': [],
        'Parameters': []
    }

    # 1. Extract Datasources and their Connections
    for ds in root.findall('.//datasource'):
        ds_name = ds.get('caption', ds.get('name', 'N/A'))
        connection = ds.find('connection')
        if connection is not None:
            conn_class = connection.get('class', 'N/A')
            db_name = connection.get('dbname', 'N/A')
            server = connection.get('server', 'N/A')
            username = connection.get('username', 'N/A')
        else:
            conn_class, db_name, server, username = 'N/A', 'N/A', 'N/A', 'N/A'
        
        data['Datasources'].append({
            'Datasource Name': ds_name,
            'Connection Class': conn_class,
            'Database Name': db_name,
            'Server': server,
            'Username': username
        })
    
    # 2. Extract Calculated Fields
    for field in root.findall('.//column'):
        if field.find('calculation') is not None:
            field_name = field.get('caption', field.get('name'))
            formula = field.find('calculation').get('formula', 'N/A').strip()
            # Find parent datasource
            parent_ds = field.find('../..')
            ds_name = parent_ds.get('caption', parent_ds.get('name', 'Unknown Datasource'))

            data['Calculated_Fields'].append({
                'Field Name': field_name,
                'Formula': formula,
                'Datasource': ds_name
            })
    
    # 3. Extract Parameters
    for param in root.findall('.//parameters/parameter'):
        param_name = param.get('name')
        param_caption = param.get('caption', param_name)
        data_type = param.get('datatype', 'N/A')
        value = param.get('value', 'N/A')
        
        data['Parameters'].append({
            'Parameter Name': param_caption,
            'Internal Name': param_name,
            'Data Type': data_type,
            'Current Value': value
        })

    # 4. Extract Worksheets
    for ws in root.findall('.//worksheets/worksheet'):
        data['Worksheets'].append({'Worksheet Name': ws.get('name')})
    
    # 5. Extract Dashboards and their contained sheets
    for window in root.findall('.//windows/window[@class="dashboard"]'):
        dashboard_name = window.get('name')
        # Find all zones in the dashboard that refer to a worksheet
        for zone in window.findall('.//zone'):
            sheet_name = zone.get('name')
            if sheet_name and sheet_name in [w['Worksheet Name'] for w in data['Worksheets']]:
                 data['Dashboards'].append({
                    'Dashboard Name': dashboard_name,
                    'Contained Worksheet': sheet_name
                })
    
    return data

def create_excel_report(data, output_path):
    """Creates a multi-sheet Excel report from the extracted data."""
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for sheet_name, sheet_data in data.items():
            if sheet_data:  # Only create a sheet if there is data
                df = pd.DataFrame(sheet_data)
                df.to_excel(writer, sheet_name=sheet_name, index=False)

# --- Flask Routes ---

@app.route('/', methods=['GET'])
def index():
    """Renders the main upload page."""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handles file upload, processing, and serves the Excel file."""
    if 'file' not in request.files:
        flash('No file part in the request.')
        return redirect(request.url)
    
    file = request.files['file']

    if file.filename == '':
        flash('No selected file.')
        return redirect(request.url)

    if file and allowed_file(file.filename):
        # Generate a unique filename to avoid conflicts
        original_filename = secure_filename(file.filename)
        unique_id = str(uuid.uuid4().hex)
        input_filename = f"{unique_id}_{original_filename}"
        output_filename = f"{unique_id}_metadata_report.xlsx"
        
        input_filepath = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
        output_filepath = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

        try:
            # Save uploaded file
            file.save(input_filepath)
            
            # Process the file
            extracted_data = parse_twb(input_filepath)
            
            # Generate Excel report
            create_excel_report(extracted_data, output_filepath)
            
            # Send the file to the user for download
            return send_from_directory(
                app.config['OUTPUT_FOLDER'],
                output_filename,
                as_attachment=True,
                download_name=f"{original_filename.replace('.twb', '')}_metadata.xlsx"
            )

        except Exception as e:
            app.logger.error(f"An error occurred: {e}")
            flash(f"An error occurred while processing the file: {e}")
            return redirect(url_for('index'))
        
        finally:
            # Clean up the temporary files
            if os.path.exists(input_filepath):
                os.remove(input_filepath)
            if os.path.exists(output_filepath):
                # NOTE: A more robust solution might delay this deletion
                # or use a background job, but for this simple app it's okay.
                # The file might not be fully downloaded yet. Flask's send_from_directory
                # streams the file, so this cleanup is generally safe after the response is sent.
                # However, for production apps, consider a scheduled cleanup task.
                pass # Keeping output for now for easier debugging, remove the 'pass' to enable cleanup.
                # os.remove(output_filepath) 

    else:
        flash('Invalid file type. Please upload a .twb file.')
        return redirect(url_for('index'))

if __name__ == '__main__':
    # Create necessary folders if they don't exist
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    # For development, use debug=True. For production, use a proper WSGI server like Gunicorn.
    app.run(debug=True)
