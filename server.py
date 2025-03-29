import os
import zipfile
import qrcode
import openpyxl
from flask import Flask, render_template, request, jsonify, send_file, redirect
import pandas as pd

app = Flask(__name__)

# Ensure directories exist
EXCEL_SHEETS_DIR = 'Excel Sheets'
QR_CODES_DIR = 'QR Codes'
os.makedirs(EXCEL_SHEETS_DIR, exist_ok=True)
os.makedirs(QR_CODES_DIR, exist_ok=True)

# Configuration for ScanTrack app
SCAN_TRACK_URL = "http://localhost:3000"  # Change this to the actual URL where ScanTrack runs

def generate_qr_codes(event_name):
    """
    Generate QR codes for all students in an event's Excel sheet.
    """
    excel_path = os.path.join(EXCEL_SHEETS_DIR, f'{event_name}_students.xlsx')
    qr_event_dir = os.path.join(QR_CODES_DIR, event_name)
    os.makedirs(qr_event_dir, exist_ok=True)

    try:
        # Use pandas to read the Excel file instead of openpyxl
        # This ensures we're working with the cleaned data
        df = pd.read_excel(excel_path)
        
        # Track successful and failed QR code generations
        success_count = 0
        failed_count = 0

        # Process each row in the DataFrame
        for idx, row in df.iterrows():
            name = row.get('Name')
            roll_number = row.get('Roll Number')
            
            if pd.notna(name) and pd.notna(roll_number):  # Check for non-null values
                # QR code data
                qr_data = f"Name: {name}\nRoll Number: {roll_number}\nEvent: {event_name}"
                
                # Create QR code
                qr = qrcode.QRCode(version=1, box_size=10, border=5)
                qr.add_data(qr_data)
                qr.make(fit=True)
                
                # Create image
                qr_img = qr.make_image(fill_color="black", back_color="white")
                
                # Save QR code with sanitized filename
                filename = f"{str(name).replace(' ', '_')}_{str(roll_number).replace(' ', '_')}.png"
                filepath = os.path.join(qr_event_dir, filename)
                qr_img.save(filepath)
                
                success_count += 1
            else:
                failed_count += 1

        return {
            'status': 'success', 
            'message': f'Generated {success_count} QR codes. {failed_count} rows skipped.'
        }

    except FileNotFoundError:
        return {
            'status': 'error', 
            'message': f'Excel file for {event_name} not found.'
        }
    except Exception as e:
        return {
            'status': 'error', 
            'message': f'Error generating QR codes: {str(e)}'
        }

def create_zip_for_event(event_name):
    """
    Create a zip file of all QR codes for a specific event.
    """
    event_qr_dir = os.path.join(QR_CODES_DIR, event_name)
    zip_path = os.path.join(QR_CODES_DIR, f'{event_name}_qrcodes.zip')

    if not os.path.exists(event_qr_dir):
        return None

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(event_qr_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, event_qr_dir)
                zipf.write(file_path, arcname=arcname)

    return zip_path

def get_available_events():
    """
    Retrieve list of available events from Excel Sheets directory.
    """
    events = []
    for filename in os.listdir(EXCEL_SHEETS_DIR):
        if filename.endswith('_students.xlsx'):
            event_name = filename.replace('_students.xlsx', '')
            events.append(event_name)
    return sorted(events)

def add_event(event_name, file):
    """
    Add a new event with its student list, handling duplicates intelligently
    """
    try:
        # Validate event name
        if not event_name or not event_name.replace(' ', '').isalnum():
            return {
                'status': 'error', 
                'message': 'Invalid event name. Use alphanumeric characters.'
            }
        
        # Check if event already exists
        existing_file = os.path.join(EXCEL_SHEETS_DIR, f'{event_name}_students.xlsx')
        if os.path.exists(existing_file):
            return {
                'status': 'error', 
                'message': 'Event already exists. Choose a different name.'
            }
        
        # Save the uploaded file temporarily
        temp_filepath = os.path.join(EXCEL_SHEETS_DIR, 'temp.xlsx')
        file.save(temp_filepath)
        
        # Validate Excel file
        try:
            # Read the original dataframe
            df = pd.read_excel(temp_filepath)
            
            # Check for minimum columns
            if len(df.columns) < 2:
                os.remove(temp_filepath)
                return {
                    'status': 'error', 
                    'message': 'Excel file must have at least Name and Roll Number columns.'
                }
            
            # Rename columns to ensure consistency
            df.columns = ['Name', 'Roll Number'] + list(df.columns[2:])
            
            # Find duplicate roll numbers
            duplicate_mask = df.duplicated(subset='Roll Number', keep='first')
            
            # Save the final filepath
            final_filepath = os.path.join(EXCEL_SHEETS_DIR, f'{event_name}_students.xlsx')
            
            # If duplicates exist
            if duplicate_mask.any():
                # Prepare duplicate details
                duplicate_details = df[duplicate_mask][['Name', 'Roll Number']].to_dict('records')
                
                # Keep only the first occurrence of each roll number
                df_cleaned = df[~duplicate_mask].reset_index(drop=True)
                
                # Save the cleaned dataframe to the final path
                df_cleaned.to_excel(final_filepath, index=False)
                
                # Remove temp file
                os.remove(temp_filepath)
                
                return {
                    'status': 'warning', 
                    'message': f'Event {event_name} added with duplicates removed.',
                    'events': get_available_events(),
                    'duplicates': duplicate_details,
                    'total_rows': len(df),
                    'kept_rows': len(df_cleaned)
                }
            
            # If no duplicates, save as is to the final path
            df.to_excel(final_filepath, index=False)
            
            # Remove temp file
            os.remove(temp_filepath)
            
            return {
                'status': 'success', 
                'message': f'Event {event_name} added successfully.',
                'events': get_available_events(),
                'total_rows': len(df)
            }
        
        except Exception as e:
            # Remove the temp file if there's an error
            if os.path.exists(temp_filepath):
                os.remove(temp_filepath)
            return {
                'status': 'error', 
                'message': f'Error processing Excel file: {str(e)}'
            }
    
    except Exception as e:
        return {
            'status': 'error', 
            'message': f'Error adding event: {str(e)}'
        }

def parse_qr_data(qr_data):
    """
    Parse QR code data into a dictionary
    """
    data_dict = {}
    for line in qr_data.split('\n'):
        if ':' in line:
            key, value = line.split(':', 1)
            data_dict[key.strip()] = value.strip()
    return data_dict

@app.route('/')
def index():
    events = get_available_events()
    return render_template('index.html', events=events)

@app.route('/add_event', methods=['POST'])
def add_new_event():
    event_name = request.form.get('event_name')
    file = request.files.get('student_list')
    
    if not file:
        return jsonify({
            'status': 'error', 
            'message': 'No file uploaded.'
        })
    
    result = add_event(event_name, file)
    return jsonify(result)

@app.route('/generate_qr', methods=['POST'])
def generate_qr():
    event_name = request.form.get('event')
    result = generate_qr_codes(event_name)
    return jsonify(result)

@app.route('/download_qr/<event_name>')
def download_qr(event_name):
    zip_path = create_zip_for_event(event_name)
    if zip_path:
        return send_file(zip_path, as_attachment=True)
    return jsonify({'status': 'error', 'message': 'No QR codes found for this event.'})

@app.route('/scan_qr')
def scan_qr():
    # Redirect to the external ScanTrack application
    return redirect(SCAN_TRACK_URL)

@app.route('/parse_qr', methods=['POST'])
def parse_qr():
    qr_data = request.form.get('qr_data', '')
    parsed_data = parse_qr_data(qr_data)
    return jsonify(parsed_data)

if __name__ == '__main__':
    app.run(debug=True)