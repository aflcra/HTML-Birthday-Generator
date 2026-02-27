from flask import Flask, render_template, request, jsonify
from docx import Document
import re
from datetime import datetime

app = Flask(__name__)

def parse_birthday_document(file):
    """Parse the uploaded .docx file and extract birthday data"""
    doc = Document(file)
    birthdays = {}
    current_date = None
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
            
        # Check if this is a date header (bold text with date format)
        if para.runs and para.runs[0].bold:
            # Extract date from bold text like "September 8" or "**September 8**"
            date_match = re.match(r'\**(.*?)\**', text)
            if date_match:
                current_date = date_match.group(1).strip()
                birthdays[current_date] = []
        elif current_date:
            # This is a name under the current date
            birthdays[current_date].append(text)
    
    return birthdays

def generate_html(birthdays):
    """Generate Bootstrap HTML from birthday data"""
    html_parts = []
    dates = list(birthdays.keys())
    
    # Process dates in groups of 4 (col-md-3 means 4 columns per row)
    for i in range(0, len(dates), 4):
        html_parts.append('<div class="row">')
        
        # Get up to 4 dates for this row
        row_dates = dates[i:i+4]
        
        for date in row_dates:
            names = birthdays[date]
            names_html = ' <br/> '.join(names)
            
            html_parts.append(f'''  <div class="col-md-3">
    <h3>{date}<br/></h3>
    <p>{names_html}</p>
  </div>''')
        
        html_parts.append('</div>')
    
    return '\n'.join(html_parts)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not file.filename.endswith('.docx'):
        return jsonify({'error': 'Please upload a .docx file'}), 400
    
    try:
        # Parse the document
        birthdays = parse_birthday_document(file)
        
        # Generate HTML
        html_output = generate_html(birthdays)
        
        # Get date range for title
        dates = list(birthdays.keys())
        title = f"{dates[0]} - {dates[-1]} Birthdays" if dates else "Birthdays"
        
        return jsonify({
            'success': True,
            'html': html_output,
            'title': title
        })
    
    except Exception as e:
        return jsonify({'error': f'Error processing file: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True)