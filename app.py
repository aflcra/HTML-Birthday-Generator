from flask import Flask, render_template, request, jsonify
from docx import Document
import re
from datetime import datetime

app = Flask(__name__)

# Month names for date-pattern fallback (dates without bold)
MONTH_NAMES = (
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
)
DATE_PATTERN = re.compile(
    r'^\**\s*(' + '|'.join(MONTH_NAMES) + r')\s+\d{1,2}(?:\s*[,.]?\s*\d{4})?\s*\**$',
    re.IGNORECASE
)


def parse_birthday_document(file):
    """Parse the uploaded .docx file and extract birthday data.
    Dates can be bold or plain; names listed below each date.
    """
    doc = Document(file)
    birthdays = {}
    current_date = None

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # Bold: can be True or None (inherited from style)
        first_run_bold = (
            para.runs
            and para.runs[0].bold is not False
        )
        # Treat as date if: first run is bold, or line looks like "Month Day"
        is_date_line = (
            (first_run_bold and re.match(r'\**(.*?)\**', text))
            or DATE_PATTERN.match(text)
        )

        if is_date_line:
            if first_run_bold:
                date_match = re.match(r'\**(.*?)\**', text)
                current_date = date_match.group(1).strip() if date_match else text
            else:
                current_date = text.strip()
            # Only add non-empty date so we don't create empty rows
            if current_date:
                birthdays[current_date] = []
        elif current_date:
            birthdays[current_date].append(text)

    return birthdays

def generate_html(birthdays):
    """Generate Bootstrap HTML from birthday data"""
    # Ignore empty date keys (avoids empty row with blank h3/p)
    dates = [d for d in birthdays.keys() if d]
    if not dates:
        return '<p class="text-muted">No birthday data found. Use bold dates (e.g. <strong>September 8</strong>) or lines like "September 8" with names listed below each date.</p>'

    html_parts = []

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
        
        # Get date range for title (ignore empty keys)
        dates = [d for d in birthdays.keys() if d]
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
