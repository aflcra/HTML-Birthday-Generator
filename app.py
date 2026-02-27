from io import BytesIO
from flask import Flask, render_template, request, jsonify
from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import Table
from docx.text.paragraph import Paragraph
import re

app = Flask(__name__)

# Full and abbreviated month names for date detection
MONTH_NAMES = (
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
)
MONTH_ABBREV = ('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                'Jul', 'Aug', 'Sep', 'Sept', 'Oct', 'Nov', 'Dec')
MONTH_PATTERN = '|'.join(MONTH_NAMES) + '|' + '|'.join(MONTH_ABBREV)
# Strict: whole line is "Month Day" or "Mar 2"
DATE_PATTERN = re.compile(
    r'^\**\s*(' + MONTH_PATTERN + r')\s+\d{1,2}(?:\s*[,.]?\s*\d{4})?\s*\**$',
    re.IGNORECASE
)
# Lenient: line starts with "Month Day" (handles trailing/invisible chars)
DATE_PATTERN_START = re.compile(
    r'^\s*\**\s*(' + MONTH_PATTERN + r')\s+\d{1,2}(?:\s*[,.]?\s*\d{4})?',
    re.IGNORECASE
)

# Strip invisible/control chars that Word sometimes inserts
INVISIBLE = re.compile(r'[\u200b\u200c\u200d\ufeff\u00a0]+')

# Set of month names (full + abbrev) for simple "Month Day" check
MONTH_SET = {m.lower() for m in MONTH_NAMES} | {m.lower() for m in MONTH_ABBREV}


def _is_date_line(text):
    """True if line looks like 'March 2' or 'Mar 2' (bulletproof fallback)."""
    parts = text.split()
    if len(parts) < 2:
        return False
    return parts[0].lower() in MONTH_SET and parts[1].isdigit() and len(parts[1]) <= 2


def _normalize(text):
    """Normalize paragraph text for matching."""
    if not text:
        return ''
    t = INVISIBLE.sub(' ', text)
    t = ' '.join(t.split())
    return t.strip()


def _iter_blocks(parent):
    """Yield paragraphs and tables in document order (body and inside tables)."""
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    else:
        parent_elm = parent._tc  # table cell
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def _iter_all_paragraphs(doc):
    """Yield every paragraph in document order, including inside tables."""
    for block in _iter_blocks(doc):
        if isinstance(block, Paragraph):
            yield block
        else:
            for row in block.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        yield para


def parse_birthday_document(file, collect_debug=False):
    """Parse the uploaded .docx file and extract birthday data.
    Dates can be bold or plain; names listed below each date.
    Reads from both body paragraphs and table cells.
    If collect_debug=True, returns (birthdays, debug_lines) for empty-result debugging.
    """
    doc = Document(file)
    birthdays = {}
    current_date = None
    debug_lines = [] if collect_debug else None

    for para in _iter_all_paragraphs(doc):
        raw = para.text
        text = _normalize(raw)
        if not text:
            continue
        if collect_debug:
            debug_lines.append(text[:80])  # first 80 chars per line

        # Bold: can be True or None (inherited from style)
        first_run_bold = (
            para.runs
            and para.runs[0].bold is not False
        )
        # Treat as date if: first run is bold, regex matches, or simple "Month Day" / "Mar 2"
        is_date_line = (
            (first_run_bold and re.match(r'\**(.*?)\**', text))
            or DATE_PATTERN.match(text)
            or DATE_PATTERN_START.match(text)
            or _is_date_line(text)
        )

        if is_date_line:
            if first_run_bold and re.match(r'\**(.*?)\**', text):
                date_match = re.match(r'\**(.*?)\**', text)
                current_date = _normalize(date_match.group(1) if date_match else text)
            elif _is_date_line(text):
                # Use first two tokens as "Month Day" (bulletproof)
                parts = text.split()
                current_date = f"{parts[0]} {parts[1]}"
            else:
                start_match = DATE_PATTERN_START.match(text)
                if start_match:
                    current_date = re.sub(r'^\s*\*+\s*|\s*\*+$', '', start_match.group(0)).strip()
                else:
                    current_date = text
            if current_date:
                birthdays[current_date] = []
        elif current_date:
            birthdays[current_date].append(text)

    if collect_debug:
        return birthdays, debug_lines
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
        # Read into a seekable stream so python-docx gets the full file (Flask uploads can be one-read)
        file_stream = BytesIO(file.read())
        result = parse_birthday_document(file_stream, collect_debug=True)
        birthdays, debug_lines = result
        
        # Generate HTML
        html_output = generate_html(birthdays)
        
        # Get date range for title (ignore empty keys)
        dates = [d for d in birthdays.keys() if d]
        title = f"{dates[0]} - {dates[-1]} Birthdays" if dates else "Birthdays"
        
        resp = {
            'success': True,
            'html': html_output,
            'title': title
        }
        if not dates and debug_lines:
            resp['debug_preview'] = debug_lines[:80]
        return jsonify(resp)
    
    except Exception as e:
        return jsonify({'error': f'Error processing file: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True)
