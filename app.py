from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import boto3
import pandas as pd
import os
from werkzeug.utils import secure_filename
import uuid
from credentials import AWS_ACCESS_KEY, AWS_SECRET_KEY, AWS_REGION  # Import credentials

app = Flask(__name__)
app.secret_key = 'supersecretkey'

# Configure upload and output folders
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'pdf'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Initialize Textract client
textract = boto3.client(
    'textract',
    aws_access_key_id=AWS_ACCESS_KEY,
    aws_secret_access_key=AWS_SECRET_KEY,
    region_name=AWS_REGION
)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_table_data(blocks):
    cells = {}
    for block in blocks:
        if block['BlockType'] == 'CELL':
            row_index = int(block['RowIndex'])
            col_index = int(block['ColumnIndex'])

            cell_text = ""
            if 'Relationships' in block:
                for rel in block['Relationships']:
                    if rel['Type'] == 'CHILD':
                        for child_id in rel['Ids']:
                            child_block = next((b for b in blocks if b['Id'] == child_id), None)
                            if child_block and child_block['BlockType'] == 'WORD':
                                cell_text += child_block.get('Text', '') + " "

            cell_text = cell_text.strip()
            cells[(row_index, col_index)] = cell_text

    return cells

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if file is uploaded
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)

            # Read file
            with open(filepath, 'rb') as f:
                img_bytes = f.read()

            try:
                # Call Textract
                response = textract.analyze_document(
                    Document={'Bytes': img_bytes},
                    FeatureTypes=['TABLES']
                )

                # Extract table data
                table_cells = extract_table_data(response['Blocks'])

                if not table_cells:
                    flash("No tables found in the document.")
                    return redirect(request.url)

                # Reconstruct table
                max_row = max(row for row, _ in table_cells.keys())
                max_col = max(col for _, col in table_cells.keys())

                table_data = []
                for row in range(1, max_row + 1):
                    row_data = []
                    for col in range(1, max_col + 1):
                        row_data.append(table_cells.get((row, col), ''))
                    table_data.append(row_data)

                # Convert to DataFrame
                headers = table_data[0]
                df = pd.DataFrame(table_data[1:], columns=headers)

                # Save to Excel
                output_filename = f"output_{uuid.uuid4().hex}.xlsx"
                output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
                df.to_excel(output_path, index=False)

                # Pass filename for download
                return render_template('index.html', download=output_filename)

            except Exception as e:
                flash(f"Error processing file: {str(e)}")
                return redirect(request.url)

        else:
            flash('Invalid file type. Only images (PNG, JPG, JPEG) and PDF allowed.')
            return redirect(request.url)

    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    if os.path.exists(path):
        return send_file(path, as_attachment=True, download_name='extracted_table.xlsx')
    else:
        flash("File not found.")
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
