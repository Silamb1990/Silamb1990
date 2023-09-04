from flask import Flask, render_template, request, redirect, url_for, send_file
import tabula
import pandas as pd
from openpyxl import Workbook
from io import BytesIO

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def upload_pdf():
    if request.method == 'POST':
        # Get the uploaded file
        uploaded_file = request.files['file']

        if uploaded_file.filename != '':
            # Save the uploaded PDF
            pdf_path = 'uploads/' + uploaded_file.filename
            uploaded_file.save(pdf_path)

            # Define parsing rules (for demonstration purposes, you may need to customize this)
            tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)

            # Combine extracted tables into a single DataFrame
            parsed_data = pd.concat(tables)

            # Export to Excel
            excel_output = BytesIO()
            with pd.ExcelWriter(excel_output, engine='openpyxl') as writer:
                parsed_data.to_excel(writer, sheet_name='ParsedData', index=False)
            excel_output.seek(0)

            return send_file(excel_output, attachment_filename='parsed_data.xlsx', as_attachment=True)

    return render_template('upload.html')

if __name__ == '__main__':
    app.run(debug=True)
