from flask import Flask, request, render_template, send_file
import pandas as pd
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import os

app = Flask(__name__)


PDF_DIR = os.path.join(os.path.dirname(__file__), 'static', 'pdf')

@app.route('/')
def upload_form():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        
        if 'file' not in request.files:
            return 'No file part'

        file = request.files['file']

        
        if file.filename == '':
            return 'No selected file'

        
        df = pd.read_csv(file)

        
        form_date = request.form['date']
        if form_date:
            date_obj = datetime.strptime(form_date, '%Y-%m-%d')  # Adjust format as per your input
        else:
            date_obj = datetime.now()
        df['Date'] = date_obj.strftime('%d.%m.%Y')

        
        os.makedirs(PDF_DIR, exist_ok=True)

        
        pdf_filename = f'output_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf'
        pdf_path = os.path.join(PDF_DIR, pdf_filename)

        # Create PDF
        pdf = SimpleDocTemplate(pdf_path, pagesize=letter)
        elements = []

        
        styles = getSampleStyleSheet()
        heading = Paragraph('Forecasting Power Report', styles['Title'])
        elements.append(heading)
        elements.append(Spacer(1, 12))

        # Convert DataFrame to list of lists
        data_list = [df.columns.to_list()] + df.values.tolist()

        # Create table with data
        table = Table(data_list)

        # Add style to the table
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ])
        table.setStyle(style)

        elements.append(table)
        pdf.build(elements)

        # Serve the generated PDF for download
        response = send_file(pdf_path, as_attachment=True)

        # Delay the file deletion to ensure it is closed properly
        @response.call_on_close
        def cleanup():
            try:
                os.remove(pdf_path)
            except Exception as e:
                print(f'Error deleting file: {str(e)}')

        return response

    except Exception as e:
        return f'Error: {str(e)}'

if __name__ == '__main__':
    app.run(debug=True)
