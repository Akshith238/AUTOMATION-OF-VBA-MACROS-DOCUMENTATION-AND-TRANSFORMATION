from flask import Flask, request, render_template, send_file, flash
from werkzeug.utils import secure_filename
import os
import logging
from macro_parser import MacroParser
from pdf_generator import generate_pdf
from gemini_enhancer import enhance_explanation_with_gemini

app = Flask(__name__)
app.secret_key = '' 

UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
ALLOWED_EXTENSIONS = {'xlsm', 'xls'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Set up logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Ensure the upload folder exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return render_template('upload.html')
        
        file = request.files['file']
        
        if file.filename == '':
            flash('No selected file')
            return render_template('upload.html')
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            logger.info(f"File saved to: {filepath}")
            
            try:
                # Process the file
                parser = MacroParser()
                parser.load_from_excel(filepath)
                parsed_macros = parser.parse_macros()
                logger.info(f"Parsed macros: {parsed_macros}")
                
                logic_explanations = parser.extract_functional_logic(parsed_macros)
                logger.info(f"Logic explanations: {logic_explanations}")
                
                # Enhance explanations with Gemini
                enhanced_explanations = []
                for explanation in logic_explanations:
                    logger.info(f"Processing explanation: {explanation}")
                    enhanced = enhance_explanation_with_gemini(str(explanation))
                    enhanced_explanations.append(enhanced)
                
                logger.info(f"Enhanced explanations: {enhanced_explanations}")
                
                functional_documentation = parser.generate_functional_documentation(enhanced_explanations)
                
                # Generate PDF
                pdf_path = generate_pdf(functional_documentation)
                
                return send_file(pdf_path, as_attachment=True)
            except Exception as e:
                logger.error(f"Error processing file: {str(e)}", exc_info=True)
                flash(f"Error processing file: {str(e)}")
                return render_template('upload.html')
            finally:
                # Clean up: remove the uploaded file
                if os.path.exists(filepath):
                    os.remove(filepath)
        else:
            flash('Invalid file type')
            return render_template('upload.html')
    
    return render_template('upload.html')

if __name__ == '__main__':
    app.run(debug=True)