from flask import Flask, request, render_template, send_file, flash, jsonify
from werkzeug.utils import secure_filename
import os
import logging
import io
import base64
from macro_parser import MacroParser
from pdf_generator import generate_pdf
from gemini_enhancer import enhance_explanation_with_gemini
from db import save_document, get_all_documents, get_document_by_id, get_all_macros, get_macros_by_document_id, get_macro_by_id

app = Flask(__name__)
app.secret_key = 'your_secret_key'

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
                with open(pdf_path, 'rb') as pdf_file:
                    pdf_data = pdf_file.read()
                
                # Save document and macros in the database
                document_id = save_document(filename, pdf_data, parsed_macros)
                logger.info(f"Document saved with ID: {document_id}")
                
                # Clean up: remove the uploaded and generated files
                if os.path.exists(filepath):
                    os.remove(filepath)
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)

                return send_file(io.BytesIO(pdf_data), attachment_filename=f"{filename}.pdf", as_attachment=True)
            except Exception as e:
                logger.error(f"Error processing file: {str(e)}", exc_info=True)
                flash(f"Error processing file: {str(e)}")
                return render_template('upload.html')
        else:
            flash('Invalid file type')
            return render_template('upload.html')
    
    return render_template('upload.html')

@app.route('/documents', methods=['GET'])
def view_all_documents():
    documents = get_all_documents()
    documents_list = [{'id': doc.id, 'name': doc.name, 'generated_pdf': base64.b64encode(doc.generated_pdf).decode('utf-8')} for doc in documents]
    return jsonify(documents_list)

@app.route('/documents/<int:document_id>', methods=['GET'])
def view_document_by_id(document_id):
    document = get_document_by_id(document_id)
    if document:
        return send_file(io.BytesIO(document.generated_pdf), attachment_filename=f"{document.name}.pdf", as_attachment=True)
    return jsonify({'error': 'Document not found'}), 404

@app.route('/macros', methods=['GET'])
def view_all_macros():
    macros = get_all_macros()
    macros_list = [{'id': macro.id, 'name': macro.name, 'document_id': macro.document_id, 'efficient': macro.efficient, 'flowchart': base64.b64encode(macro.flowchart).decode('utf-8') if macro.flowchart else None} for macro in macros]
    return jsonify(macros_list)

# @app.route('/macros/<int:document_id>', methods=['GET'])
# def view_macros_by_document_id(document_id):
#     macros = get_macros_by_document_id(document_id)
#     macros_list = [{'id': macro.id, 'name': macro.name, 'efficient': macro.efficient, 'flowchart': base64.b64encode(macro.flowchart).decode('utf-8') if macro.flowchart else None} for macro in macros]
#     return jsonify(macros_list)

@app.route('/macros/<int:macro_id>', methods=['GET'])
def view_macro_by_id(macro_id):
    macro = get_macro_by_id(macro_id)
    if macro:
        macro_data = {
            'id': macro.id,
            'name': macro.name,
            'document_id': macro.document_id,
            'efficient': macro.efficient,
            'flowchart': base64.b64encode(macro.flowchart).decode('utf-8') if macro.flowchart else None
        }
        return jsonify(macro_data)
    return jsonify({'error': 'Macro not found'}), 404

if __name__ == '__main__':
    app.run(debug=True)
