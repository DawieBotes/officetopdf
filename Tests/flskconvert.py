import os
import shutil
import uno
from flask import Flask, request, jsonify
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Configure upload folder and allowed extensions
UPLOAD_FOLDER = 'C:/temp/uploads'
OUTPUT_FOLDER = 'C:/temp/output'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Set the URE_BOOTSTRAP for UNO initialization
libreoffice_path = "C:/Program Files/LibreOffice/program"  # Adjust this path based on your setup
os.environ["URE_BOOTSTRAP"] = f"file:///{libreoffice_path}/fundamental.ini"

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    # Save the uploaded file
    input_filename = secure_filename(file.filename)
    input_path = os.path.join(UPLOAD_FOLDER, input_filename)
    file.save(input_path)
    
    output_filename = input_filename.replace('.xlsx', '.pdf')
    output_path = os.path.join(OUTPUT_FOLDER, output_filename)
    
    try:
        # Initialize UNO and perform conversion
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context
        )
        context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        desktop = context.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", context
        )

        # Open the document
        document = desktop.loadComponentFromURL(f"file:///{input_path}", "_blank", 0, ())

        # Recalculate formulas
        document.calculateAll()

        # Export to PDF
        pdf_filter = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        pdf_filter.Name = "FilterName"
        pdf_filter.Value = "calc_pdf_Export"
        document.storeToURL(f"file:///{output_path}", (pdf_filter,))

        # Close document
        document.close(True)

        return jsonify({"message": f"File converted successfully!", "output_file": output_path})

    except Exception as e:
        return jsonify({"error": f"Conversion failed: {e}"}), 500


@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({"status": "API is up and running!"})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
