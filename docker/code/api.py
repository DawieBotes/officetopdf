import os
from flask import Flask, request, send_file, jsonify
import subprocess
import uuid

app = Flask(__name__)

# Get port from environment variable, default to 5000 if not set
port = os.getenv("PORT", 5000)

UPLOAD_FOLDER = "/app/files/"
OUTPUT_FOLDER = "/app/files/"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    # Validate file extension
    allowed_extensions = ['.xlsx', '.xls', '.doc', '.docx']
    if not any(file.filename.lower().endswith(ext) for ext in allowed_extensions):
        return jsonify({"error": f"Only {', '.join(allowed_extensions)} files are allowed"}), 400

    # Generate a unique filename to avoid conflicts
    unique_filename = f"{uuid.uuid4()}{os.path.splitext(file.filename)[1]}"
    input_path = os.path.join(UPLOAD_FOLDER, unique_filename)
    output_path = input_path.rsplit(".", 1)[0] + ".pdf"

    # Save the uploaded file
    file.save(input_path)

    # Run the existing convert.py script
    try:
        result = subprocess.run(
            ["python3", "/app/code/convert.py", input_path, output_path],
            check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True
        )
    except subprocess.CalledProcessError as e:
        return jsonify({"error": "Conversion failed", "details": e.stderr}), 500

    # Ensure the output file exists before attempting to send it
    if not os.path.exists(output_path):
        return jsonify({"error": "Conversion failed: Output file not created"}), 500

    response = send_file(output_path, as_attachment=True)

    # Cleanup: Delete both input and output files after sending the response
    try:
        os.remove(input_path)
        os.remove(output_path)
    except Exception as e:
        print(f"Error deleting files: {e}")

    return response

@app.route('/health', methods=['GET'])
def health_check():
    """Simple health check endpoint"""
    return jsonify({"status": "healthy"}), 200

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=port)
