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

    # Get optional sheet name from query string
    sheet_name = request.args.get('sheet')

    # Generate unique filename
    unique_filename = f"{uuid.uuid4()}{os.path.splitext(file.filename)[1]}"
    input_path = os.path.join(UPLOAD_FOLDER, unique_filename)
    output_path = input_path.rsplit(".", 1)[0] + ".pdf"

    file.save(input_path)

    # Build subprocess command
    command = ["python3", "/app/code/convert.py", input_path, output_path]
    if sheet_name:
        command.append(sheet_name)

    try:
        result = subprocess.run(
            command,
            check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True
        )
    except subprocess.CalledProcessError as e:
        return jsonify({"error": "Conversion failed", "details": e.stderr}), 500

    if not os.path.exists(output_path):
        return jsonify({"error": "Conversion failed: Output file not created"}), 500

    response = send_file(output_path, as_attachment=True)

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


@app.route('/recalc', methods=['POST'])
def recalc():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    # Validate file extension
    allowed_extensions = ['.xlsx', '.xls']
    if not any(file.filename.lower().endswith(ext) for ext in allowed_extensions):
        return jsonify({"error": f"Only {', '.join(allowed_extensions)} files are allowed"}), 400

    unique_filename = f"{uuid.uuid4()}{os.path.splitext(file.filename)[1]}"
    input_path = os.path.join(UPLOAD_FOLDER, unique_filename)
    output_path = input_path.rsplit(".", 1)[0] + "_recalc" + os.path.splitext(file.filename)[1]

    file.save(input_path)

    try:
        result = subprocess.run(
            ["python3", "/app/code/convert.py", "recalc", input_path, output_path],
            check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True
        )
    except subprocess.CalledProcessError as e:
        return jsonify({"error": "Recalculation failed", "details": e.stderr}), 500

    if not os.path.exists(output_path):
        return jsonify({"error": "Recalculation failed: Output file not created"}), 500

    response = send_file(output_path, as_attachment=True)

    try:
        os.remove(input_path)
        os.remove(output_path)
    except Exception as e:
        print(f"Error deleting files: {e}")

    return response



if __name__ == '__main__':
    app.run(host='0.0.0.0', port=port)
