import os
import uno
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse
import sys

# FastAPI app
app = FastAPI()

# Debugging environment info on startup
@app.on_event("startup")
async def startup_event():
    print("Python executable:", sys.executable)
    print("Python path:", sys.path)
    print("Environment variables:", os.environ)

# Test UNO connection
def test_uno_connection():
    try:
        # Set the UNO environment variable for LibreOffice connection (update with correct path)
        libreoffice_path = "C:/Program Files/LibreOffice/program"  # Adjust to your LibreOffice path
        os.environ["URE_BOOTSTRAP"] = f"file:///{libreoffice_path}/fundamental.ini"

        # Initialize UNO
        local_context = uno.getComponentContext()
        print("UNO context successfully initialized!")
    except Exception as e:
        print(f"Failed to initialize UNO context: {e}")

# Function to convert XLSX to PDF
def export_pdf(input_file, output_file):
    try:
        # Initialize UNO
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context
        )
        context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        desktop = context.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", context
        )

        # Open the document
        print(f"Attempting to open file: {input_file}")
        document = desktop.loadComponentFromURL(f"file:///{input_file}", "_blank", 0, ())

        # Recalculate formulas
        document.calculateAll()

        # Export to PDF
        pdf_filter = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        pdf_filter.Name = "FilterName"
        pdf_filter.Value = "calc_pdf_Export"
        print(f"Exporting to: {output_file}")
        document.storeToURL(f"file:///{output_file}", (pdf_filter,))

        # Close document
        document.close(True)

        # Check if the PDF file exists after conversion
        if os.path.exists(output_file):
            print(f"PDF file created: {output_file}")
        else:
            print(f"PDF file not found: {output_file}")

    except Exception as e:
        print(f"Error: {e}")

# Endpoint for file conversion
@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    # Set up paths
    temp_dir = "C:/temp/"
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)  # Create directory if it doesn't exist

    input_path = os.path.join(temp_dir, file.filename)
    output_path = input_path.replace(".xlsx", ".pdf")

    # Save the uploaded file
    print(f"Saving uploaded file to: {input_path}")
    with open(input_path, "wb") as f:
        f.write(await file.read())

    # Log the input and output paths
    print(f"Input file saved at: {input_path}")
    print(f"Output file saved at: {output_path}")

    # Convert to PDF
    export_pdf(input_path, output_path)

    # Check if the PDF file was created
    if not os.path.exists(output_path):
        raise RuntimeError(f"File at path {output_path} does not exist.")

    # Return the generated PDF
    return FileResponse(output_path, filename=os.path.basename(output_path), media_type="application/pdf")

# Health check
@app.get("/health")
async def health_check():
    test_uno_connection()
    return {"status": "ok"}

if __name__ == "__main__":
    # Test the UNO connection during startup


    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
