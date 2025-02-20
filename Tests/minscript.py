import os
import uno
from fastapi import FastAPI, File, UploadFile
import shutil

app = FastAPI()

# Ensure UNO is initialized at app startup
@app.on_event("startup")
async def startup_event():
    try:
        # Set the UNO environment variable (Ensure it points to the correct fundamental.ini file)
        libreoffice_path = "C:/Program Files/LibreOffice/program"  # Adjust this path as needed
        os.environ["URE_BOOTSTRAP"] = f"file:///{libreoffice_path}/fundamental.ini"

        # Ensure the UNO environment is ready before trying to use it
        local_context = uno.getComponentContext()
        print("UNO context successfully initialized!")
    except Exception as e:
        print(f"Failed to initialize UNO context: {e}")
        raise e

# Test UNO initialization via a route
@app.get("/test_uno")
async def test_uno():
    try:
        libreoffice_path = "C:/Program Files/LibreOffice/program"  # Adjust this path as needed
        os.environ["URE_BOOTSTRAP"] = f"file:///{libreoffice_path}/fundamental.ini"
        
        # Initialize UNO
        local_context = uno.getComponentContext()
        return {"message": "UNO context initialized successfully!"}
    except Exception as e:
        return {"error": str(e)}

# Convert the uploaded Excel file to PDF
@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    input_path = f"C:/temp/{file.filename}"
    output_path = f"C:/temp/{file.filename.replace('.xlsx', '.pdf')}"
    
    # Save the uploaded file to disk
    with open(input_path, "wb") as f:
        shutil.copyfileobj(file.file, f)
    print(f"Input file saved at: {input_path}")

    # Try to initialize UNO and perform the conversion
    try:
        # Set the UNO environment variable for LibreOffice
        libreoffice_path = "C:/Program Files/LibreOffice/program"  # Adjust this path as needed
        os.environ["URE_BOOTSTRAP"] = f"file:///{libreoffice_path}/fundamental.ini"

        # Initialize UNO context
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context
        )
        context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        desktop = context.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", context
        )

        # Open the document
        document = desktop.loadComponentFromURL(
            f"file:///{input_path}", "_blank", 0, ()
        )

        # Recalculate formulas
        document.calculateAll()

        # Export to PDF
        pdf_filter = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        pdf_filter.Name = "FilterName"
        pdf_filter.Value = "calc_pdf_Export"
        document.storeToURL(f"file:///{output_path}", (pdf_filter,))

        # Close document
        document.close(True)
        print(f"Output file saved at: {output_path}")

    except Exception as e:
        return {"error": f"Conversion failed: {e}"}

    return {"message": f"File converted successfully! Output: {output_path}"}

# Health check route to test if the API is responsive
@app.get("/health")
async def health_check():
    return {"status": "API is up and running!"}

