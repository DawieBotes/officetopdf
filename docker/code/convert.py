import sys
import uno
import os

def export_pdf(input_file, output_file):
    # Detect the file type
    ext = os.path.splitext(input_file)[1].lower()
    
    if ext not in ['.xlsx', '.doc', '.docx']:
        raise ValueError(f"Unsupported file type: {ext}. Only .xlsx, .doc, and .docx are supported.")
    
    try:
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
            f"file:///{input_file}", "_blank", 0, ()
        )

        # Recalculate formulas for Excel files
        if ext == '.xlsx':
            document.calculateAll()

        # Export to PDF
        pdf_filter = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        pdf_filter.Name = "FilterName"
        pdf_filter.Value = "writer_pdf_Export" if ext in ['.doc', '.docx'] else "calc_pdf_Export"
        document.storeToURL(f"file:///{output_file}", (pdf_filter,))

        # Close document
        document.close(True)

    except uno.Exception as e:
        raise RuntimeError(f"LibreOffice UNO API error: {e}")
    except Exception as e:
        raise RuntimeError(f"An error occurred during the conversion process: {e}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python3 convert.py <input_file> <output_file>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    try:
        export_pdf(input_file, output_file)
        print(f"Successfully converted {input_file} to {output_file}")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)
