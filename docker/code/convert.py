import sys
import uno

def export_pdf(input_file, output_file):
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

    # Recalculate formulas
    document.calculateAll()

    # Export to PDF
    pdf_filter = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    pdf_filter.Name = "FilterName"
    pdf_filter.Value = "calc_pdf_Export"
    document.storeToURL(f"file:///{output_file}", (pdf_filter,))

    # Close document
    document.close(True)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python3 convert.py <input_file> <output_file>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    export_pdf(input_file, output_file)
