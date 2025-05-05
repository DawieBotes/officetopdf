import sys
import uno
import os
import re  # >>> MODIFIED <<<


def export_pdf(input_file, output_file, sheet_name=None):
    # Detect the file type
    ext = os.path.splitext(input_file)[1].lower()
    
    if ext not in ['.xlsx', '.xls', '.doc', '.docx']:
        raise ValueError(f"Unsupported file type: {ext}. Only .xlsx, xls, .doc, and .docx are supported.")
    
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
        if ext in ['.xlsx', '.xls']:
            

            if sheet_name:
                sheets = document.getSheets()
                if not sheets.hasByName(sheet_name):
                    raise ValueError(f"Sheet '{sheet_name}' not found in document.")
                
                # Add a temporary sheet to force state reset
                temp_sheet_name = "TempSheet"
                i = 0
                while sheets.hasByName(temp_sheet_name):
                    i += 1
                    temp_sheet_name = f"TempSheet_{i}"

                sheets.insertNewByName(temp_sheet_name, sheets.getCount())


                # Hide all sheets first
                for i in range(sheets.getCount()):
                    sheet = sheets.getByIndex(i)
                    sheet.setPropertyValue("IsVisible", False)

                # Show only the target sheet
                target_sheet = sheets.getByName(sheet_name)
                target_sheet.setPropertyValue("IsVisible", True)

                # Set target sheet as active
                controller = document.getCurrentController()
                controller.setActiveSheet(target_sheet)

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




def recalculate_excel(input_file, output_file):
    ext = os.path.splitext(input_file)[1].lower()
    if ext not in ['.xls', '.xlsx']:
        raise ValueError(f"Unsupported file type: {ext}. Only .xls and .xlsx are supported.")

    try:
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context
        )
        context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        desktop = context.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", context
        )

        document = desktop.loadComponentFromURL(
            f"file:///{input_file}", "_blank", 0, ()
        )

        # Recalculate formulas
        document.calculateAll()

        # Set the output filter to Excel
        output_ext = os.path.splitext(output_file)[1].lower()
        if output_ext == '.xlsx':
            filter_name = "Calc MS Excel 2007 XML"
        elif output_ext == '.xls':
            filter_name = "MS Excel 97"
        else:
            raise ValueError("Unsupported output format for recalculated Excel file.")

        filter_props = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        filter_props.Name = "FilterName"
        filter_props.Value = filter_name

        document.storeToURL(f"file:///{output_file}", (filter_props,))

        document.close(True)

    except uno.Exception as e:
        raise RuntimeError(f"LibreOffice UNO API error: {e}")
    except Exception as e:
        raise RuntimeError(f"An error occurred during recalculation: {e}")
    
if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage:")
        print("  python3 convert.py <input_file> <output_file>              # Convert to PDF")
        print("  python3 convert.py recalc <input_file> <output_file>       # Recalculate and save Excel")
        sys.exit(1)

    try:
        if sys.argv[1] == "recalc":
            input_file = sys.argv[2]
            output_file = sys.argv[3]
            recalculate_excel(input_file, output_file)
            print(f"Successfully recalculated {input_file} to {output_file}")
        else:
            input_file = sys.argv[1]
            output_file = sys.argv[2]
            sheet_name = sys.argv[3] if len(sys.argv) > 3 else None
            export_pdf(input_file, output_file, sheet_name)
            print(f"Successfully converted {input_file} to {output_file}")

    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)

