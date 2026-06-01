
### ⚙️ The Solution: The Word COM API

Since you are already relying on the Windows COM interface to communicate with ETABS, the most elegant and native solution is to use the COM interface to briefly command Microsoft Word in the background.

Once** **`docxtpl` finishes injecting the data and saves the file, your Python script will silently open the generated document via Word, trigger the TOC update, and save it again. The user will never even see Word open.

Here is the exact code block to implement this.

### 1. The TOC Updater Function

You will need the** **`pywin32` library (`pip install pywin32`). We will wrap this in a try/except block and an OS check so that it doesn't crash your development environment if you are testing on macOS/Linux.

**Python**

```
import os
import platform

def update_word_toc(file_path):
    """
    Silently opens Microsoft Word via COM to update the Table of Contents field codes.
    Skips execution if running on a non-Windows (Mock) environment.
    """
    if platform.system() != "Windows":
        print("DEV MODE: Skipping TOC Update (Requires Windows COM).")
        return

    # win32com is standard for controlling Office apps via Python
    import win32com.client 

    # Convert to absolute path (Word COM requires absolute paths)
    abs_path = os.path.abspath(file_path)

    try:
        # Launch a hidden instance of Word
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
      
        # Open the freshly generated document
        doc = word.Documents.Open(abs_path)
      
        # Update all Tables of Contents in the document
        if doc.TablesOfContents.Count > 0:
            for toc in doc.TablesOfContents:
                toc.Update()
            print("Table of Contents updated successfully.")
        else:
            print("No Table of Contents found in the document.")
          
        # Save and close cleanly
        doc.Close(SaveChanges=True)
      
    except Exception as e:
        print(f"Warning: Could not update TOC automatically. Error: {e}")
      
    finally:
        # Ensure the hidden Word process is terminated
        word.Quit()
```

### 2. Integration into the Pipeline

This function gets called at the absolute end of your workflow, immediately after** **`docxtpl` saves the** **`.docx` file.

**Python**

```
from docxtpl import DocxTemplate

def generate_report(context_data, template_path, output_path):
    print("Injecting ETABS data into Word Template...")
  
    # 1. Initialize Template
    doc = DocxTemplate(template_path)
  
    # 2. Render all tables, text, and images
    doc.render(context_data)
  
    # 3. Save the raw XML-modified document
    doc.save(output_path)
    print(f"Document saved to {output_path}.")
  
    # 4. Trigger the Word COM engine to finalize page numbers and TOC
    print("Calculating final page numbers...")
    update_word_toc(output_path)
  
    print("Report Generation Complete! 🚀")
```

### ⚠️ A Warning on Word Field Codes

For this to work flawlessly, the Table of Contents in your base** **`Report_Template_Tagged.docx` must be a** ** **native Microsoft Word automatic table** .

If you just typed out a list that looks like a TOC, the COM command will do nothing. You must insert it in Word by going to** ** **References > Table of Contents > Automatic Table** . Word uses hidden field codes (like** **`TOC \o "1-3" \h \z \u`) to identify these blocks. The Python script looks for these exact field codes to trigger the refresh.
