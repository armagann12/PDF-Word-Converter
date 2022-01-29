import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import PyPDF2 as pdf2
import docx

# create the root window
root = tk.Tk()
root.title('Tkinter Open File Dialog')
root.resizable(False, False)
root.geometry('250x250')

# Select a PDF
def select_file():
    filetypes = (
        ('pdf files', '*.pdf'),
        ('All files', '*.*')
    )
    
    filename = fd.askopenfilename(
        title='Open a file',
        filetypes=filetypes
        )
    
    showinfo(
        title='Selected File',
        message=filename
    )
    
    # Open PDF and Extract it
    with open(filename, mode='rb') as f:

        reader = pdf2.PdfFileReader(f, strict=False)
        if reader.isEncrypted:
            reader.decrypt("")

        page = reader.getPage(1)
        global extracted ##global important
        extracted = page.extractText()

# Get Docx name and create s
# Add Extract Data to Docx
def submit():
    doc=doc_name.get()
    document = docx.Document()
    document.add_paragraph(extracted) 
    document.save(doc) 


doc_name=tk.StringVar()    

name_label = tk.Label(root, text = 'Doc name with .docx', font=('calibre',10, 'bold'))
  
name_entry = tk.Entry(root,textvariable = doc_name, font=('calibre',10,'normal'))

sub_btn=tk.Button(root,text = 'Submit', command = submit)
     

# open button
open_button = ttk.Button(
    root,
    text='Select PDF',
    command=select_file
)

#Position

open_button.pack(expand=True)
name_label.pack(expand=True)
name_entry.pack(expand=True)
sub_btn.pack(expand=True)


# run the application
root.mainloop()
