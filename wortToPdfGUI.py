from tkinter import Tk, Label, Button, filedialog, Entry, messagebox, Frame
from PyPDF2 import PdfMerger
from docx2pdf import convert as convert_docx2pdf
import os

selected_conversion_files = []
selected_merge_files = []
output_file_entry = None

def convert_to_pdf(files):
    pdf_files = []
    for file in files:
        try:
            if file.endswith(".docx"):
                convert_docx2pdf(file)
                pdf_file = file.replace(".docx", ".pdf") 
                if os.path.exists(pdf_file): 
                    pdf_files.append(pdf_file)  
            else:
                print(f"Unsupported file format: {file}")
        except FileNotFoundError:
            print(f"File not found: {file}")
        except Exception as e:
            print(f"An error occurred while converting {file} to PDF: {str(e)}")
            print("Conversion failed.")

    if pdf_files:
        print(f"PDF files: {pdf_files}")
    else:
        print("No PDF files converted.")

    return pdf_files

def merge_pdfs(pdf_files, output_file):
    merger = PdfMerger()
    for pdf_file in pdf_files:
        merger.append(pdf_file)
    merger.write(output_file)
    merger.close()

def remove_file(file_name, file_frame):
    file_frame.destroy()
    if file_name in selected_conversion_files:
        selected_conversion_files.remove(file_name)
    conversion_files_label.config(text=f"{len(selected_conversion_files)} file(s) selected:")

def select_conversion_files():
    global selected_conversion_files
    selected_conversion_files = list(filedialog.askopenfilenames(filetypes=[("Word Files", "*.docx")]))
    if selected_conversion_files:
        for file in selected_conversion_files:
            file_name = os.path.basename(file)
            file_frame = Frame(conversion_frame)
            file_frame.pack(anchor='w', padx=5, pady=2)
            file_label = Label(file_frame, text=file_name)
            file_label.pack(side='left')
            remove_button = Button(file_frame, text="x", command=lambda f=file_name, l=file_frame: remove_file(f, l))
            remove_button.pack(side='left')
        conversion_files_label.config(text=f"{len(selected_conversion_files)} file(s) selected:")




def select_merge_files():
    global selected_merge_files
    selected_merge_files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
    if selected_merge_files:
        merge_files_label.config(text=f"{len(selected_merge_files)} file(s) selected.")

def select_output():
    output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
    if output_file:
        output_file_entry.delete(0, 'end')
        output_file_entry.insert(0, output_file)

def convert_files():
    global selected_conversion_files
    if not selected_conversion_files:
        messagebox.showerror("Error", "No files selected for conversion.")
        return

    initial_word_files_count = len([file for file in selected_conversion_files if file.endswith(".docx")])

    converted_files = []
    for file in selected_conversion_files:
        if os.path.exists(file):  # Check if the file still exists
            pdf_files = convert_to_pdf([file])
            if pdf_files:
                converted_files.extend(pdf_files)

    converted_count = len(converted_files)
    if converted_count > 0:
        messagebox.showinfo("Success", f"Conversion completed. {converted_count} PDF file(s) created.")
    else:
        messagebox.showwarning("Warning", "No files converted to PDF.")

    # Update the selected_conversion_files list with only the files that were successfully converted
    selected_conversion_files = [file for file in selected_conversion_files if file.replace(".docx", ".pdf") in converted_files]


def merge_files():
    global selected_merge_files
    if not selected_merge_files:
        messagebox.showerror("Error", "No files selected for merging.")
        return

    output_file = output_file_entry.get()
    if not output_file:
        messagebox.showerror("Error", "No output file selected.")
        return

    merge_pdfs(selected_merge_files, output_file)
    messagebox.showinfo("Success", "PDF files merged successfully.")

# Create the main window
root = Tk()
root.title("File Converter and Merger")
root.geometry("700x400")

# Create frames for each section
conversion_frame = Frame(root, width=root.winfo_width() // 2, height=root.winfo_height(), bg='lightgrey')
conversion_frame.pack(side='left', fill='both', expand=True)

merge_frame = Frame(root, width=root.winfo_width() // 2, height=root.winfo_height(), bg='lightblue')
merge_frame.pack(side='right', fill='both', expand=True)

# Widgets for conversion frame
conversion_files_label = Label(conversion_frame, text="No files selected for conversion.")
conversion_files_label.pack()

select_conversion_files_button = Button(conversion_frame, text="Select Word Files", command=select_conversion_files)
select_conversion_files_button.pack()

convert_button = Button(conversion_frame, text="Convert to PDF", command=convert_files)
convert_button.pack()

# Widgets for merge frame
merge_files_label = Label(merge_frame, text="No files selected for merging.")
merge_files_label.pack()

select_merge_files_button = Button(merge_frame, text="Select PDF Files", command=select_merge_files)
select_merge_files_button.pack()

output_label = Label(merge_frame, text="Output PDF file name:")
output_label.pack()

output_file_entry = Entry(merge_frame, width=50)
output_file_entry.pack()

output_button = Button(merge_frame, text="Save As...", command=select_output)
output_button.pack()

merge_button = Button(merge_frame, text="Merge PDF Files", command=merge_files)
merge_button.pack()

root.mainloop()
