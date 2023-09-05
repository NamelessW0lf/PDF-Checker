import PyPDF2
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import os
import pandas as pd

def calculate_coordinates(field_rect):
    x_left = field_rect[0]
    y_bottom = field_rect[1]
    x_right = field_rect[2]
    y_top = field_rect[3]
    
    x_center = (x_left + x_right) / 2
    y_center = (y_bottom + y_top) / 2
    
    x_width = x_right - x_left
    y_height = y_top - y_bottom
    
    x_top_left = x_left
    y_top_left = y_top
    
    x_top_right = x_right
    y_top_right = y_top
    
    x_bottom_left = x_left
    y_bottom_left = y_bottom
    
    x_bottom_right = x_right
    y_bottom_right = y_bottom
    
    return x_top_left, y_top_left, x_top_right, y_top_right, x_bottom_left, y_bottom_left, x_bottom_right, y_bottom_right, x_center, y_center, x_width, y_height

def extract_form_fields(pdf_path):
    form_fields = []
    
    pdf_file = open(pdf_path, 'rb')
    pdf_reader = PyPDF2.PdfFileReader(pdf_file)
    
    num_pages = pdf_reader.numPages
    
    for page_num in range(num_pages):
        page = pdf_reader.getPage(page_num)
        annotations = page['/Annots']
        
        for annotation in annotations:
            annotation_dict = annotation.getObject()
            if '/Subtype' in annotation_dict and annotation_dict['/Subtype'] == '/Widget':
                field_name = annotation_dict.get('/T', '')
                field_rect = annotation_dict['/Rect']
                
                # Koordinaten berechnen
                x_top_left, y_top_left, x_top_right, y_top_right, x_bottom_left, y_bottom_left, x_bottom_right, y_bottom_right, x_center, y_center, x_width, y_height = calculate_coordinates(field_rect)
                
                form_fields.append((field_name, x_top_left, y_top_left, x_top_right, y_top_right, x_bottom_left, y_bottom_left, x_bottom_right, y_bottom_right, x_center, y_center, x_width, y_height))
    
    pdf_file.close()
    return num_pages, form_fields

def browse_pdf():
    global current_pdf
    current_pdf = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if current_pdf:
        num_pages, form_fields = extract_form_fields(current_pdf)
        display_results(num_pages, form_fields)

def export_csv():
    global current_pdf
    csv_filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if csv_filename:
        _, form_fields = extract_form_fields(current_pdf)
        df = pd.DataFrame(form_fields, columns=["Feldname", "X-TopLeft", "Y-TopLeft", "X-TopRight", "Y-TopRight", "X-BottomLeft", "Y-BottomLeft", "X-BottomRight", "Y-BottomRight", "X-Center", "Y-Center", "X-Width", "Y-Height"])
        df.to_csv(csv_filename, index=False)
        tk.messagebox.showinfo("Erfolg", "CSV-Datei wurde erfolgreich exportiert.")

def export_excel():
    global current_pdf
    excel_filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if excel_filename:
        _, form_fields = extract_form_fields(current_pdf)
        df = pd.DataFrame(form_fields, columns=["Feldname", "X-TopLeft", "Y-TopLeft", "X-TopRight", "Y-TopRight", "X-BottomLeft", "Y-BottomLeft", "X-BottomRight", "Y-BottomRight", "X-Center", "Y-Center", "X-Width", "Y-Height"])
        df.to_excel(excel_filename, index=False)
        tk.messagebox.showinfo("Erfolg", "Excel-Datei wurde erfolgreich exportiert.")

def display_results(num_pages, form_fields):
    # Tabelle leeren
    for item in result_table.get_children():
        result_table.delete(item)
    
    # PDF-Datei-Anzeige aktualisieren und nur den Dateinamen anzeigen
    pdf_filename = os.path.basename(current_pdf)
    pdf_info_label.config(text=f"Aktuelle PDF: {pdf_filename}, Seiten: {num_pages}, Formularfelder: {len(form_fields)}")
    
    for field_name, x_top_left, y_top_left, x_top_right, y_top_right, x_bottom_left, y_bottom_left, x_bottom_right, y_bottom_right, x_center, y_center, x_width, y_height in form_fields:
        result_table.insert("", "end", values=(field_name, x_top_left, y_top_left, x_top_right, y_top_right, x_bottom_left, y_bottom_left, x_bottom_right, y_bottom_right, x_center, y_center, x_width, y_height))

# GUI erstellen
root = tk.Tk()
root.title("PDF Formularfelder Extractor")

# Button-Stil
button_style = ("Helvetica", 12)
file_button = tk.Button(root, text="PDF auswählen", command=browse_pdf, font=button_style)
file_button.pack(pady=10)

# PDF-Anzeige
pdf_info_label = tk.Label(root, text="Aktuelle PDF: Keine ausgewählt", font=("Helvetica", 12))
pdf_info_label.pack()

# Tabelle erstellen mit den neuen Spaltennamen
columns = ("Feldname", "X-TopLeft", "Y-TopLeft", "X-TopRight", "Y-TopRight", "X-BottomLeft", "Y-BottomLeft", "X-BottomRight", "Y-BottomRight", "X-Center", "Y-Center", "X-Width", "Y-Height")
result_table = ttk.Treeview(root, columns=columns, show="headings", height=10)
result_table.heading("Feldname", text="Feldname")
result_table.heading("X-TopLeft", text="X-TopLeft")
result_table.heading("Y-TopLeft", text="Y-TopLeft")
result_table.heading("X-TopRight", text="X-TopRight")
result_table.heading("Y-TopRight", text="Y-TopRight")
result_table.heading("X-BottomLeft", text="X-BottomLeft")
result_table.heading("Y-BottomLeft", text="Y-BottomLeft")
result_table.heading("X-BottomRight", text="X-BottomRight")
result_table.heading("Y-BottomRight", text="Y-BottomRight")
result_table.heading("X-Center", text="X-Center")
result_table.heading("Y-Center", text="Y-Center")
result_table.heading("X-Width", text="X-Width")
result_table.heading("Y-Height", text="Y-Height")
result_table.column("Feldname", width=200)
result_table.column("X-TopLeft", width=100)
result_table.column("Y-TopLeft", width=100)
result_table.column("X-TopRight", width=100)
result_table.column("Y-TopRight", width=100)
result_table.column("X-BottomLeft", width=100)
result_table.column("Y-BottomLeft", width=100)
result_table.column("X-BottomRight", width=100)
result_table.column("Y-BottomRight", width=100)
result_table.column("X-Center", width=100)
result_table.column("Y-Center", width=100)
result_table.column("X-Width", width=100)
result_table.column("Y-Height", width=100)
result_table.pack(padx=10, pady=10)

# Export-Buttons hinzufügen
export_csv_button = tk.Button(root, text="Exportiere als CSV", command=export_csv, font=button_style)
export_csv_button.pack(pady=10)

export_excel_button = tk.Button(root, text="Exportiere als Excel", command=export_excel, font=button_style)
export_excel_button.pack()

current_pdf = ""

root.mainloop()
