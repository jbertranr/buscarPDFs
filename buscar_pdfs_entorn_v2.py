import pdfplumber
import os
import csv
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Frame, Button, Label, Entry, Scrollbar, Style, Treeview

def buscar_pdfs():
    folder_path = folder_path_entry.get()
    search_terms = search_terms_entry.get().strip().split(",")
    if not folder_path or not os.path.isdir(folder_path):
        messagebox.showerror("Error", "La ruta de la carpeta no és vàlida.")
        return
    if not search_terms:
        messagebox.showerror("Error", "Introdueix almenys una paraula clau.")
        return

    # Neteja els resultats existents
    for row in result_table.get_children():
        result_table.delete(row)

    pdf_files = [file for file in os.listdir(folder_path) if file.endswith(".pdf")]
    total_files = len(pdf_files)

    # Acumular resultats
    accumulated_results = []

    for i, file in enumerate(pdf_files, start=1):
        processing_label.config(text=f"Processant: {file} ({i}/{total_files})")
        processing_label.update()

        pdf_path = os.path.join(folder_path, file)
        with pdfplumber.open(pdf_path) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                text = page.extract_text().replace("\n", " ") if page.extract_text() else ""
                if text:
                    for term in search_terms:
                        if term in text:
                            start_idx = max(0, text.find(term) - 50)
                            end_idx = min(len(text), text.find(term) + 50 + len(term))
                            fragment = text[start_idx:end_idx].replace(term, term.upper())

                            # Acumula resultats
                            accumulated_results.append(("Sí", file, page_number, term, fragment))

    # Actualitza la interfície gràfica un cop acabada la cerca
    for result in accumulated_results:
        item = result_table.insert("", "end", values=result)
        update_row_style(item)

    processing_label.config(text="Processament complet!")

def toggle_checkbox(event):
    """Canvia l'estat del checkbox (Sí/No) a la primera columna i actualitza l'estil de la fila."""
    region = result_table.identify("region", event.x, event.y)
    if region == "cell":
        column = result_table.identify_column(event.x)
        item = result_table.identify_row(event.y)
        if column == "#1" and item:  # Primera columna (Seleccionar)
            current_value = result_table.item(item)["values"][0]
            new_value = "Sí" if current_value == "No" else "No"
            result_table.item(item, values=(new_value, *result_table.item(item)["values"][1:]))
            update_row_style(item)

def update_row_style(item):
    """Actualitza l'estil d'una fila segons si està seleccionada o no."""
    values = result_table.item(item)["values"]
    if values[0] == "No":
        result_table.tag_configure("unselected", foreground="#888")
        result_table.item(item, tags=("unselected",))
    else:
        result_table.tag_configure("selected", background="white")
        result_table.item(item, tags=("selected",))

def guardar_csv():
    output_csv = os.path.join(folder_path_entry.get(), "resultats_busqueda.csv")

    # Només guarda resultats seleccionats (checkbox marcat com a "Sí")
    filtered_results = [
        result_table.item(item)["values"]
        for item in result_table.get_children()
        if result_table.item(item)["values"][0] == "Sí"
    ]

    if not filtered_results:
        messagebox.showerror("Error", "No hi ha resultats seleccionats per guardar.")
        return

    try:
        with open(output_csv, mode='w', newline='', encoding='utf-8') as csv_file:
            fieldnames = ["Nom del fitxer", "Pàgina", "Paraula clau", "Fragment"]
            writer = csv.DictWriter(csv_file, fieldnames=fieldnames, quoting=csv.QUOTE_ALL, quotechar='"')
            writer.writeheader()

            for row in filtered_results:
                writer.writerow({
                    "Nom del fitxer": row[1],
                    "Pàgina": row[2],
                    "Paraula clau": row[3],
                    "Fragment": row[4]
                })

        messagebox.showinfo("Guardat", f"Fitxer CSV guardat a: {output_csv}")
    except Exception as e:
        messagebox.showerror("Error", f"No s'ha pogut guardar el fitxer: {e}")

def seleccionar_carpeta():
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta amb PDFs")
    folder_path_entry.delete("0", "end")
    folder_path_entry.insert("0", carpeta)

# Crear interfície gràfica
root = tk.Tk()
root.title("Cercador de PDFs")
root.geometry("1000x700")

style = Style()
style.configure("Processing.TLabel", foreground="blue", background="#f0f0f0", font=("Helvetica", 10))

# Títol
title_label = Label(root, text="Cercador de paraules clau en PDFs", font=("Helvetica", 20, "bold"))
title_label.pack(pady=(20, 10))

# Descripció
description_label = Label(root, text="Aquest programa permet buscar paraules clau específiques dins de fitxers PDF. "
                                     "Introdueix una carpeta amb fitxers PDF i paraules clau separades per comes. "
                                     "Les coincidències es mostraran amb les paraules clau ressaltades.",
                           font=("Helvetica", 12), wraplength=750, justify="center")
description_label.pack(pady=(0, 20))

# Carpeta
folder_label = Label(root, text="Carpeta amb PDFs:")
folder_label.pack(anchor="w", padx=20)
folder_frame = Frame(root)
folder_frame.pack(fill="x", padx=20, pady=(0, 10))
folder_path_entry = Entry(folder_frame, width=60)
folder_path_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
folder_button = Button(folder_frame, text="Seleccionar carpeta", command=seleccionar_carpeta)
folder_button.pack(side="right")

# Paraules clau
search_label = Label(root, text="Paraules clau (separades per comes):")
search_label.pack(anchor="w", padx=20)
search_terms_entry = Entry(root, width=80)
search_terms_entry.pack(fill="x", padx=20, pady=(0, 10))

# Botó de cerca
search_button = Button(root, text="Cercar", command=buscar_pdfs)
search_button.pack(pady=10)

# Etiqueta de processament
processing_label = Label(root, text="", style="Processing.TLabel")
processing_label.pack(pady=5)

# Resultats
result_label = Label(root, text="Resultats:")
result_label.pack(anchor="w", padx=20)

result_frame = Frame(root)
result_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))

# Crear Treeview per mostrar els resultats amb checkboxs
columns = ("Seleccionar", "Nom del fitxer", "Pàgina", "Paraula clau", "Fragment")
result_table = Treeview(result_frame, columns=columns, show="headings", selectmode="browse")
result_table.pack(side="left", fill="both", expand=True)

# Afegir capçaleres
result_table.heading("Seleccionar", text="✔")
result_table.column("Seleccionar", width=30, anchor="center")
result_table.heading("Nom del fitxer", text="Nom del fitxer")
result_table.column("Nom del fitxer", width=200, anchor="w")
result_table.heading("Pàgina", text="Pàgina")
result_table.column("Pàgina", width=20, anchor="center")
result_table.heading("Paraula clau", text="Paraula clau")
result_table.column("Paraula clau", width=100, anchor="w")
result_table.heading("Fragment", text="Fragment")
result_table.column("Fragment", width=400, anchor="w")

# Afegir scrollbar
scrollbar = Scrollbar(result_frame, orient="vertical", command=result_table.yview)
scrollbar.pack(side="right", fill="y")
result_table.config(yscrollcommand=scrollbar.set)

# Afegir esdeveniment per fer clic al checkbox
result_table.bind("<Button-1>", toggle_checkbox)

# Botó de guardar
save_button = Button(root, text="Guardar CSV", command=guardar_csv)
save_button.pack(pady=10)

# Peu
footer_label = Label(root, text="Ajuntament de Mataró - SSIT - Secció Aplicacions", font=("Helvetica", 10))
footer_label.pack(side="bottom", fill="x", pady=(10, 0))

root.mainloop()
