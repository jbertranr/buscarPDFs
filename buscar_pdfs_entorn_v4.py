#import pdfplumber
import fitz
import os
import csv
import tkinter as tk
import threading
import datetime
from tkinter import filedialog, messagebox
from tkinter.ttk import Frame, Button, Label, Entry, Scrollbar, Style, Treeview
from tkinter.filedialog import asksaveasfilename
import gc
import unicodedata
import re

def netejar_text(text):
    """Elimina caràcters estranys i neteja el text extret."""
    # Normalitza el text per eliminar diacrítics i caràcters estranys
    text = unicodedata.normalize("NFKD", text)
    # Elimina caràcters no imprimibles
    text = "".join(c for c in text if c.isprintable())
    # Substitueix espais múltiples o salts de línia per un únic espai
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def process_pdf_with_fitz(pdf_path, search_terms):
    """Processa un PDF amb PyMuPDF i cerca termes específics."""
    results = []  # Emmagatzema els resultats de les coincidències
    doc = fitz.open(pdf_path)  # Obre el PDF

    for page_number, page in enumerate(doc, start=1):
        text = page.get_text("text")  # Extreu tot el text de la pàgina
        text = netejar_text(text)  # Neteja el text extret
        if text:
            # Normalitza el text per fer la cerca independent de majúscules/minúscules
            text_lower = text.lower()
            for term in search_terms:
                term_lower = term.lower()  # Normalitza el terme
                if term_lower in text_lower:
                    start_idx = text_lower.find(term_lower)
                    start_fragment = max(0, start_idx - 50)  # Mantenir 50 caràcters abans
                    end_fragment = min(len(text), start_idx + 50 + len(term_lower))  # Mantenir 50 caràcters després
                    fragment = text[start_fragment:end_fragment].replace(term_lower, term.upper())
                    results.append((pdf_path, page_number, term, fragment))

    doc.close()  # Tanca el PDF
    return results

def buscar_pdfs():
    global stop_search
    stop_search = False  # Restablir l'estat de cerca

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

    # Comprovar si cal buscar en subdirectoris
    if search_in_subdirs.get():
        pdf_files = [os.path.join(root, file) for root, _, files in os.walk(folder_path) for file in files if file.endswith(".pdf")]
    else:
        pdf_files = [os.path.join(folder_path, file) for file in os.listdir(folder_path) if file.endswith(".pdf")]

    total_files = len(pdf_files)

    for i, pdf_path in enumerate(pdf_files, start=1):
        if stop_search:
            break

        file = os.path.basename(pdf_path)  # Nom del fitxer sense la ruta
        pdf_path_norm = os.path.normpath(pdf_path)  # Convertir ruta al format Windows
        processing_label.config(text=f"Processant: {file} ({i}/{total_files})")
        processing_label.update()

        # Processa el PDF amb fitz
        results = process_pdf_with_fitz(pdf_path_norm, search_terms)

        # Afegeix els resultats a la taula
        for result in results:
            file_name, page_number, term, fragment = result
            result_table.insert("", "end", values=("Sí", file_name, page_number, term, fragment, pdf_path_norm))

        # Alliberar memòria després de processar el PDF
        gc.collect()

    if not stop_search:
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

def iniciar_cerca():
    cerca_thread = threading.Thread(target=buscar_pdfs)
    cerca_thread.start()   
     
def guardar_csv():
    # Obtenir la data i l’hora actual
    now = datetime.datetime.now()
    timestamp = now.strftime("%Y%m%d_%H%M%S")  # Format: YYYYMMDD_HHMMSS

    # Generar el nom proposat per al fitxer
    default_filename = f"resultats_busqueda_{timestamp}.csv"

    # Obrir la finestra de diàleg per seleccionar on guardar el fitxer
    output_csv = asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("Fitxers CSV", "*.csv")],
        initialfile=default_filename,
        title="Guardar com"
    )

    # Si l'usuari cancel·la, sortir de la funció
    if not output_csv:
        return

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
            fieldnames = ["Nom del fitxer", "Pàgina", "Paraula clau", "Fragment", "Ruta completa"]
            writer = csv.DictWriter(csv_file, fieldnames=fieldnames, quoting=csv.QUOTE_ALL, quotechar='"')
            writer.writeheader()

            for row in filtered_results:
                writer.writerow({
                    "Nom del fitxer": row[1],
                    "Pàgina": row[2],
                    "Paraula clau": row[3],
                    "Fragment": row[4],
                    "Ruta completa": row[5]
                })

        messagebox.showinfo("Guardat", f"Fitxer CSV guardat a: {output_csv}")
    except Exception as e:
        messagebox.showerror("Error", f"No s'ha pogut guardar el fitxer: {e}")

def seleccionar_carpeta():
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta amb PDFs")
    folder_path_entry.delete("0", "end")
    folder_path_entry.insert("0", carpeta)
    actualitzar_comptador_fitxers()

def aturar_cerca():
    global stop_search
    stop_search = True
    processing_label.config(text="Cerca aturada!")

def actualitzar_comptador_fitxers():
    folder_path = folder_path_entry.get()
    if not folder_path or not os.path.isdir(folder_path):
        file_count_label.config(text="Fitxers a processar: 0")
        return

    # Comptar els fitxers segons l'opció de subdirectoris
    if search_in_subdirs.get():
        pdf_files = [os.path.join(root, file) for root, _, files in os.walk(folder_path) for file in files if file.endswith(".pdf")]
    else:
        pdf_files = [os.path.join(folder_path, file) for file in os.listdir(folder_path) if file.endswith(".pdf")]

    total_files = len(pdf_files)
    file_count_label.config(text=f"Fitxers a processar: {total_files}")


def seleccionar_carpeta():
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta amb PDFs")
    folder_path_entry.delete("0", "end")
    folder_path_entry.insert("0", carpeta)
    actualitzar_comptador_fitxers()



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

folder_path_entry.bind("<FocusOut>", lambda e: actualitzar_comptador_fitxers())


# Paraules clau
search_label = Label(root, text="Paraules clau (separades per comes):")
search_label.pack(anchor="w", padx=20)
search_terms_entry = Entry(root, width=80)
search_terms_entry.pack(fill="x", padx=20, pady=(0, 10))

# Crear un frame contenidor
line_frame = Frame(root)
line_frame.pack(anchor="w", padx=20, pady=(0, 10))

file_count_label = Label(line_frame, text="Fitxers a processar: 0", font=("Helvetica", 10))
file_count_label.pack(side="left")

# Variable per al checkbox
search_in_subdirs = tk.BooleanVar(value=False)

# Vincular el canvi del checkbox amb el comptador
search_in_subdirs.trace_add("write", lambda *args: actualitzar_comptador_fitxers())

# Checkbox per a l'opció de buscar en subdirectoris
subdirs_checkbox = tk.Checkbutton(line_frame, text="Buscar en subdirectoris", variable=search_in_subdirs)
subdirs_checkbox.pack(side="left", padx=20)

# Botó de cerca
search_button = Button(line_frame, text="Cercar", command=iniciar_cerca)
search_button.pack(side="left", padx=20)

# Botó d'aturar la cerca
stop_button = Button(line_frame, text="Aturar", command=aturar_cerca)
stop_button.pack(side="left", padx=20)

# Etiqueta de processament
processing_label = Label(root, text="", style="Processing.TLabel")
processing_label.pack(pady=5)

# Resultats
result_label = Label(root, text="Resultats:")
result_label.pack(anchor="w", padx=20)

# Marc per contenir el Treeview i les barres de desplaçament
result_frame = Frame(root)
result_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))

# Crear Treeview per mostrar els resultats amb checkboxs
columns = ("Seleccionar", "Nom del fitxer", "Pàgina", "Paraula clau", "Fragment", "Ruta completa")
result_table = Treeview(result_frame, columns=columns, show="headings", selectmode="browse")

# Configurar les capçaleres de les columnes
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
result_table.heading("Ruta completa", text="Ruta completa")
result_table.column("Ruta completa", width=400, anchor="w")

# Crear scrollbar vertical
scrollbar_y = Scrollbar(result_frame, orient="vertical", command=result_table.yview)
scrollbar_y.pack(side="right", fill="y")

# Crear scrollbar horitzontal
scrollbar_x = Scrollbar(result_frame, orient="horizontal", command=result_table.xview)
scrollbar_x.pack(side="bottom", fill="x")

# Vincular les scrollbars al Treeview
result_table.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
result_table.pack(side="left", fill="both", expand=True)

# Afegir esdeveniment per fer clic al checkbox
result_table.bind("<Button-1>", toggle_checkbox)

# Botó de guardar
save_button = Button(root, text="Guardar CSV", command=guardar_csv)
save_button.pack(pady=10)

# Peu
footer_label = Label(root, text="Ajuntament de Mataró - SSIT - Secció Aplicacions", font=("Helvetica", 10))
footer_label.place(relx=0.5, rely=1.0, anchor="s")  # Centra l'etiqueta al peu de la finestra


root.mainloop()
