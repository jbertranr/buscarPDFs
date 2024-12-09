import pdfplumber
import os
import csv
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Frame, Button, Label, Entry, Scrollbar, Style

def buscar_pdfs():
    folder_path = folder_path_entry.get()
    search_terms = search_terms_entry.get().strip().split(",")
    output_csv = os.path.join(folder_path, "resultats_busqueda.csv")
    
    if not folder_path or not os.path.isdir(folder_path):
        messagebox.showerror("Error", "La ruta de la carpeta no és vàlida.")
        return
    if not search_terms:
        messagebox.showerror("Error", "Introdueix almenys una paraula clau.")
        return

    results = []

    pdf_files = [file for file in os.listdir(folder_path) if file.endswith(".pdf")]
    total_files = len(pdf_files)

    for i, file in enumerate(pdf_files, start=1):
        # Actualitza l'etiqueta per mostrar el fitxer processat
        processing_label.config(text="")
        processing_label.update()  # Força l'actualització de la interfície gràfica
        processing_label.config(text=f"Processant: {file} ({i}/{total_files})")
        processing_label.update()  # Força l'actualització de la interfície gràfica

        pdf_path = os.path.join(folder_path, file)
        with pdfplumber.open(pdf_path) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                text = page.extract_text().replace("\n", " ") if page.extract_text() else ""
                if text:
                    for term in search_terms:
                        if term in text:
                            # Trobar l'índex de la paraula clau
                            start_idx = max(0, text.find(term) - 50)  # 50 caràcters abans
                            end_idx = min(len(text), text.find(term) + 50 + len(term))  # 50 caràcters després
                            fragment = text[start_idx:end_idx]

                            results.append({
                                "Nom del fitxer": file,
                                "Pàgina": page_number,
                                "Paraula clau": term,
                                "Fragment": fragment
                            })

    # Missatge final un cop processats tots els fitxers
    processing_label.config(text="Processament complet!")
    result_box.delete("1.0", "end")
    result_box.tag_configure("highlight", background="yellow", foreground="black")

    if results:
        for result in results:
            result_box.insert("end", f"Fitxer: {result['Nom del fitxer']}\n")
            result_box.insert("end", f"Pàgina: {result['Pàgina']}\n")
            result_box.insert("end", f"Paraula clau trobada: {result['Paraula clau']}\n")
            result_box.insert("end", "Fragment: ")

            # Inserir el fragment amb paraules ressaltades
            fragment = result['Fragment']
            for term in search_terms:
                while term in fragment:
                    start = fragment.find(term)
                    end = start + len(term)
                    result_box.insert("end", fragment[:start])
                    result_box.insert("end", fragment[start:end], "highlight")
                    fragment = fragment[end:]
            result_box.insert("end", fragment + "\n")
            result_box.insert("end", "-" * 50 + "\n")
    else:
        result_box.insert("end", "No s'han trobat coincidències.\n")

    save_button.config(state="normal", command=lambda: guardar_csv(results, output_csv))

def guardar_csv(results, output_csv):
    try:
        with open(output_csv, mode='w', newline='', encoding='utf-8') as csv_file:
            fieldnames = ["Nom del fitxer", "Pàgina", "Paraula clau", "Fragment"]
            writer = csv.DictWriter(csv_file, fieldnames=fieldnames, quoting=csv.QUOTE_ALL, quotechar='"')
            writer.writeheader()
            writer.writerows(results)
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
root.geometry("800x800")

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

result_box = tk.Text(result_frame, wrap="word", height=10, font=("Helvetica", 12))
result_box.pack(side="left", fill="both", expand=True)

scrollbar = Scrollbar(result_frame, orient="vertical", command=result_box.yview)
scrollbar.pack(side="right", fill="y")

result_box.config(yscrollcommand=scrollbar.set)

# Botó de guardar
save_button = Button(root, text="Guardar CSV", state="disabled")
save_button.pack(pady=10)

# Peu
footer_label = Label(root, text="Ajuntament de Mataró - SSIT - Secció Aplicacions", font=("Helvetica", 10))
footer_label.pack(pady=(10, 0))

root.mainloop()
