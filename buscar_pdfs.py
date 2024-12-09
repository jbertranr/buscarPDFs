import pdfplumber
import os
import csv

# Ruta de la carpeta amb els PDFs
folder_path = "C:\\Users\\jbertran.MATARO\\AppData\\Local\\buscarPDFs\\pdfs"
# Llista de paraules clau per buscar
search_terms = ["permetre", "solució", "integració"]  # Afegeix aquí les paraules clau
# Nom del fitxer CSV de sortida
output_csv = "resultats_busqueda.csv"

# Crea una llista per guardar els resultats
results = []

# Processa tots els fitxers PDF a la carpeta
for file in os.listdir(folder_path):
    if file.endswith(".pdf"):
        pdf_path = os.path.join(folder_path, file)
        with pdfplumber.open(pdf_path) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                text = page.extract_text().replace("\n", " ") if page.extract_text() else ""
                if text:
                    for term in search_terms:
                        if term in text:
                            # Guarda el resultat
                            results.append({
                                "Nom del fitxer": file,
                                "Pàgina": page_number,
                                "Paraula clau": term,
                                "Fragment": text[:200]  # Mostra un fragment del text trobat
                            })

# Escriu els resultats en un fitxer CSV
with open(output_csv, mode='w', newline='', encoding='utf-8') as csv_file:
    fieldnames = ["Nom del fitxer", "Pàgina", "Paraula clau", "Fragment"]
    writer = csv.DictWriter(csv_file, fieldnames=fieldnames, quoting=csv.QUOTE_ALL, quotechar='"')
    
    # Escriu les capçaleres i les dades
    writer.writeheader()
    writer.writerows(results)

print(f"Resultats guardats al fitxer: {output_csv}")
