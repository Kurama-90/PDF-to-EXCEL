#Kuram-90 ( https://github.com/Kurama-90 )

import fitz  # PyMuPDF
import pandas as pd  # Pour créer un fichier Excel
import os

# Chemins des dossiers
UPLOADS_DIR = 'uploads'
OUTPUTS_DIR = 'outputs'

def extract_text_and_tables_from_pdf(pdf_path):
    # Ouvrir le PDF
    pdf_document = fitz.open(pdf_path)
    
    # Variables pour stocker les données
    combined_data = []  # Pour combiner texte et tableaux
    tables = []  # Pour les tables

    # Parcourir chaque page du PDF
    for page_num in range(len(pdf_document)):
        print(f"Traitement de la page {page_num + 1}...")
        page = pdf_document[page_num]
        
        # Extraire les tables de la page avec leurs coordonnées
        tables_in_page = page.find_tables()  # Détecter les tables dans la page
        table_rects = []  # Pour stocker les zones des tableaux

        for table in tables_in_page:
            # Extraire les données de la table
            table_data = table.extract()
            tables.append(table_data)
            
            # Stocker les coordonnées de la table (x0, y0, x1, y1)
            table_rect = fitz.Rect(table.bbox)  # Convertir le tuple en objet Rect
            table_rects.append(table_rect)
        
        # Extraire le texte brut de la page
        text_blocks = page.get_text("blocks")  # Obtenir les blocs de texte avec leurs coordonnées
        
        # Trier les blocs de texte et les tableaux par ordre d'apparition (y0)
        all_blocks = []
        
        # Ajouter les blocs de texte
        for block in text_blocks:
            x0, y0, x1, y1, block_text, block_no, block_type = block
            block_rect = fitz.Rect(x0, y0, x1, y1)  # Zone du bloc de texte
            
            # Vérifier si le bloc de texte est dans une zone de tableau
            is_in_table = False
            for table_rect in table_rects:
                if block_rect.intersects(table_rect):  # Si le bloc est dans une zone de tableau
                    is_in_table = True
                    break
            
            if not is_in_table:
                all_blocks.append({"type": "text", "y0": y0, "content": block_text.strip()})
        
        # Ajouter les tableaux
        for i, table_rect in enumerate(table_rects):
            all_blocks.append({"type": "table", "y0": table_rect.y0, "content": tables[i]})
        
        # Trier tous les blocs par ordre d'apparition (y0)
        all_blocks.sort(key=lambda x: x["y0"])
        
        # Ajouter les blocs triés à la liste combinée
        for block in all_blocks:
            combined_data.append(block)
    
    # Fermer le PDF
    pdf_document.close()
    
    return combined_data

def export_to_excel(combined_data, output_excel_path):
    # Créer un fichier Excel avec une seule feuille
    with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
        # Créer un DataFrame pour la feuille unique
        df_data = []
        
        for block in combined_data:
            if block["type"] == "text":
                # Ajouter le texte hors tableau
                df_data.append([block["content"]])
            elif block["type"] == "table":
                # Ajouter les lignes du tableau
                table_data = block["content"]
                for row in table_data:
                    df_data.append(row)
        
        # Convertir en DataFrame
        df = pd.DataFrame(df_data)
        
        # Exporter dans une seule feuille
        df.to_excel(writer, sheet_name="Feuille unique", index=False, header=False)
    
    print(f"Fichier Excel généré : {output_excel_path}")

def process_pdf(pdf_path, output_excel_path):
    # Extraire le texte hors tableau et les tables du PDF
    combined_data = extract_text_and_tables_from_pdf(pdf_path)
    
    # Exporter les données dans un fichier Excel
    export_to_excel(combined_data, output_excel_path)

if __name__ == '__main__':
    # Demander le nom du fichier PDF
    pdf_name = input("Veuillez entrer le nom du fichier PDF (avec l'extension .pdf) : ")
    
    # Chemin du fichier PDF dans le dossier uploads
    pdf_file = os.path.join(UPLOADS_DIR, pdf_name)
    
    # Vérifier si le fichier existe
    if not os.path.exists(pdf_file):
        print(f"Erreur : Le fichier {pdf_file} n'existe pas dans le dossier 'uploads'.")
        exit()
    
    # Chemin du fichier Excel de sortie dans le dossier outputs
    output_excel = os.path.join(OUTPUTS_DIR, 'output.xlsx')
    
    # Vérifier si le dossier outputs existe, sinon le créer
    if not os.path.exists(OUTPUTS_DIR):
        print(f"Création du dossier {OUTPUTS_DIR}...")
        os.makedirs(OUTPUTS_DIR)
    
    # Traiter le PDF et générer le fichier Excel
    process_pdf(pdf_file, output_excel)
    print(f"Le fichier Excel a été généré avec succès.")