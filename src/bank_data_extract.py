#!/usr/bin/env python3

"""

FILENAME :      bank_data_extract.py             
AUTHOR :        Dennis Hartel        
START DATE :    18.01.2025
VERSION :       1.0
PYTHON :        3.9.2

REQUIREMENTS :  PyPDF2
DESCRIPTION :   Automatic bank sheet data extraction from PDF files
STATUS :        In development

"""

import PyPDF2
import re
import os
import shutil
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Constants
DATE_FORMAT = "%Y%m%d"
EXTRACTED_DATA_TEMPLATE = {
    "Konto Inhaber": "Nicht gefunden",
    "Depotnummer": "Nicht gefunden",
    "ISIN": "Nicht gefunden",
    "Datum": "Nicht gefunden",
    "Dokumenttyp": "Nicht gefunden"
}

# File Paths
current_folder = os.getcwd()
sup_folder = os.path.join(current_folder, "sup")

class PDFProcessor:
    @staticmethod
    def extract_info_from_pdf(pdf_path):
        try:
            with open(pdf_path, "rb") as file:
                reader = PyPDF2.PdfReader(file)
                text = "\n".join([page.extract_text() for page in reader.pages])
            extracted_data = EXTRACTED_DATA_TEMPLATE.copy()

            # Datum extrahieren und formatieren
            date_match = re.search(r"Frankfurt,\s(\d{2})\.(\d{2})\.(\d{4})", text)
            if date_match:
                extracted_data["Datum"] = f"{date_match.group(3)}{date_match.group(2)}{date_match.group(1)}"
            
            # ISIN extrahieren (Kauf/Verkauf oder allgemeine ISIN)
            isin_match = re.search(r"(Kauf|Verkauf)\s+.*?\((?P<isin>[A-Z0-9]{12})/[A-Z0-9]+\)", text)
            if not isin_match:
                isin_match = re.search(r"ISIN\s*:\s*([A-Z]{2}[A-Z0-9]{10})", text)
            
            if isin_match:
                extracted_data["ISIN"] = isin_match.group("isin") if "isin" in isin_match.groupdict() else isin_match.group(1)
           
            # Depotnummer extrahieren
            depot_match = re.search(r"Ihre Depotnummer:\s*(\d+)", text)
            if depot_match:
                extracted_data["Depotnummer"] = depot_match.group(1)
            
            # Depotinhaber extrahieren (Herr oder Frau)
            inhaber_match = re.search(r"(Herrn|Frau)\s*(.*?)(?:\n|$)", text, re.DOTALL)
            if inhaber_match:
                extracted_data["Konto Inhaber"] = inhaber_match.group(2).replace("\n", " ").strip()
            
            # Dokumenttyp extrahieren (Spezifische Begriffe)
            type_match = re.search(r"(Sammelabrechnung|Abrechnung|Wertpapierabrechnung\s+(?:Kauf|Verkauf|Vorabpauschale)|Storno|Dividendenabrechnung|Zinsabrechnung|Steuerbescheinigung|Kontoauszug|Rechnung|Gutschrift|Beleg|Transaktionsübersicht)", text)
            if type_match:
                extracted_data["Dokumenttyp"] = type_match.group(1)
            
            return extracted_data
        except Exception as e:
            logging.error(f"Fehler beim Verarbeiten von {pdf_path}: {e}")
            return None

    @staticmethod
    def process_all_pdfs(folder_path):
        pdf_files = [f for f in os.listdir(folder_path) if f.endswith(".pdf")]
        extracted_info_list = []
        
        if pdf_files:
            new_folder = os.path.join(folder_path, "NEW")
            os.makedirs(new_folder, exist_ok=True)
            output_txt_path = os.path.join(new_folder, "extracted_info.txt")
        
        for pdf_file in pdf_files:
            pdf_path = os.path.join(folder_path, pdf_file)
            extracted_info = PDFProcessor.extract_info_from_pdf(pdf_path)
            if extracted_info:
                extracted_info_list.append(extracted_info)
                
                # Neuen Dateinamen generieren
                new_filename = PDFProcessor.generate_new_filename(extracted_info)
                new_filepath = os.path.join(new_folder, new_filename)
                
                # Datei kopieren und umbenennen
                try:
                    shutil.copy(pdf_path, new_filepath)
                    logging.info(f"Kopiert: {pdf_file} -> {new_filename}")
                except Exception as e:
                    logging.error(f"Fehler beim Kopieren von {pdf_file}: {e}")
        
        PDFProcessor.save_extracted_info(output_txt_path, extracted_info_list)

    @staticmethod
    def generate_new_filename(extracted_info):
        max_length = 50
        truncated_inhaber = (extracted_info.get('Konto Inhaber') or 'Unbekannt')[:max_length]
        truncated_dokumenttyp = (extracted_info.get('Dokumenttyp') or 'Unbekannt')[:max_length]
        truncated_isin = (extracted_info.get('ISIN') or 'Unbekannt')[:max_length]
        
        new_filename = f"{extracted_info['Datum']}_{extracted_info['Depotnummer']}_{truncated_dokumenttyp}_{truncated_inhaber}_{truncated_isin}.pdf"
        return new_filename.replace(" ", "_")

    @staticmethod
    def save_extracted_info(output_txt_path, extracted_info_list):
        result_data = PDFProcessor.generate_result_data(extracted_info_list)
        try:
            PDFProcessor.write_to_file(output_txt_path, result_data)
        except Exception as e:
            logging.error(f"Fehler beim Speichern der Daten: {e}")
            
    @staticmethod
    def generate_result_data(extracted_info_list):
        result_data = f"Anzahl Dokumente: {len(extracted_info_list)}\n\n"
        for entry in extracted_info_list:
            result_data += "\n".join([f"{key}: {value}" for key, value in entry.items()]) + "\n\n" + ("-" * 50) + "\n"
        return result_data

    @staticmethod
    def write_to_file(output_txt_path, data):
        try:
            with open(output_txt_path, "w", encoding="utf-8") as out_file:
                out_file.write(data)
            logging.info(f"Daten erfolgreich in {output_txt_path} gespeichert.")
        except Exception as e:
            logging.error(f"Fehler beim Speichern der Daten: {e}")

if __name__ == "__main__":
    if os.path.exists(sup_folder):
        PDFProcessor.process_all_pdfs(sup_folder)
    else:
        logging.warning(f"Ordner {sup_folder} existiert nicht. Verwende aktuellen Pfad.")
        PDFProcessor.process_all_pdfs(current_folder)
