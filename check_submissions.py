import os
import pypdf
import pandas as pd

# List of expected professors
expected_names_text = """
Aleksova Aneta, Barbi Egidio, Bernardi Stella, Boscolo Rizzo Paolo, Buoite Stella Alex, Cadenaro Milena, Cattalini Marco, Comar Manola, Contardo Luca, D'Errico Stefano, Fornasiero Eugenio, Franca Raffaella, Gasparini Paolo, Girotto Giorgia, Giuffrè Mauro, Lucafò Marianna, Marchesi Giulio, Mardirossian Mario, Merlo Marco, Murena Luigi, Ottaviani Giulia, Palmisano Silvia, Ricci Giuseppe, Romano Maurizio, Ruaro Barbara, Sinagra Gianfranco, Sorrentino Giovanni, Stampalija Tamara, Stocco Gabriele, Taddio Andrea, Tommasini Alberto, Turco Gianluca, Zacchigna Serena, Fabris Enrico, Dal Ferro Matteo, Salton Francesco, Zerbato Verena, Ratti Chiara, Mazzà Daniela, Travan Laura, Romano Federico, Maestro Roberta, Cecchin Erika, Spessotto Paola
"""

# Clean up and create a list of surnames for matching
# Assuming "Surname Name" format in the text, but we'll match primarily on Surname as it's more unique usually
# actually some are "Surname Name" and some might be "Name Surname".
# Let's store full names and try to match parts.

def parse_names(text):
    # Remove newlines and split by comma
    raw_names = text.replace("\n", " ").split(",")
    names = []
    for n in raw_names:
        n = n.strip()
        if n:
            # Split into parts
            parts = n.split()
            # Heuristic: usually Surname Name in Italian lists, but let's keep the whole string for display
            # and use parts for matching.
            names.append({"full": n, "parts": [p.lower() for p in parts]})
    return names

expected_people = parse_names(expected_names_text)

target_dir = "/Users/piuzzo/Library/CloudStorage/OneDrive-Personale/UNITS/Dottorato/ciclo_42/Excel_docenti"

found_people = []
files_processed = []

print(f"Checking files in: {target_dir}")

if not os.path.exists(target_dir):
    print("Directory not found!")
    exit(1)

for filename in os.listdir(target_dir):
    # Check all files
    file_path = os.path.join(target_dir, filename)
    files_processed.append(filename)
    
    matched_person = None
    
    # 1. Try text extraction if PDF
    if filename.lower().endswith(".pdf"):
        try:
            reader = pypdf.PdfReader(file_path)
            text = ""
            for page in reader.pages:
                text += page.extract_text() + "\n"
            
            text_lower = text.lower()
            
            for person in expected_people:
                parts_matched = 0
                for part in person['parts']:
                    if part in text_lower:
                        parts_matched += 1
                
                if parts_matched == len(person['parts']):
                    matched_person = person
                    break
            
            if matched_person:
                 print(f"[FOUND] {filename} -> {matched_person['full']} (matched in text)")
        except Exception as e:
            print(f"[ERROR] Could not read PDF {filename}: {e}")

    # 2. Try text extraction if Excel
    elif filename.lower().endswith((".xls", ".xlsx")):
        try:
            # Read excel, convert to string to search
            # header=None to read all cells including headers
            df = pd.read_excel(file_path, header=None)
            text = df.to_string()
            text_lower = text.lower()
            
            for person in expected_people:
                parts_matched = 0
                for part in person['parts']:
                    if part in text_lower:
                        parts_matched += 1
                
                if parts_matched == len(person['parts']):
                    matched_person = person
                    break
            
            if matched_person:
                 print(f"[FOUND] {filename} -> {matched_person['full']} (matched in excel content)")
        except Exception as e:
             # Some files might be corrupted or password protected
             print(f"[ERROR] Could not read Excel {filename}: {e}")

    # 2. Fallback: check filename (if not already matched)
    if not matched_person:
        filename_lower = filename.lower()
        best_match = None
        max_parts_matched = 0
        
        for person in expected_people:
            parts_matched = 0
            for part in person['parts']:
                if part in filename_lower:
                    parts_matched += 1
            
            # We require at least the surname (first part) to be present usually, 
            # but let's just go with max parts for now.
            # Actually, let's require at least 1 part.
            
            if parts_matched > 0:
                if parts_matched > max_parts_matched:
                    max_parts_matched = parts_matched
                    best_match = person
                elif parts_matched == max_parts_matched:
                    # Tie breaking?
                    # If we have "Romano" and matches both, we might have an issue if filename is just "Romano.pdf".
                    # But if filename is "Federico Romano", Federico gets 2, Maurizio gets 1.
                    pass
        
        if best_match:
             # Optional: set a threshold? e.g. if only 1 part matched and it's a common surname...
             # But for now, best match is better than first match.
             matched_person = best_match
             print(f"[FOUND] {filename} -> {matched_person['full']} (matched in filename with {max_parts_matched} parts)")
        else:
            print(f"[UNKNOWN] {filename} -> Could not identify person from text or filename")
    
    if matched_person:
        if matched_person['full'] not in found_people:
            found_people.append(matched_person['full'])


print("\n" + "="*30)
print("MISSING SUBMISSIONS:")
print("="*30)

missing_count = 0
for person in expected_people:
    if person['full'] not in found_people:
        print(f"- {person['full']}")
        missing_count += 1

print(f"\nTotal expected: {len(expected_people)}")
print(f"Total found: {len(found_people)}")
print(f"Total missing: {missing_count}")
