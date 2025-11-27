import pypdf
import pandas as pd
import urllib.parse
import os

def process_pdf_to_whatsapp_excel(pdf_filename,excel_filename, phone_number):
    """
     Reads text from a pdf , generates a whatsapp link, and  saves  to an excel file.

     :param pdf_filename: The name of the input pdf file.
     :param excel_filename: The name of the output excel file.
     :param phone_number: The recipient's phone number (with country code, no '+' or spaces).
    """
    try:
        reader = pypdf.PdfReader(pdf_filename)
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
        print(f"Sucessfully extracted text from  {pdf_filename}")

    except Exception as e:
        print(f"Error reading PDF file: {e}")
        return  

    encoded_text = urllib.parse.quote_plus(text)
    whatsapp_url = f"https://wa.me/{phone_number}?text={encoded_text}"
    print(f"Generated whatsapp URL: {whatsapp_url}")

    data = {
        'pdf Filename': [pdf_filename],
        'Excel Filename': [text],
        'whatsapp link' : [whatsapp_url]
    }      

    df = pd.DataFrame(data)

    try:
        df.to_excel(excel_filename, index=False, engine='openpyxl')
        print(f"Successfully saved data to {excel_filename}")

    except Exception as e:
        print(f"Error saving to Excel file: {e}")

if __name__ == "__main__":
    PDF_FILE = 'Election.pdf'
    EXCEL_FILE = 'output_links.xlsx'
    PHONE = '8428757942'

    if os.path.exists(PDF_FILE):
        process_pdf_to_whatsapp_excel(PDF_FILE, EXCEL_FILE, PHONE)

    else:
        print(f"Error: {PDF_FILE} not found. please create the file or update the script.")
                