import os
import re
import zipfile
import docx2txt
from PyPDF2 import PdfReader
from openpyxl import Workbook

def datafromdocument(file_path):
    txt = docx2txt.process(file_path)
    emails = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', txt)
    phonenum = re.findall(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]', txt)
    return {'email': emails, 'phone': phonenum, 'text': txt}

def datafrompdf(file_path):
    txt = ""
    with open(file_path, 'rb') as fle:
        reading = PdfReader(fle)
        for page in reading.pages:
            txt += page.extract_text()
    emails = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', txt)
    phonenums = re.findall(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]', txt)
    return {'email': emails, 'phone': phonenums, 'text': txt}

def datafromzip(zip_file_path, outputdirectory):
    with zipfile.ZipFile(zip_file_path, 'r') as zipinput:
        zipinput.extractall(outputdirectory)
        subdirectories = [name for name in zipinput.namelist() if os.path.isdir(os.path.join(outputdirectory, name))]
        if subdirectories:
            return os.path.join(outputdirectory, subdirectories[0]) #assume in first subdirectory
        else:
            return outputdirectory

def handlingsubfolder(directory, worksheet):
    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)
        if os.path.isdir(file_path):
            # If it's a directory, recursively process its contents
            handlingsubfolder(file_path, worksheet)
        elif os.path.isfile(file_path):
            if filename.endswith('.docx'):
                info = datafromdocument(file_path)
            elif filename.endswith('.pdf'):
                info = datafrompdf(file_path)
            else:
                continue

            worksheet.append([filename, ', '.join(info['email']), ', '.join(info['phone']), info['text']])

def main():
    # Create a workbook
    wb = Workbook()
    ws = wb.active
    ws.append(['File Name', 'Email', 'Phone', 'Text'])

    # Directory containing the extracted CVs
    extracteddirectory = 'extracted_cvsff'  # Replace with your desired directory name

    # Extract CVs from the zip file
    zip_file_path = 'Sample2-20240406T093029Z-001.zip'  # Replace with your zip file path
    cv_directory = datafromzip(zip_file_path, extracteddirectory)

    handlingsubfolder(cv_directory, ws)

    # Save the workbook
    wb.save('cv_infoff.xlsx')

if __name__ == "__main__":
    main()
