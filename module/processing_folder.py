from module.save_excel import save_infos_to_excel
from module.extract_information import extract_infos_from_pdf


import os


def processing_folder(folder_path, excel_file):
    infos_list = []

    for root, _, files in os.walk(folder_path):
        for filename in files:
            if filename.lower().endswith('.pdf'):
                pdf_file_path = os.path.join(root, filename)
                infos = extract_infos_from_pdf(pdf_file_path)
                infos_list.extend(infos)

    save_infos_to_excel(infos_list, excel_file)
