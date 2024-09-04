from module.save_excel import save_infos_to_excel
from module.extract_information import extract_infos_from_pdf, extract_infos_from_hwp


import os


def processing_folder(folder_path, excel_file):
    """폴더 내부를 순회하며, pdf, hwp, xlsx 파일을 찾아 개인정보를 찾습니다."""
    infos_list = []

    for root, _, files in os.walk(folder_path):
        for filename in files:
            if filename.lower().endswith('.pdf'):
                pdf_file_path = os.path.join('\\\\?\\', root, filename)
                pdf_result = extract_infos_from_pdf(pdf_file_path)
                infos_list.extend(pdf_result)
            elif filename.lower().endswith('.hwp'):
                hwp_file_path = os.path.join('\\\\?\\', root, filename)
                hwp_result = extract_infos_from_hwp(hwp_file_path)
                infos_list.extend(hwp_result)

    save_infos_to_excel(infos_list, excel_file)
