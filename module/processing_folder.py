"""
입력받은 폴더를 순회하면서 pdf, hwp, xlsx 파일의 경로를 구하고 extract_information으로 넘깁니다
"""

import os


from module.save_excel import save_infos_to_excel
from module.extract_information import processing_pdf, processing_hwp, processing_xlsx


def processing_folder(folder_path, excel_file):
    """폴더 내부를 순회하며, pdf, hwp, xlsx 파일을 찾아 개인정보를 찾습니다."""
    infos_list = []

    for root, _, files in os.walk(folder_path):
        for filename in files:
            if filename.lower().endswith('.pdf'):
                pdf_file_path = os.path.join('\\\\?\\', root, filename)
                pdf_result = processing_pdf(pdf_file_path)
                infos_list.extend(pdf_result)
            elif filename.lower().endswith('.hwp'):
                hwp_file_path = os.path.join('\\\\?\\', root, filename)
                hwp_result = processing_hwp(hwp_file_path)
                infos_list.extend(hwp_result)
            elif filename.lower().endswith('.xlsx'):
                xlsx_file_path = os.path.join('\\\\?\\', root, filename)
                xlsx_result = processing_xlsx(xlsx_file_path)
                infos_list.extend(xlsx_result)

    save_infos_to_excel(infos_list, excel_file)
