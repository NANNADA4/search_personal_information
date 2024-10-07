"""
입력받은 폴더를 순회하면서 pdf, hwp, xlsx 파일의 경로를 구하고 extract_information으로 넘깁니다
"""

import os
from natsort import natsorted


from module.processing_file import processing_hwp, processing_pdf, processing_excel
from module.save_excel import save_infos_to_excel


def processing_folder(folder_path, excel_file):
    """폴더 내부를 순회하며, pdf, hwp, xlsx 파일을 찾아 개인정보를 찾습니다."""
    for root, _, files in os.walk(folder_path):
        infos_list = []
        for filename in natsorted(files):
            infos_list = []
            if filename.lower().endswith('.pdf'):
                pdf_file_path = os.path.join('\\\\?\\', root, filename)
                print(f"진행중 : {pdf_file_path}")
                pdf_result = processing_pdf(folder_path, pdf_file_path)
                infos_list.extend(pdf_result)
            elif filename.lower().endswith('.hwp') or filename.lower().endswith('.hwpx'):
                hwp_file_path = os.path.join('\\\\?\\', root, filename)
                print(f"진행중 : {hwp_file_path}")
                hwp_result = processing_hwp(folder_path, hwp_file_path)
                infos_list.extend(hwp_result)
            elif filename.lower().endswith('.xlsx'):
                xlsx_file_path = os.path.join('\\\\?\\', root, filename)
                print(f"진행중 : {xlsx_file_path}")
                xlsx_result = processing_excel(folder_path, xlsx_file_path)
                infos_list.extend(xlsx_result)
        save_infos_to_excel(infos_list, excel_file)
