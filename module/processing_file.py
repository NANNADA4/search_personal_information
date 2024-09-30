"""파일별 구분하여 스크립트를 진행합니다"""


import pandas as pd
import fitz


import win32com.client as win32
from module.extract_information import extract_personal_information


def processing_pdf(folder_path, pdf_file):
    """pdf파일을 처리후, pdf_infos에 모든 결과를 리스트로 저장하여 return합니다"""
    pdf_infos = []
    try:
        doc = fitz.open(pdf_file)
        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)
            text = page.get_text()
            pdf_infos.extend(extract_personal_information(folder_path,
                                                          pdf_file, text=text, page_num=page_num))
    except Exception as e:  # pylint: disable=W0703
        error_log = str(e)
        pdf_infos.extend(
            extract_personal_information(folder_path, pdf_file, error=error_log))
        print(pdf_file, e)

    return pdf_infos


def processing_hwp(folder_path, hwp_file):
    """hwp 파일을 처리 후, hwp_infos에 모든 결과를 리스트로 저장하여 반환합니다"""
    hwp_infos = []
    hwp = None

    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        hwp.Open(hwp_file)
        hwp.InitScan()

        while True:
            state, text = hwp.GetText()
            hwp.MovePos(201)
            if state in [0, 1]:
                break
            hwp_infos.extend(
                extract_personal_information(folder_path, hwp_file, text=text,
                                             page_num=hwp.KeyIndicator()[3]))

    except Exception as e:  # pylint: disable=W0703
        error_log = str(e)
        hwp_infos.extend(
            extract_personal_information(folder_path, hwp_file, error=error_log))
        print(hwp_file, e)

    finally:
        if hwp:
            hwp.ReleaseScan()
            hwp.Quit()

    return hwp_infos


def processing_excel(folder_path, excel_file):
    """엑셀 파일을 처리 후, excel_infos에 모든 결과를 리스트로 저장하여 반환합니다"""
    excel_infos = []

    try:
        xls = pd.ExcelFile(excel_file)

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            for row_index, row in df.iterrows():
                if row.isnull().all():
                    continue
                for col_index, cell in enumerate(row):
                    if pd.isna(cell):
                        continue
                    xlsx_index = f"[{row_index + 1}, {col_index + 1}]"
                    text = str(cell).strip()
                    excel_infos.extend(
                        extract_personal_information(folder_path, excel_file,
                                                     text=text, page_num=xlsx_index))

    except Exception as e:  # pylint: disable=W0703
        error_log = str(e)
        excel_infos.extend(
            extract_personal_information(folder_path, excel_file, error=error_log))
        print(excel_file, e)

    return excel_infos
