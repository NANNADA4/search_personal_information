"""
동일한 정규표현식을 사용하여 각 파일에 맞는 방법으로 개인정보를 추출합니다 
"""

import os
import re
import win32com.client as win32
import fitz
from openpyxl import load_workbook


PATTERN_EMAILS = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
PATTERN_JUMINS = r'\d{2}[01]\d[0123]\d- [1-4]\d{6}'
PATTERN_CREDIT_NUMS = r'\b\d{4}-\d{4}-\d{4}-\d{4}\b'
PATTERN_CELLPHONE_NUMS = r'\b(010-\d{4}-\d{4}|01[16789]-\d{3,4}-\d{4})\b'
PATTERN_DRIVER_NUMS = r'\d{2}-\d{2}-\d{6}-\d{2}'


PATTERNS = {
    '이메일': PATTERN_EMAILS,
    '주민등록번호': PATTERN_JUMINS,
    '신용카드번호': PATTERN_CREDIT_NUMS,
    '휴대전화번호': PATTERN_CELLPHONE_NUMS,
    '운전면허번호': PATTERN_DRIVER_NUMS
}


def _extract_personal_information(file, text, page_num=None, sheet_name=None):
    """정규표현식으로 개인정보를 추출하여 리스트로 return합니다"""
    infos = []

    for info_type, pattern in PATTERNS.items():
        matches = re.findall(pattern, text)
        for match in matches:
            if page_num is not None:
                infos.append(
                    (os.path.basename(file), page_num + 1, info_type, match))
            elif sheet_name is not None:
                infos.append(
                    (os.path.basename(file), sheet_name, info_type, match))
            else:
                infos.append((os.path.basename(file), None, info_type, match))

    return infos


def processing_pdf(pdf_file):
    """pdf파일을 처리후, pdf_infos에 모든 결과를 리스트로 저장하여 return합니다"""
    doc = fitz.open(pdf_file)
    pdf_infos = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text()
        pdf_infos.extend(_extract_personal_information(
            pdf_file, text, page_num))

    return pdf_infos


def processing_hwp(hwp_file):
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
            if state in [0, 1]:
                break
            hwp_infos.extend(
                _extract_personal_information(hwp_file, text))

    except Exception as e:  # pylint: disable=W0703
        print(f"hwp 오류 발생 : {e}")

    finally:
        if hwp:
            hwp.ReleaseScan()
            hwp.Quit()

    return hwp_infos


def processing_excel(excel_file):
    """엑셀 파일을 처리 후, excel_infos에 모든 결과를 리스트로 저장하여 반환합니다"""
    excel_infos = []
    wb = load_workbook(excel_file)

    for sheet in wb.sheetnames:
        ws = wb[sheet]
        text = ""
        for row in ws.iter_rows(values_only=True):
            for cell in row:
                if cell:
                    text += str(cell) + " "
        excel_infos.extend(_extract_personal_information(
            excel_file, text, sheet_name=sheet))

    return excel_infos
