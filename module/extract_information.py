"""
동일한 정규표현식을 사용하여 각 파일에 맞는 방법으로 개인정보를 추출합니다
"""

import os
import re
import pathlib
import pandas as pd
import win32com.client as win32
import fitz
import phonenumbers
from phonenumbers import NumberParseException


from module.patterns import PATTERNS


def _extract_personal_information(folder_path, file, text=None, page_num=None, error=None):
    """정규표현식으로 개인정보를 추출하여 리스트로 return합니다"""
    infos = []

    if os.path.basename(folder_path).find(' ') != -1:
        if os.path.basename(folder_path).find('_') != -1:
            cmt = os.path.basename(folder_path)[
                os.path.basename(folder_path).find(' ') + 1:os.path.basename(folder_path).find('_')]
        else:
            cmt = os.path.basename(folder_path)[
                os.path.basename(folder_path).find(' ') + 1:]
    else:
        cmt = os.path.basename(folder_path)

    relative_path = os.path.relpath(file, os.path.dirname(folder_path))

    if text is None:
        infos.append((
            cmt, relative_path.split(os.sep)[1],
            os.path.basename(file), pathlib.Path(
                file).suffix.lstrip('.').lower(),
            None, None, None, error
        ))
        return infos

    for info_type, pattern in PATTERNS.items():
        matches = re.findall(pattern, text)
        for match in matches:
            infos.append((
                cmt, relative_path.split(os.sep)[1],
                os.path.basename(file), pathlib.Path(
                    file).suffix.lstrip('.').lower(),
                page_num if isinstance(page_num, str) else (
                    page_num + 1 if page_num is not None else None),
                info_type, match, None
            ))

    for word in text.split():
        try:
            number = phonenumbers.parse(word, None)
            if phonenumbers.is_valid_number(number):
                phone_number = phonenumbers.format_number(
                    number, phonenumbers.PhoneNumberFormat.INTERNATIONAL)
                infos.append((
                    cmt, relative_path.split(os.sep)[1],
                    os.path.basename(file), pathlib.Path(
                        file).suffix.lstrip('.').lower(),
                    page_num + 1 if page_num is not None else None,
                    '해외전화번호', phone_number, None
                ))
        except NumberParseException:
            continue

    return infos


def processing_pdf(folder_path, pdf_file):
    """pdf파일을 처리후, pdf_infos에 모든 결과를 리스트로 저장하여 return합니다"""
    pdf_infos = []
    try:
        doc = fitz.open(pdf_file)
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text = page.get_text()
            pdf_infos.extend(_extract_personal_information(folder_path,
                                                           pdf_file, text=text, page_num=page_num))
    except Exception as e:  # pylint: disable=W0703
        error_log = str(e)
        pdf_infos.extend(
            _extract_personal_information(folder_path, pdf_file, error=error_log))
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
                _extract_personal_information(folder_path, hwp_file, text=text,
                                              page_num=hwp.KeyIndicator()[3]))

    except Exception as e:  # pylint: disable=W0703
        error_log = str(e)
        hwp_infos.extend(
            _extract_personal_information(folder_path, hwp_file, error=error_log))
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
                        _extract_personal_information(folder_path, excel_file,
                                                      text=text, page_num=xlsx_index))

    except Exception as e:  # pylint: disable=W0703
        error_log = str(e)
        excel_infos.extend(
            _extract_personal_information(folder_path, excel_file, error=error_log))
        print(excel_file, e)

    return excel_infos
