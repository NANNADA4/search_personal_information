import fitz


import os
import re
import win32com.client as win32


PATTERN_EMAILS = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
PATTERN_JUMINS = r'\d{2}[01]\d[0123]\d- [1-4]\d{6}'
PATTERN_CREDIT_NUMS = r'\b\d{4}-\d{4}-\d{4}-\d{4}\b'
PATTERN_CELLPHONE_NUMS = r'\b(010-\d{4}-\d{4}|01[16789]-\d{3,4}-\d{4})\b'
PATTERN_DRIVER_NUMS = r'\d{2}-\d{2}-\d{6}-\d{2}'


def _extract_personal_information(file, text, page_num):
    """정규표현식으로 개인정보를 추출하여 리스트로 return합니다"""
    infos = []

    pattern_email = re.findall(PATTERN_EMAILS, text)
    pattern_jumin = re.findall(PATTERN_JUMINS, text)
    pattern_credit_num = re.findall(PATTERN_CREDIT_NUMS, text)
    pattern_cellphone_num = re.findall(PATTERN_CELLPHONE_NUMS, text)
    pattern_driver = re.findall(PATTERN_DRIVER_NUMS, text)

    for email in pattern_email:
        infos.append((os.path.basename(file),
                      page_num + 1, '이메일', email))
    for jumin in pattern_jumin:
        infos.append((os.path.basename(file),
                      page_num + 1, '주민등록번호', jumin))
    for credit in pattern_credit_num:
        infos.append((os.path.basename(file),
                      page_num + 1, '신용카드번호', credit))
    for cellphone in pattern_cellphone_num:
        infos.append((os.path.basename(file),
                      page_num + 1, '휴대전화번호', cellphone))
    for driver in pattern_driver:
        infos.append((os.path.basename(file),
                     page_num+1, '운전면허번호', driver))

    return infos


def extract_infos_from_pdf(pdf_file):
    """pdf파일을 처리후, pdf_infos에 모든 결과를 리스트로 저장하여 return합니다"""
    doc = fitz.open(pdf_file)
    pdf_infos = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text()
        pdf_infos.append(_extract_personal_information(
            pdf_file, text, page_num))

    return pdf_infos


def _extract_personal_information_hwp(hwp_file, text, hwp):
    """HWP 문서에서 개인 정보를 추출하여 리스트로 반환합니다"""
    infos = []

    patterns = {
        '이메일': PATTERN_EMAILS,
        '주민등록번호': PATTERN_JUMINS,
        '신용카드번호': PATTERN_CREDIT_NUMS,
        '휴대전화번호': PATTERN_CELLPHONE_NUMS,
        '운전면허번호': PATTERN_DRIVER_NUMS
    }

    for info_type, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            infos.append((os.path.basename(hwp_file), hwp.KeyIndicator()[
                         3], info_type, match.group()))

    return infos


def extract_infos_from_hwp(hwp_file):
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
                _extract_personal_information_hwp(hwp_file, text, hwp))

    except Exception as e:
        print(f"hwp 오류 발생 : {e}")

    finally:
        if hwp:
            hwp.ReleaseScan()
            hwp.Quit()

    return hwp_infos
