"""
동일한 정규표현식을 사용하여 각 파일에 맞는 방법으로 개인정보를 추출합니다
"""

import os
import re
import pathlib
import phonenumbers
from phonenumbers import NumberParseException


from module.data import PATTERNS


def _find_name(folder_path, file):
    """위원회, 피감기관을 파일명에서 검색 후 추출합니다"""
    if os.path.basename(folder_path).find(' ') != -1:
        if os.path.basename(folder_path).find('_') != -1:
            cmt = os.path.basename(folder_path)[
                os.path.basename(folder_path).find(' ') + 1:os.path.basename(folder_path).find('_')]
        else:
            cmt = os.path.basename(folder_path)[
                os.path.basename(folder_path).find(' ') + 1:]
    else:
        cmt = os.path.basename(folder_path)

    org = os.path.relpath(file, os.path.dirname(folder_path)).split(os.sep)[1]

    return cmt, org


def _extract_info_patterns(file, text, name, page_num, infos):
    """정규표현식 패턴을 사용하여 개인정보 추출을 시도합니다"""
    cmt, org = name
    for info_type, pattern in PATTERNS.items():
        if info_type == '계좌번호':
            for account_pattern in pattern:
                matches = re.findall(account_pattern, text)
                for match in matches:
                    infos.append((
                        cmt, org,
                        os.path.basename(file), pathlib.Path(
                            file).suffix.lstrip('.').lower(),
                        page_num if isinstance(page_num, str) else (
                            page_num + 1 if page_num is not None else None),
                        info_type, match, None
                    ))
        else:
            matches = re.findall(pattern, text)
            for match in matches:
                infos.append((
                    cmt, org,
                    os.path.basename(file), pathlib.Path(
                        file).suffix.lstrip('.').lower(),
                    page_num if isinstance(page_num, str) else (
                        page_num + 1 if page_num is not None else None),
                    info_type, match, None
                ))

    return infos


def _extract_info_phonenum(file, text, name, page_num, infos):
    """phonenumbers 라이브러리로 국제전화번호를 추출합니다"""
    cmt, org = name
    for word in text.split():
        try:
            number = phonenumbers.parse(word, None)
            if phonenumbers.is_valid_number(number):
                phone_number = phonenumbers.format_number(
                    number, phonenumbers.PhoneNumberFormat.INTERNATIONAL)
                infos.append((
                    cmt, org,
                    os.path.basename(file), pathlib.Path(
                        file).suffix.lstrip('.').lower(),
                    page_num + 1 if page_num is not None else None,
                    '해외전화번호', phone_number, None
                ))
        except NumberParseException:
            continue

    return infos


def extract_personal_information(folder_path, file, text=None, page_num=None, error=None):
    """정규표현식으로 개인정보를 추출하여 리스트로 return합니다"""
    infos = []

    cmt, org = _find_name(folder_path, file)

    if text is None:
        infos.append((
            cmt, org,
            os.path.basename(file), pathlib.Path(
                file).suffix.lstrip('.').lower(),
            None, None, None, error
        ))
        return infos

    _extract_info_patterns(file, text, _find_name(
        folder_path, file), page_num, infos)

    _extract_info_phonenum(file, text, _find_name(
        folder_path, file), page_num, infos)

    return infos
