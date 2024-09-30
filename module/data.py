"""
정규표현식을 모아놓은 파일입니다
"""


PATTERN_GOOKMIN_ACCOUNTS = r'\b\d{6}-\d{2}-\d{6}\b'
PATTERN_SHINHAN_ACCOUNTS = r'\b\d{3}-\d{3}-\d{6}\b'
PATTERN_HANA_ACCOUNTS = r'\b\d{3}-\d{6}-\d{5}\b'
PATTERN_WOORI_ACCOUNTS = r'\b\d{3}-\d{6}-\d{6}\b'
PATTERN_KAKAO_ACCOUNTS = r'\b3333-\d{2,3}-\d{6}\b'
PATTERN_NH_ACCOUNTS = r'\b\d{3}-\d{4}-\d{4}-\d{2}\b'
PATTERN_EMAILS = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
PATTERN_JUMINS = r'\d{2}[01]\d[0123]\d- [1-4]\d{6}'
PATTERN_CREDIT_NUMS = r'\b\d{4}-\d{4}-\d{4}-\d{4}\b'
PATTERN_CELLPHONE_NUMS = r'\b(010-\d{4}-\d{4}|01[16789]-\d{3,4}-\d{4})\b'
PATTERN_PHONE_NUMS = r'\b(02-\d{4}-\d{4}|0[3-9]\d-\d{3,4}-\d{4})\b'
PATTERN_REPRESENTATIVE_PHONE_NUMBERS = r'\b(?:1566|1600|1670|1533|1551|1577|1588|1899|1522|1544|1644|1661|1660|1599|1800|1688|1666|1668|1555|1855|1811|1877)-\d{4}-\d{4}\b'  # pylint: disable=C0301
PATTERN_DRIVER_NUMS = r'(?<!\+)\d{2}-\d{2}-\d{6}-\d{2}'


PATTERNS = {
    '이메일': PATTERN_EMAILS,
    '주민등록번호': PATTERN_JUMINS,
    '신용카드번호': PATTERN_CREDIT_NUMS,
    '휴대전화번호': PATTERN_CELLPHONE_NUMS,
    '일반전화번호': PATTERN_PHONE_NUMS,
    '전국대표번호': PATTERN_REPRESENTATIVE_PHONE_NUMBERS,
    '운전면허번호': PATTERN_DRIVER_NUMS,
    '계좌번호': [
        PATTERN_GOOKMIN_ACCOUNTS,
        PATTERN_SHINHAN_ACCOUNTS,
        PATTERN_HANA_ACCOUNTS,
        PATTERN_WOORI_ACCOUNTS,
        PATTERN_KAKAO_ACCOUNTS,
        PATTERN_NH_ACCOUNTS
    ]
}
