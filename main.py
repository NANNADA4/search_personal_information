"""
main 함수. 폴더를 순회하며 hwp, pdf, xlsx 파일에 한해 개인정보를 수집합니다
"""
from module.processing_folder import processing_folder

if __name__ == "__main__":
    print("\n>>>>>>개인정보 추출기<<<<<<\n")
    print("-"*24)
    folder_path = input("작업할 폴더 경로를 입력하세요: ")
    excel_file = input(
        "엑셀파일 경로를 입력하세요(확장자포함. 파일이 존재하지 않을 경우 새로 생성): ")
    processing_folder(folder_path, excel_file)
    print(f"{excel_file}에 개인정보목록이 생성되었습니다.")
