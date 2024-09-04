from module.processing_folder import processing_folder
"""main 함수"""

if __name__ == "__main__":
    folder_path = input("작업할 폴더 경로를 입력하세요: ")
    excel_file = input(
        "엑셀파일 경로를 입력하세요(확장자포함. 파일이 존재하지 않을 경우 새로 생성): ")
    processing_folder(folder_path, excel_file)
    print(f"{excel_file}에 개인정보목록이 생성되었습니다.")
