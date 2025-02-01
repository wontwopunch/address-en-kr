import sys
import traceback
import requests
import pandas as pd
import re
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QTableWidget, \
    QTableWidgetItem, QMessageBox

# 구글 지오코딩 API 키 설정
GOOGLE_API_KEY = "API_KEY"


# 📌 핸드폰 번호를 010-xxxx-xxxx 형식으로 변환하는 함수
def format_phone_number(phone):
    phone = str(phone).strip()
    phone = re.sub(r'[^0-9]', '', phone)  # 숫자 이외의 문자 제거

    if len(phone) == 10:
        return f"010-{phone[3:6]}-{phone[6:]}"
    elif len(phone) == 11 and phone.startswith("010"):
        return f"{phone[:3]}-{phone[3:7]}-{phone[7:]}"
    return phone  # 변환 실패 시 원본 반환


# 📌 대한민국 제거 및 주소 세부 정리 함수
def clean_address(address):
    address = address.replace("대한민국 ", "").strip()  # "대한민국" 제거
    address = re.sub(r'\s+', ' ', address)  # 중복 공백 제거
    return address


# 📌 구글 API를 이용해 주소 변환
def get_korean_address_google(address):
    try:
        url = f"https://maps.googleapis.com/maps/api/geocode/json?address={address}&key={GOOGLE_API_KEY}&language=ko"
        response = requests.get(url)
        result = response.json()
        print("구글 API 응답:", result)  # ✅ 응답 출력 (디버깅용)

        if "results" in result and result["results"]:
            formatted_address = result["results"][0]["formatted_address"]
            return clean_address(formatted_address)
        return "변환 실패"

    except Exception as e:
        print(f"주소 변환 중 오류 발생: {e}")
        traceback.print_exc()
        return "변환 실패"


# 📌 영어와 한글 주소 분리 및 변환 후 병합
def separate_and_convert_address(address):
    try:
        parts = address.split()
        english_part = []
        korean_part = []

        for part in parts:
            if re.search(r'[a-zA-Z]', part):
                english_part.append(part)
            else:
                korean_part.append(part)

        if english_part:
            converted_address = get_korean_address_google(' '.join(english_part))  # 영어 주소 변환
            if converted_address != "변환 실패":
                # 중복되는 지역 정보를 제거하고 병합
                converted_main = converted_address.split()
                combined_address = []

                for part in converted_main:
                    if part not in korean_part:
                        combined_address.append(part)

                combined_address.extend(korean_part)
                return re.sub(r'\s+', ' ', ' '.join(combined_address))  # ✅ 공백 정리
        return address  # 변환 실패 시 원본 유지

    except Exception as e:
        print(f"주소 변환 중 오류 발생: {e}")
        traceback.print_exc()
        return address


# PyQt GUI 생성
class ExcelConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.label = QLabel('엑셀 파일을 선택하세요')
        layout.addWidget(self.label)

        self.btnLoad = QPushButton('엑셀 파일 선택')
        self.btnLoad.clicked.connect(self.loadExcel)
        layout.addWidget(self.btnLoad)

        self.btnConvert = QPushButton('데이터 변환')
        self.btnConvert.clicked.connect(self.convertData)
        layout.addWidget(self.btnConvert)

        self.btnSave = QPushButton('변환된 데이터 저장')
        self.btnSave.clicked.connect(self.saveExcel)
        layout.addWidget(self.btnSave)

        self.tableWidget = QTableWidget()
        layout.addWidget(self.tableWidget)

        self.setLayout(layout)
        self.setWindowTitle('엑셀 데이터 변환기')
        self.setGeometry(100, 100, 600, 400)

    def loadExcel(self):
        try:
            options = QFileDialog.Options()
            filePath, _ = QFileDialog.getOpenFileName(self, "엑셀 파일 선택", "", "Excel Files (*.xlsx);;All Files (*)",
                                                      options=options)
            if filePath:
                self.filePath = filePath
                self.df = pd.read_excel(filePath, dtype=str)  # ✅ 엑셀에서 0이 사라지는 문제 해결
                self.label.setText(f'파일 로드 완료: {filePath}')
                self.displayData()

        except Exception as e:
            print(f"엑셀 파일 로드 중 오류 발생: {e}")
            traceback.print_exc()
            QMessageBox.critical(self, "오류", f"엑셀 파일을 로드하는 중 오류 발생: {e}")

    def displayData(self):
        try:
            if hasattr(self, 'df'):
                self.tableWidget.setRowCount(self.df.shape[0])
                self.tableWidget.setColumnCount(self.df.shape[1])
                self.tableWidget.setHorizontalHeaderLabels(self.df.columns)

                for row in range(self.df.shape[0]):
                    for col in range(self.df.shape[1]):
                        self.tableWidget.setItem(row, col, QTableWidgetItem(str(self.df.iat[row, col])))

        except Exception as e:
            print(f"데이터 표시 중 오류 발생: {e}")
            traceback.print_exc()
            QMessageBox.critical(self, "오류", f"데이터를 표시하는 중 오류 발생: {e}")

    def convertData(self):
        try:
            if hasattr(self, 'df'):
                self.label.setText('변환 중...')  # ✅ 변환 중 표시
                QMessageBox.information(self, "변환 진행 중", "데이터 변환이 진행 중입니다. 잠시 기다려 주세요.")

                self.df["수령자휴대폰번호"] = self.df["수령자휴대폰번호"].apply(format_phone_number)
                self.df["변환된 주소"] = self.df["주소"].apply(separate_and_convert_address)  # ✅ 영어와 한글 분리 후 변환

                self.displayData()
                self.label.setText('변환 완료!')  # ✅ 변환 완료 표시
                QMessageBox.information(self, "변환 완료", "데이터 변환이 완료되었습니다!")

        except Exception as e:
            print(f"데이터 변환 중 오류 발생: {e}")
            traceback.print_exc()
            QMessageBox.critical(self, "오류", f"데이터 변환 중 오류 발생: {e}")

    def saveExcel(self):
        try:
            if hasattr(self, 'df'):
                options = QFileDialog.Options()
                filePath, _ = QFileDialog.getSaveFileName(self, "변환된 파일 저장", "converted.xlsx",
                                                          "Excel Files (*.xlsx);;All Files (*)", options=options)
                if filePath:
                    self.df.to_excel(filePath, index=False)
                    self.label.setText(f'변환된 파일 저장 완료: {filePath}')
                    QMessageBox.information(self, "저장 완료", "변환된 파일이 저장되었습니다!")

        except Exception as e:
            print(f"엑셀 저장 중 오류 발생: {e}")
            traceback.print_exc()
            QMessageBox.critical(self, "오류", f"엑셀 저장 중 오류 발생: {e}")


if __name__ == '__main__':
    try:
        app = QApplication(sys.argv)
        ex = ExcelConverterApp()
        ex.show()
        sys.exit(app.exec_())

    except Exception as e:
        print(f"프로그램 실행 중 오류 발생: {e}")
        traceback.print_exc()
        input("오류가 발생했습니다. Enter 키를 눌러 종료하세요...")
