import sys
import traceback
import pandas as pd
import re
from deep_translator import GoogleTranslator
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QTableWidget, \
    QTableWidgetItem, QMessageBox

# 📌 핸드폰 번호를 010-xxxx-xxxx 형식으로 변환하는 함수
def format_phone_number(phone):
    phone = str(phone).strip()
    phone = re.sub(r'[^0-9]', '', phone)  # 숫자 이외의 문자 제거

    if len(phone) == 10 and phone.startswith("010"):
        return f"{phone[:3]}-{phone[3:6]}-{phone[6:]}"

    elif len(phone) == 10:  # 10자리 번호인데 "010"이 없으면 강제로 010 추가
        return f"010-{phone[2:6]}-{phone[6:]}"

    elif len(phone) == 11 and phone.startswith("010"):  # 정상적인 11자리 핸드폰 번호
        return f"{phone[:3]}-{phone[3:7]}-{phone[7:]}"

    return phone  # 변환 실패 시 원본 반환


# 📌 영문 주소를 한글 발음으로 변환
def translate_english_to_korean(address):
    try:
        translated = GoogleTranslator(source='en', target='ko').translate(address)
        return translated
    except Exception as e:
        print(f"번역 오류 발생: {e}")
        traceback.print_exc()
        return address  # 번역 실패 시 원본 유지


# 📌 영어와 한글 주소 분리 및 변환 후 병합
def separate_and_convert_address(address):
    try:
        if re.search(r'[a-zA-Z]', address):  # ✅ 영문 포함된 경우
            translated_address = translate_english_to_korean(address)
            return translated_address

        return address  # ✅ 한글 주소는 그대로 유지

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
                self.df = pd.read_excel(filePath, dtype=str)
                self.label.setText(f'파일 로드 완료: {filePath}')
                self.displayData()

        except Exception as e:
            print(f"엑셀 파일 로드 오류: {e}")
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
            print(f"데이터 표시 오류: {e}")
            traceback.print_exc()
            QMessageBox.critical(self, "오류", f"데이터를 표시하는 중 오류 발생: {e}")

    def convertData(self):
        try:
            if hasattr(self, 'df'):
                self.label.setText('변환 중...')
                QMessageBox.information(self, "변환 진행 중", "데이터 변환이 진행 중입니다. 잠시 기다려 주세요.")

                if "수령자휴대폰번호" in self.df.columns:
                    self.df["수령자휴대폰번호"] = self.df["수령자휴대폰번호"].apply(format_phone_number)

                if "주소" in self.df.columns:
                    self.df["변환된 주소"] = self.df["주소"].apply(separate_and_convert_address)

                self.displayData()
                self.label.setText('변환 완료!')
                QMessageBox.information(self, "변환 완료", "데이터 변환이 완료되었습니다!")

        except Exception as e:
            print(f"데이터 변환 오류: {e}")
            traceback.print_exc()
            QMessageBox.critical(self, "오류", f"데이터 변환 중 오류 발생: {e}")

    def saveExcel(self):
        try:
            if hasattr(self, 'df'):
                options = QFileDialog.Options()
                filePath, _ = QFileDialog.getSaveFileName(self, "변환된 파일 저장", "converted.xlsx",
                                                          "Excel Files (*.xlsx);;All Files (*)", options=options)
                if filePath:
                    if "수령자휴대폰번호" in self.df.columns:
                        self.df["수령자휴대폰번호"] = self.df["수령자휴대폰번호"].astype(str)

                    self.df.to_excel(filePath, index=False, sheet_name='변환된 데이터')
                    self.label.setText(f'변환된 파일 저장 완료: {filePath}')
                    QMessageBox.information(self, "저장 완료", "변환된 파일이 저장되었습니다!")

        except Exception as e:
            print(f"엑셀 저장 오류: {e}")
            traceback.print_exc()
            QMessageBox.critical(self, "오류", f"엑셀 저장 중 오류 발생: {e}")


if __name__ == '__main__':
    try:
        app = QApplication(sys.argv)
        ex = ExcelConverterApp()
        ex.show()
        sys.exit(app.exec_())

    except Exception as e:
        print(f"프로그램 실행 오류: {e}")
        traceback.print_exc()
        input("오류가 발생했습니다. Enter 키를 눌러 종료하세요...")
