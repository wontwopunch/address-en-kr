import sys
import traceback
import pandas as pd
import re
import asyncio
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


# 📌 번역 제외할 주소 구성요소 리스트 (한글 그대로 유지)
EXCLUDE_WORDS = ["읍", "면", "동", "리", "길", "호", "상가", "빌리지", "타운", "아파트", "부동산", "원룸"]

# 📌 반드시 번역해야 하는 행정 구역 (도, 시, 군, 구)
TRANSLATABLE_ADMIN_REGIONS = ["do", "si", "gun", "gu"]


# 📌 주소를 번역 전, 예외처리할 함수
def preprocess_address_for_translation(address):
    words = address.split()
    translatable_part = []
    preserved_part = []

    for word in words:
        # ✅ 번역에서 제외할 부분(한국어 주소 요소)
        if any(excluded in word for excluded in EXCLUDE_WORDS):
            preserved_part.append(word)
        else:
            translatable_part.append(word)  # ✅ 번역할 부분

    return ' '.join(translatable_part), ' '.join(preserved_part)


# 📌 비동기 번역 처리 함수
async def async_translate_english_to_korean(address):
    try:
        translatable_part, preserved_part = preprocess_address_for_translation(address)

        # ✅ 번역 수행 (도/시/군/구 포함)
        translated = await asyncio.to_thread(GoogleTranslator(source='en', target='ko').translate, translatable_part)

        # ✅ 번역된 결과에 여전히 영문이 남아 있는 경우, 다시 번역 시도
        if any(region in translated for region in TRANSLATABLE_ADMIN_REGIONS):
            translated = await asyncio.to_thread(GoogleTranslator(source='en', target='ko').translate, translated)

        # ✅ 번역된 부분 + 유지한 원본 부분 합치기
        return f"{translated} {preserved_part}".strip()

    except Exception as e:
        print(f"번역 오류 발생: {e}")
        traceback.print_exc()
        return address  # 번역 실패 시 원본 유지


# 📌 모든 주소를 한 번에 번역 (비동기 처리)
async def async_separate_and_convert_addresses(addresses):
    tasks = [async_translate_english_to_korean(addr) if re.search(r'[a-zA-Z]', addr) else asyncio.sleep(0) for addr in addresses]
    results = await asyncio.gather(*tasks)  # ✅ `await` 추가
    return results


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
        self.btnSave.clicked.connect(self.saveExcel)  # ✅ saveExcel 함수 추가
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
                    addresses = self.df["주소"].tolist()
                    loop = asyncio.new_event_loop()
                    asyncio.set_event_loop(loop)
                    translated_addresses = loop.run_until_complete(async_separate_and_convert_addresses(addresses))  # ✅ 수정된 코드
                    self.df["변환된 주소"] = translated_addresses

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
    app = QApplication(sys.argv)
    ex = ExcelConverterApp()
    ex.show()
    sys.exit(app.exec_())
