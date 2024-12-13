import os
import sys
from datetime import datetime
from openpyxl import load_workbook, Workbook
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QLineEdit, QPushButton, QLabel, QTextEdit, QFileDialog, QWidget
)
from PySide6.QtCore import QFileSystemWatcher

DEFAULT_EXCEL_FILE = "data.xlsm"  # 기본 데이터베이스 파일 이름
QSS_FILE = "style.qss"  # QSS 파일 경로


class BarcodeApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("바코드 처리 프로그램")
        self.current_file = DEFAULT_EXCEL_FILE  # 현재 데이터베이스 파일

        # 중앙 위젯 설정
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        self.layout = QVBoxLayout()

        # 바코드 입력창
        self.input_line = QLineEdit()
        self.input_line.setPlaceholderText("바코드를 입력하세요")
        self.layout.addWidget(self.input_line)

        # 처리 버튼
        self.button = QPushButton("처리")
        self.button.clicked.connect(self.on_process_clicked)
        self.layout.addWidget(self.button)

        # 결과 표시
        self.result_label = QLabel("")
        self.layout.addWidget(self.result_label)

        # 최근 항목 표시
        self.recent_label = QLabel("최근 항목")
        self.layout.addWidget(self.recent_label)

        self.recent_items_text = QTextEdit()
        self.recent_items_text.setReadOnly(True)
        self.layout.addWidget(self.recent_items_text)

        central_widget.setLayout(self.layout)

        # 메뉴바 추가
        self.create_menu_bar()

        self.update_recent_items()

    def create_menu_bar(self):
        """메뉴바 생성"""
        menu_bar = self.menuBar()

        # 파일 메뉴
        file_menu = menu_bar.addMenu("파일")

        # "다른 이름으로 저장하기" 메뉴
        save_action = file_menu.addAction("다른 이름으로 저장하기")
        save_action.triggered.connect(self.save_as)

        # "불러오기" 메뉴
        load_action = file_menu.addAction("불러오기")
        load_action.triggered.connect(self.load_file)

    def update_file_name_label(self):
        """현재 파일명을 메뉴바 옆에 업데이트"""
        self.file_name_label.setText(f"현재 파일: {os.path.basename(self.current_file)}")

    def on_process_clicked(self):
        barcode = self.input_line.text().strip()
        if not barcode:
            self.result_label.setText("바코드를 입력해주세요.")
            return

        msg = process_barcode(barcode, self.current_file)
        self.result_label.setText(msg)
        self.input_line.clear()
        self.update_recent_items()

    def save_as(self):
        """다른 이름으로 파일 저장"""
        file_path, _ = QFileDialog.getSaveFileName(self, "파일 저장", "", "Excel Files (*.xlsm *.xlsx)")
        if file_path:
            try:
                wb = load_workbook(self.current_file, keep_vba=True)
                wb.save(file_path)
                wb.close()
                self.result_label.setText(f"파일이 저장되었습니다: {file_path}")
            except Exception as e:
                self.result_label.setText(f"파일 저장 중 오류 발생: {str(e)}")

    def load_file(self):
        """다른 파일 불러오기"""
        file_path, _ = QFileDialog.getOpenFileName(self, "파일 열기", "", "Excel Files (*.xlsm *.xlsx)")
        if file_path:
            self.current_file = file_path
            self.update_recent_items()
            self.result_label.setText(f"파일이 변경되었습니다: {file_path}")

    def update_recent_items(self):
        """최근 항목 업데이트"""
        recent_items = get_recent_items(self.current_file)
        self.recent_items_text.setText(recent_items)


def ensure_excel_file(file_path):
    """데이터 관리를 위한 엑셀 파일 생성 확인"""
    try:
        wb = load_workbook(file_path, keep_vba=True)
        wb.close()
    except FileNotFoundError:
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "바코드"
        ws["C1"] = "날짜/시간"
        ws["D1"] = "중복횟수"
        ws["E1"] = "상태"
        wb.save(file_path)


def process_barcode(barcode, file_path):
    """바코드를 엑셀 데이터베이스에 추가하고 처리"""
    wb = load_workbook(file_path, keep_vba=True)
    ws = wb.active

    # 다음 추가할 행 찾기
    max_row = ws.max_row + 1

    # 중복 횟수 계산
    count = sum(1 for row in range(2, max_row) if ws.cell(row=row, column=1).value == barcode)

    # 데이터 추가
    ws.cell(row=max_row, column=1, value=barcode)
    ws.cell(row=max_row, column=3, value=datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    ws.cell(row=max_row, column=4, value=count + 1)

    wb.save(file_path)
    return f"바코드 {barcode} 처리 완료 (중복: {count + 1})"


def get_recent_items(file_path, limit=10):
    """최근 항목 가져오기"""
    wb = load_workbook(file_path, keep_vba=True)
    ws = wb.active

    max_row = ws.max_row
    recent_items = []

    for row in range(max(max_row - limit + 1, 2), max_row + 1):  # 마지막 `limit`개만 가져옴
        barcode = ws.cell(row=row, column=1).value
        timestamp = ws.cell(row=row, column=3).value
        duplicate_count = ws.cell(row=row, column=4).value
        recent_items.append(f"{barcode} | {timestamp} | {duplicate_count}")

    wb.close()
    return "\n".join(recent_items)


def load_qss(app, qss_file):
    """QSS 파일을 로드하여 스타일 적용"""
    try:
        with open(qss_file, "r", encoding="utf-8") as f:
            app.setStyleSheet(f.read())
    except FileNotFoundError:
        print(f"QSS 파일을 찾을 수 없습니다: {qss_file}")


if __name__ == "__main__":
    ensure_excel_file(DEFAULT_EXCEL_FILE)
    app = QApplication(sys.argv)

    # QSS 로드 및 핫 리로드 설정
    load_qss(app, QSS_FILE)
    watcher = QFileSystemWatcher([QSS_FILE])
    watcher.fileChanged.connect(lambda: load_qss(app, QSS_FILE))  # QSS 변경 시 자동 적용

    window = BarcodeApp()
    window.show()
    sys.exit(app.exec())
