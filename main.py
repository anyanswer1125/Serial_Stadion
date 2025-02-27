import os
import sys
from datetime import datetime
from openpyxl import load_workbook, Workbook
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QLineEdit, QPushButton, QLabel, QTextEdit, QFileDialog, QWidget, QInputDialog
)
from PySide6.QtCore import QFileSystemWatcher
from PySide6.QtCore import QEvent
from PySide6.QtWidgets import QMessageBox

DEFAULT_EXCEL_FILE = "data.xlsm"  # 기본 데이터베이스 파일 이름
QSS_FILE = "style.qss"  # QSS 파일 경로


class BarcodeApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("할인쿠폰R0.2_preview")
        self.current_file = DEFAULT_EXCEL_FILE  # 현재 데이터베이스 파일
        
        
        self.watcher = QFileSystemWatcher([self.current_file])
        self.watcher.fileChanged.connect(self.on_file_changed)

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
        self.activateWindow()
        self.input_line.setFocus()
        self.input_line.returnPressed.connect(self.on_process_clicked)
        
        self.file_label = QLabel(f"불러온 파일: {os.path.basename(self.current_file)}")
        self.layout.addWidget(self.file_label)


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

        msg = process_barcode(barcode, self.current_file, self)
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
        file_path, _ = QFileDialog.getOpenFileName(self, "파일 열기", "", "Excel Files (*.xlsm *.xlsx)")
        if file_path:
            self.current_file = file_path
            self.file_label.setText(f"불러온 파일: {os.path.basename(self.current_file)}")  # 파일명 업데이트
            self.watcher.removePaths(self.watcher.files())  # 기존 감시자 제거
            self.watcher.addPath(file_path)  # 새 파일 감시자 등록
            self.update_recent_items()
            self.result_label.setText(f"파일이 변경되었습니다: {file_path}")

    # def update_recent_items(self):
        """최근 항목 업데이트 후 스크롤 위치 유지"""
    #     scrollbar = self.recent_items_text.verticalScrollBar()
    #     scroll_position = scrollbar.value()  # 현재 스크롤 위치 저장

    #     recent_items = get_recent_items(self.current_file)
    #     self.recent_items_text.setText(recent_items)

    #     scrollbar.setValue(scroll_position)  # 이전 스크롤 위치로 복원
    
    
    def update_recent_items(self):
        """최근 항목 업데이트 후 스크롤을 맨 아래로 이동"""
        recent_items = get_recent_items(self.current_file)
        self.recent_items_text.setText(recent_items)

        scrollbar = self.recent_items_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())  # 스크롤을 맨 아래로 이동

    
    def prompt_max_duplicate(self):
    # 최대 중복 횟수를 입력받는 팝업
        max_duplicate, ok = QInputDialog.getInt(
            self,                           # 부모 위젯 (여기서는 BarcodeApp 클래스의 인스턴스)
            "최대 중복 횟수 설정",           # 팝업창의 제목 (타이틀)
            "최대 중복 횟수를 입력해주세요:", # 사용자에게 보여줄 메시지 (라벨 텍스트)
            10,                              # 기본값 (기본으로 입력창에 표시되는 값, 여기서는 10)
            1,                              # 최소값 (사용자가 입력할 수 있는 가장 작은 값, 여기서는 1)
            1000,                           # 최대값 (사용자가 입력할 수 있는 가장 큰 값, 여기서는 1000)
            1                               # 스텝 값 (입력 시 숫자가 얼마나 증가/감소하는지, 여기서는 1씩)
        )                                                                                           
        return max_duplicate if ok else None
    def on_file_changed(self):
        # 엑셀 파일이 변경되었을 때 자동으로 최근 항목 업데이트
        self.update_recent_items()
    def changeEvent(self, event):
        """창 활성화 시 바코드 입력창에 포커스를 강제로 설정"""
        if event.type() == QEvent.ActivationChange and self.isActiveWindow():
            self.input_line.setFocus()
        super().changeEvent(event)

def ensure_excel_file(file_path):
    """데이터 관리를 위한 엑셀 파일 생성 확인"""
    try:
        wb = load_workbook(file_path, keep_vba=True)
        ws = wb.active

        # 기존에 F1까지 사용하던 헤더를 새로운 형식으로 변경
        if ws.max_column < 6:
            ws["A1"] = "바코드"
            ws["B1"] = "날짜"
            ws["C1"] = "횟수"
            ws["D1"] = "비고"
            ws["E1"] = "최대중복"
            wb.save(file_path)
        wb.close()
    except FileNotFoundError:
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "바코드"
        ws["B1"] = "날짜"
        ws["C1"] = "횟수"
        ws["D1"] = "비고"
        ws["E1"] = "최대중복"
        wb.save(file_path)

def show_max_duplicate_popup(app_window, barcode, max_duplicate):
    """최대 중복 횟수 초과 시 경고 팝업"""
    msg_box = QMessageBox(app_window)
    msg_box.setIcon(QMessageBox.Warning)
    msg_box.setWindowTitle("최대 중복 횟수 초과")
    msg_box.setText(f"바코드 {barcode}의 최대 중복 횟수 ({max_duplicate})를 초과했습니다.\n더 이상 등록할 수 없습니다.")
    msg_box.setStandardButtons(QMessageBox.Ok)
    msg_box.exec()

def process_barcode(barcode, file_path, app_window):
    """바코드를 엑셀 데이터베이스에 추가하고 처리"""
    wb = load_workbook(file_path, keep_vba=True)
    ws = wb.active

    now_str = datetime.now().strftime('%Y-%m-%d %H:%M')

    # 바코드별 중복 횟수 계산
    barcode_counts = {}
    for row in range(2, ws.max_row + 1):
        code = ws.cell(row=row, column=1).value
        if code:
            barcode_counts[code] = barcode_counts.get(code, 0) + 1

    count = barcode_counts.get(barcode, 0) + 1  # 현재 바코드 카운트 증가
    max_duplicate = 10  # 기본 최대 중복 횟수 (변경 가능)
    
    if not barcode_counts.get(barcode):
        # 신규 등록
        max_duplicate = app_window.prompt_max_duplicate()
        if max_duplicate is None:
            wb.close()
            return "등록이 취소되었습니다."
        
        remark = "신규 등록"
    else:
        # 중복 확인
        existing_rows = [row for row in range(2, ws.max_row + 1) if ws.cell(row=row, column=1).value == barcode]
        max_duplicate = int(ws.cell(existing_rows[0], column=5).value.replace("회", ""))  # 최대 중복 횟수

        if count > max_duplicate:
            # 🚨 팝업창 추가 (최대 한도 초과 시)
            show_max_duplicate_popup(app_window, barcode, max_duplicate)
            wb.close()
            return f"등록 불가: 바코드 {barcode}의 최대 중복 횟수({max_duplicate})를 초과했습니다."

        remark = "최대 한도 도달" if count == max_duplicate else "중복 사용"

    # 엑셀에 기록
    new_row = ws.max_row + 1
    ws.cell(row=new_row, column=1, value=barcode)
    ws.cell(row=new_row, column=2, value=now_str)  # 날짜
    ws.cell(row=new_row, column=3, value=f"{count} 회")  # 횟수
    ws.cell(row=new_row, column=4, value=remark)  # 비고
    ws.cell(row=new_row, column=5, value=f"{max_duplicate}회")  # 최대중복

    wb.save(file_path)
    wb.close()

    return f"바코드 {barcode} 처리 완료 (중복: {count}/{max_duplicate})"

def get_recent_items(file_path, limit=1000):
    """최근 항목 가져오기 (일정한 간격 유지)"""
    wb = load_workbook(file_path, keep_vba=True)
    ws = wb.active

    max_row = ws.max_row
    recent_items = []

    # 칼럼 
    # 너비 설정 (문자열 길이를 균일하게 유지)
    COL_WIDTH = {
        "barcode": 22,   # 바코드 길이 고정
        "timestamp": 18, # YYYY-MM-DD HH:MM
        "count": 12,      # "99 회"
        "status": 20,    # "최대 사용 횟수 도달"
        "max_count": 6   # "99회"
    }

    for row in range(max(max_row - limit + 1, 2), max_row + 1):  # 마지막 `limit`개만 가져옴
        barcode = str(ws.cell(row=row, column=1).value).ljust(COL_WIDTH["barcode"])  # 왼쪽 정렬
        timestamp = str(ws.cell(row=row, column=2).value).rjust(COL_WIDTH["timestamp"])
        duplicate_count = str(ws.cell(row=row, column=3).value).rjust(COL_WIDTH["count"])  # 오른쪽 정렬
        status = str(ws.cell(row=row, column=4).value).rjust(COL_WIDTH["status"])  # 왼쪽 정렬
        max_count = str(ws.cell(row=row, column=5).value).rjust(COL_WIDTH["max_count"])  # 오른쪽 정렬

        # recent_items.append(f"{barcode} {timestamp} {duplicate_count} {status} {max_count}") #백업
        # recent_items.append(f"   {barcode} {timestamp} {duplicate_count}회 {status}") 
        recent_items.append(f"   {barcode} {timestamp} {duplicate_count} {status}") 

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
