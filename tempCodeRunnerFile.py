    def update_recent_items(self, scroll_to_bottom=False):
        """최근 항목을 테이블 형식으로 업데이트 (최신 항목이 아래로 오도록)"""
        recent_items = get_recent_items(self.current_file)  # 최근 데이터 가져오기
        lines = list(reversed(recent_items.split("\n")))  # ✅ 최신 항목이 아래로 가도록 정렬

        self.recent_table.setRowCount(len(lines))  # 행 개수 설정

        for row_idx, line in enumerate(lines):
            parts = line.split()
            if len(parts) < 4:
                continue  # 데이터 부족하면 건너뜀

            date, barcode, time, count = parts[:4]  # ✅ 컬럼 순서 조정 (가독성 향상)
            self.recent_table.setItem(row_idx, 0, QTableWidgetItem(barcode))
            self.recent_table.setItem(row_idx, 1, QTableWidgetItem(date))
            self.recent_table.setItem(row_idx, 2, QTableWidgetItem(time))
            self.recent_table.setItem(row_idx, 3, QTableWidgetItem(count))

        if scroll_to_bottom:
            scrollbar = self.recent_table.verticalScrollBar()
            scrollbar.setValue(scrollbar.maximum())  # ✅ 스크롤을 최신 항목으로 이동


        def perform_search(self, barcode):
            """입력된 바코드를 검색하여 테이블에 표시"""
            result = find_barcode_in_excel(self.current_file, barcode)

            if result:
                self.update_search_results(result, barcode)
            else:
                QMessageBox.information(self, "검색 결과", "해당 바코드가 없습니다.")

        def update_search_results(self, results, barcode):
            """검색 결과 테이블을 업데이트하고 '최대 중복' 컬럼 추가"""
            self.search_table.setColumnCount(4)  # ✅ 컬럼 5개로 변경
            self.search_table.setHorizontalHeaderLabels(["날짜", "횟수", "비고", "삭제"])

            self.search_table.setRowCount(len(results))

            for row, (date, count, remark, max_dup) in enumerate(results):
                self.search_table.setItem(row, 0, QTableWidgetItem(date))
                self.search_table.setItem(row, 1, QTableWidgetItem(count))
                self.search_table.setItem(row, 2, QTableWidgetItem(remark))
                # self.search_table.setItem(row, 3, QTableWidgetItem(max_dup))  # ✅ 최대 중복 컬럼 추가

                # 삭제 버튼 추가
                delete_button = QPushButton("X")
                delete_button.setStyleSheet("border: 0px ; border-radius: 5px;font-size: 20px; font-weight: 1000;color: white; background-color: rgb(255, 60, 60);")
                delete_button.clicked.connect(lambda _, r=row: self.delete_barcode_entry(barcode, r))
                self.search_table.setCellWidget(row, 3, delete_button)  # ✅ 5번째 컬럼 (삭제 버튼)

        def delete_barcode_entry(self, barcode, row):
            """선택한 바코드 행을 삭제"""
            date = self.search_table.item(row, 0).text()
            count = self.search_table.item(row, 1).text()

            if delete_barcode_from_excel(self.current_file, barcode, date, count):
                QMessageBox.information(self, "삭제 완료", "바코드 항목이 삭제되었습니다.")
                self.perform_search(barcode)  # ✅ 검색 결과 갱신
                self.update_recent_items(scroll_to_bottom=True)  # ✅ 삭제 후 최근 항목도 갱신
            else:
                QMessageBox.warning(self, "삭제 실패", "삭제할 수 없습니다.")

        def changeEvent(self, event):
            """창 활성화 시 바코드 입력창에 포커스 강제 설정"""
            if event.type() == QEvent.ActivationChange and self.isActiveWindow():
                self.input_line.setFocus()
            super().changeEvent(event)
