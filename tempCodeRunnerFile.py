
        # 검색
        self.search_button = QPushButton("검색")
        self.search_button.clicked.connect(self.open_search_dialog)
        self.layout.addWidget(self.search_button)