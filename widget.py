# This Python file uses the following encoding: utf-8
import sys
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication, QWidget, QFormLayout, QVBoxLayout, QLabel, QFontComboBox,
    QListWidget, QListWidgetItem, QHBoxLayout, QPushButton, QSpinBox, QFileDialog,
    QGroupBox, QSpacerItem, QSizePolicy,QMessageBox
)
import csv2pptx
import random


class Widget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.data = []
        self.samples = []

        self.setWindowTitle("成语PPT生成系统 IdiomUI")

        main_layout = QVBoxLayout()

        # Group for form elements
        form_group_box = QGroupBox("设置参数")
        form_layout = QFormLayout()
        form_layout.setLabelAlignment(Qt.AlignmentFlag.AlignLeft)

        self.select_num = QSpinBox()
        self.select_num.setValue(20)
        self.select_num.setMaximum(2**31 - 1)

        form_layout.addRow(QLabel("成语数量:"), self.select_num)

        # Font selection
        self.font_edit = QFontComboBox()
        self.font_edit.setCurrentFont("Kai")
        form_layout.addRow(QLabel("字体："), self.font_edit)

        # Row and Column Settings
        self.row_edit = QSpinBox()
        self.row_edit.setMaximum(2**31 - 1)
        self.row_edit.setValue(1)
        form_layout.addRow(QLabel("起始阅读行号:"), self.row_edit)

        self.column_edit = QSpinBox()
        self.column_edit.setMaximum(2**31 - 1)
        self.column_edit.setValue(1)
        form_layout.addRow(QLabel("数据源列号:"), self.column_edit)

        self.fontsize_edit = QSpinBox()
        self.fontsize_edit.setMaximum(2**31 - 1)
        self.fontsize_edit.setValue(150)
        form_layout.addRow(QLabel("字体大小:"), self.fontsize_edit)

        # Add the new SpinBox for additional settings
        self.text_width = QSpinBox()
        self.text_width.setMaximum(2**31 - 1)
        self.text_width.setValue(720)
        form_layout.addRow(QLabel("文本框宽度:"), self.text_width)

        self.text_height = QSpinBox()
        self.text_height.setMaximum(2**31 - 1)
        self.text_height.setValue(200)
        form_layout.addRow(QLabel("文本框高度:"), self.text_height)

        self.text_top = QSpinBox()
        self.text_top.setMaximum(2**31 - 1)
        self.text_top.setValue(180)
        form_layout.addRow(QLabel("文本上间距:"), self.text_top)

        self.text_left = QSpinBox()
        self.text_left.setMaximum(2**31 - 1)
        self.text_left.setValue(0)
        form_layout.addRow(QLabel("文本左间距:"), self.text_left)

        form_group_box.setLayout(form_layout)
        main_layout.addWidget(form_group_box)

        # Preview Area
        preview_group_box = QGroupBox("数据")
        preview_layout = QHBoxLayout()

        # Loaded Items
        loaded_label = QLabel("已加载：")
        loaded_label.setAlignment(Qt.AlignmentFlag.AlignTop)
        preview_layout.addWidget(loaded_label)

        self.list_widget = QListWidget()
        preview_layout.addWidget(self.list_widget)

        # Selected Items
        selected_label = QLabel("已选取：")
        selected_label.setAlignment(Qt.AlignmentFlag.AlignTop)
        preview_layout.addWidget(selected_label)

        self.selected_list_widget = QListWidget()
        preview_layout.addWidget(self.selected_list_widget)

        preview_group_box.setLayout(preview_layout)
        main_layout.addWidget(preview_group_box)

        # File Operations
        op_area = QHBoxLayout()
        load_btn = QPushButton("加载成语文件 (.csv)")
        random_btn = QPushButton("随机选取")
        export_btn = QPushButton("导出 (.pptx)")

        load_btn.clicked.connect(self.open_csv)
        random_btn.clicked.connect(self.random_select)
        export_btn.clicked.connect(self.export_file)

        op_area.addWidget(load_btn)
        op_area.addWidget(random_btn)
        op_area.addWidget(export_btn)

        # Add a spacer to balance the buttons
        op_area.addSpacerItem(QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding))

        main_layout.addLayout(op_area)

        # Label to show the loaded file path
        file_panel = QHBoxLayout()
        file_panel.setAlignment(Qt.AlignmentFlag.AlignLeft)

        self.loadfile_label = QLabel("(未加载)")
        file_panel.addWidget(QLabel("加载文件:"))
        file_panel.addWidget(self.loadfile_label)

        main_layout.addLayout(file_panel)

        about_label = QLabel("Powered by QixinyNet 由七夕泥开发 Version 1.0.0 2024/07/02")
        about_label.setStyleSheet('font-size: 8px;')
        main_layout.addWidget(about_label)

        # Set the main layout
        self.setLayout(main_layout)

    def open_csv(self):
        dialog = QFileDialog(self)
        dialog.setNameFilter("CSV Files (*.csv)")
        filename, _ = dialog.getOpenFileName(self, '选择 CSV 文件', '', 'CSV Files (*.csv)')
        if filename:
            self.loadfile_label.setText(filename)

            self.data = csv2pptx.read_csv(filename, self.row_edit.value(), self.column_edit.value())
            self.list_widget.clear()
            for item in self.data:
                self.list_widget.addItem(QListWidgetItem(item))

    def random_select(self):
        if not self.data:
            return
        self.samples = random.sample(self.data, min(self.select_num.value(), len(self.data)))
        self.selected_list_widget.clear()
        for item in self.samples:
            self.selected_list_widget.addItem(QListWidgetItem(item))

    def export_file(self):
        file_name, _ = QFileDialog.getSaveFileName(self, '保存文件', '', 'PPTX Files (*.pptx);;All Files (*)')
        if file_name:
            font_name = self.font_edit.currentText()
            font_size = self.fontsize_edit.value()
            presentation = csv2pptx.generate_presentation_by_array(
                self.samples,
                font_size,
                self.text_width.value(),
                self.text_height.value(),
                self.text_top.value(),
                self.text_left.value(),
                font_name
            )
            # Save the presentation to the specified file
            presentation.save(file_name)
            print(f"Presentation saved to {file_name}")
            QMessageBox.information(self, '信息', '成语幻灯片导出成功')


if __name__ == "__main__":
    app = QApplication(sys.argv)
    widget = Widget()
    widget.show()
    sys.exit(app.exec())
