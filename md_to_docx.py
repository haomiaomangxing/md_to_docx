import sys
import os
import tempfile
import subprocess
from PySide6.QtWidgets import (QApplication, QWidget, QVBoxLayout,
                               QTextEdit, QPushButton, QFileDialog, QMessageBox)
from PySide6.QtCore import QFile, Qt

# 内置Pandoc路径配置（打包时需要包含pandoc.exe）
if getattr(sys, 'frozen', False):
    BUNDLE_DIR = sys._MEIPASS
else:
    BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))

PANDOC_PATH = os.path.join(BUNDLE_DIR, 'pandoc.exe')

class MdToDocxConverter(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.check_pandoc()
        self.check_template()  # 新增模板检查

    def check_template(self):
        # """检查模板文件是否存在"""
        self.template_path = os.path.join(BUNDLE_DIR, "reference.docx")
        if not os.path.exists(self.template_path):
            QMessageBox.critical(
                self,
                "缺少模板文件",
                "未找到reference.docx模板文件，请确保模板文件与程序在同一目录",
                QMessageBox.Ok
            )
            self.convert_btn.setEnabled(False)
            
    def initUI(self):
        self.setWindowTitle("MD转DOCX工具")
        self.setGeometry(300, 300, 600, 400)

        layout = QVBoxLayout()
        
        # 输入文本框
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("请在此粘贴Markdown内容...")
        layout.addWidget(self.text_edit)

        # 转换按钮
        self.convert_btn = QPushButton("转换为DOCX")
        self.convert_btn.clicked.connect(self.convert_md)
        layout.addWidget(self.convert_btn)

        self.setLayout(layout)

    def check_pandoc(self):
        if not os.path.exists(PANDOC_PATH):
            QMessageBox.critical(
                self,
                "错误",
                "未找到Pandoc引擎，请确保pandoc.exe与程序在同一目录",
                QMessageBox.Ok
            )
            self.convert_btn.setEnabled(False)

    def convert_md(self):
        # 获取输入内容
        md_content = self.text_edit.toPlainText().strip()
        if not md_content:
            QMessageBox.warning(self, "警告", "输入内容不能为空！", QMessageBox.Ok)
            return

        # 选择保存路径
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存文件",
            os.path.expanduser("~/untitled.docx"),
            "Word文档 (*.docx)"
        )
        if not save_path:
            return

        # 创建临时文件
        with tempfile.NamedTemporaryFile(mode="w", encoding='utf-8', delete=False, suffix=".md") as tmp:
            tmp.write(md_content)
            tmp_path = tmp.name

        try:
            # 执行pandoc转换
            subprocess.run(
            [
                PANDOC_PATH,
                tmp_path,
                "-o", save_path,
                "--from", "markdown+raw_attribute-raw_html+emoji+tex_math_single_backslash",
                "--metadata", "lang=zh-CN",
                "--reference-doc", self.template_path,  # 关键修改点
                "--citeproc",  # 使用citeproc替代已弃用的pandoc-citeproc
            ],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            QMessageBox.information(
                self,
                "转换成功",
                f"文件已保存至：\n{save_path}",
                QMessageBox.Ok
            )
        except subprocess.CalledProcessError as e:
            QMessageBox.critical(
                self,
                "转换失败",
                f"Pandoc错误：\n{e.stderr.decode()}",
                QMessageBox.Ok
            )
        finally:
            os.remove(tmp_path)



if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MdToDocxConverter()
    window.show()
    sys.exit(app.exec())
