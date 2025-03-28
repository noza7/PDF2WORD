import os
import sys
import platform
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                            QFileDialog, QProgressBar, QMessageBox, QFrame)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QIcon, QFont, QColor, QPalette

class ConversionThread(QThread):
    """用于在后台执行PDF到Word转换的线程"""
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(bool, str)
    
    def __init__(self, pdf_path, word_path):
        super().__init__()
        self.pdf_path = pdf_path
        self.word_path = word_path
    
    def run(self):
        try:
            # 使用Word打开PDF并另存为Word文档
            self.convert_with_word()
            
            # 检查生成的文件是否有效
            if os.path.exists(self.word_path) and os.path.getsize(self.word_path) > 100:
                # 修复上标数字问题
                try:
                    self.fix_superscript_numbers(self.word_path)
                except Exception as e:
                    print(f"修复上标数字失败: {str(e)}")
                
                self.finished_signal.emit(True, "")
                return
            else:
                raise Exception("转换后的文件无效或为空")
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.finished_signal.emit(False, str(e))
    
    def convert_with_word(self):
        """使用Word打开PDF并另存为Word文档"""
        if platform.system() != "Windows":
            raise Exception("Word COM自动化仅在Windows系统上可用")
        
        try:
            import win32com.client
            import pythoncom
            
            # 初始化COM
            pythoncom.CoInitialize()
            
            # 创建Word应用实例
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            # 不再尝试设置可能不存在的自动格式选项
            # 发送进度信号
            self.progress_signal.emit(20)
            
            try:
                # 打开PDF文件
                doc = word.Documents.Open(self.pdf_path)
                
                # 发送进度信号
                self.progress_signal.emit(50)
                
                # 保存为Word文档
                doc.SaveAs2(self.word_path, FileFormat=16)  # 16 表示.docx格式
                
                # 发送进度信号
                self.progress_signal.emit(80)
                
                # 关闭文档
                doc.Close()
            finally:
                # 退出Word应用
                word.Quit()
                
                # 释放COM资源
                pythoncom.CoUninitialize()
        except Exception as e:
            raise Exception(f"使用Word转换失败: {str(e)}")
    
    def fix_superscript_numbers(self, docx_path):
        """修复Word文档中的上标数字，只处理数字而不影响其他格式"""
        import win32com.client
        import pythoncom
        
        # 初始化COM
        pythoncom.CoInitialize()
        
        try:
            # 创建Word应用实例
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            # 打开Word文档
            doc = word.Documents.Open(docx_path)
            
            # 使用查找替换功能专门处理上标数字
            for i in range(10):  # 处理0-9的数字
                # 设置查找条件：查找上标格式的数字
                word.Selection.Find.ClearFormatting()
                word.Selection.Find.Text = str(i)
                word.Selection.Find.SuperScript = True
                
                # 设置替换内容：相同的数字但不是上标
                word.Selection.Find.Replacement.ClearFormatting()
                word.Selection.Find.Replacement.Text = str(i)
                
                # 执行替换
                word.Selection.Find.Execute(
                    FindText=str(i),
                    MatchCase=True,
                    MatchWholeWord=True,
                    MatchWildcards=False,
                    MatchSoundsLike=False,
                    MatchAllWordForms=False,
                    Forward=True,
                    Wrap=1,  # wdFindContinue
                    Format=True,
                    ReplaceWith=str(i),
                    Replace=2  # wdReplaceAll
                )
            
            # 保存文档
            doc.Save()
            
            # 关闭文档
            doc.Close()
        finally:
            # 退出Word应用
            word.Quit()
            
            # 释放COM资源
            pythoncom.CoUninitialize()


class PDFToWordConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF转Word转换器")
        self.setMinimumSize(600, 400)
        
        # 设置应用样式
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QLabel {
                color: #333333;
            }
            QLineEdit {
                padding: 8px;
                border: 1px solid #cccccc;
                border-radius: 4px;
                background-color: white;
            }
            QPushButton {
                padding: 8px 15px;
                border: none;
                border-radius: 4px;
                background-color: #4a86e8;
                color: white;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #3a76d8;
            }
            QPushButton:pressed {
                background-color: #2a66c8;
            }
            QPushButton#convert_btn {
                padding: 10px 20px;
                font-size: 14px;
                background-color: #4CAF50;
            }
            QPushButton#convert_btn:hover {
                background-color: #3d9c40;
            }
            QPushButton#convert_btn:pressed {
                background-color: #2e8c30;
            }
            QProgressBar {
                border: 1px solid #cccccc;
                border-radius: 4px;
                text-align: center;
                background-color: #f0f0f0;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                border-radius: 3px;
            }
        """)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(30, 30, 30, 30)
        main_layout.setSpacing(20)
        
        # 标题
        title_label = QLabel("PDF 转 Word 转换工具")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; margin-bottom: 20px; color: #333;")
        main_layout.addWidget(title_label)
        
        # 分隔线
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        line.setStyleSheet("background-color: #cccccc;")
        main_layout.addWidget(line)
        
        # PDF文件选择
        pdf_layout = QHBoxLayout()
        pdf_label = QLabel("PDF文件:")
        pdf_label.setMinimumWidth(80)
        self.pdf_path_edit = QLineEdit()
        self.pdf_path_edit.setPlaceholderText("选择PDF文件...")
        pdf_browse_btn = QPushButton("浏览...")
        pdf_browse_btn.clicked.connect(self.browse_pdf)
        
        pdf_layout.addWidget(pdf_label)
        pdf_layout.addWidget(self.pdf_path_edit)
        pdf_layout.addWidget(pdf_browse_btn)
        main_layout.addLayout(pdf_layout)
        
        # Word文件选择
        word_layout = QHBoxLayout()
        word_label = QLabel("输出目录:")
        word_label.setMinimumWidth(80)
        self.word_path_edit = QLineEdit()
        self.word_path_edit.setPlaceholderText("选择输出目录...")
        word_browse_btn = QPushButton("浏览...")
        word_browse_btn.clicked.connect(self.browse_word_dir)
        
        word_layout.addWidget(word_label)
        word_layout.addWidget(self.word_path_edit)
        word_layout.addWidget(word_browse_btn)
        main_layout.addLayout(word_layout)
        
        # 转换按钮
        convert_btn_layout = QHBoxLayout()
        convert_btn_layout.addStretch()
        self.convert_btn = QPushButton("开始转换")
        self.convert_btn.setObjectName("convert_btn")
        self.convert_btn.setMinimumWidth(150)
        self.convert_btn.setMinimumHeight(40)
        self.convert_btn.clicked.connect(self.convert)
        convert_btn_layout.addWidget(self.convert_btn)
        convert_btn_layout.addStretch()
        main_layout.addLayout(convert_btn_layout)
        
        # 进度条
        progress_layout = QVBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p%")
        self.progress_bar.setMinimumHeight(20)
        progress_layout.addWidget(self.progress_bar)
        main_layout.addLayout(progress_layout)
        
        # 状态标签
        self.status_label = QLabel("准备就绪")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("color: #666; font-style: italic;")
        main_layout.addWidget(self.status_label)
        
        # 添加依赖库状态信息
        self.dependency_label = QLabel("正在检查依赖库...")
        self.dependency_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.dependency_label.setStyleSheet("color: #888; font-size: 10px;")
        main_layout.addWidget(self.dependency_label)
        
        # 添加弹性空间
        main_layout.addStretch()
        
        # 初始化转换线程
        self.conversion_thread = None
        
        # 检查依赖库
        self.check_dependencies()
    
    def check_dependencies(self):
        """检查必要的依赖库是否已安装"""
        # 检查Word COM自动化
        word_available = False
        if platform.system() == "Windows":
            try:
                import win32com.client
                import pythoncom  # 也检查 pythoncom
                # 尝试创建Word应用实例
                word = win32com.client.Dispatch("Word.Application")
                word.Quit()
                word_available = True
                self.dependency_label.setText("已检测到Microsoft Word，可以开始转换")
                self.dependency_label.setStyleSheet("color: #388e3c; font-size: 10px;")
            except ImportError:
                self.dependency_label.setText("警告: 缺少pywin32库，无法使用Word进行转换。请运行: pip install pywin32")
                self.dependency_label.setStyleSheet("color: #d32f2f; font-size: 10px;")
            except Exception as e:
                self.dependency_label.setText(f"Word检测错误: {str(e)}")
                self.dependency_label.setStyleSheet("color: #d32f2f; font-size: 10px;")
        else:
            self.dependency_label.setText("警告: 此程序仅在Windows系统上有效，因为需要使用Microsoft Word")
            self.dependency_label.setStyleSheet("color: #d32f2f; font-size: 10px;")
    
    def browse_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择PDF文件", "", "PDF文件 (*.pdf)"
        )
        if file_path:
            self.pdf_path_edit.setText(file_path)
            # 如果输出目录为空，则设置为PDF文件所在目录
            if not self.word_path_edit.text():
                self.word_path_edit.setText(os.path.dirname(file_path))
    
    def browse_word_dir(self):
        dir_path = QFileDialog.getExistingDirectory(
            self, "选择输出目录", ""
        )
        if dir_path:
            self.word_path_edit.setText(dir_path)
    
    def convert(self):
        pdf_path = self.pdf_path_edit.text()
        output_dir = self.word_path_edit.text()
        
        # 验证输入
        if not pdf_path:
            QMessageBox.warning(self, "警告", "请选择PDF文件")
            return
        
        if not os.path.exists(pdf_path):
            QMessageBox.warning(self, "警告", f"文件 '{pdf_path}' 不存在")
            return
        
        if not output_dir:
            QMessageBox.warning(self, "警告", "请选择输出目录")
            return
        
        if not os.path.exists(output_dir):
            QMessageBox.warning(self, "警告", f"目录 '{output_dir}' 不存在")
            return
        
        # 检查Word是否可用
        if platform.system() != "Windows":
            QMessageBox.critical(self, "错误", "此程序仅在Windows系统上有效，因为需要使用Microsoft Word")
            return
        
        try:
            import win32com.client
        except ImportError:
            QMessageBox.critical(self, "错误", "缺少pywin32库，无法使用Word进行转换。请运行: pip install pywin32")
            return
        
        # 生成输出Word文件路径
        pdf_filename = os.path.basename(pdf_path)
        word_filename = os.path.splitext(pdf_filename)[0] + '.docx'
        word_path = os.path.join(output_dir, word_filename)
        
        # 禁用按钮，显示进度
        self.convert_btn.setEnabled(False)
        self.status_label.setText("正在转换中...")
        self.progress_bar.setValue(10)  # 初始进度
        
        # 创建并启动转换线程
        self.conversion_thread = ConversionThread(pdf_path, word_path)
        self.conversion_thread.progress_signal.connect(self.update_progress)
        self.conversion_thread.finished_signal.connect(self.conversion_finished)
        self.conversion_thread.start()
    
    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)
    
    def conversion_finished(self, success, error_msg):
        # 重新启用按钮
        self.convert_btn.setEnabled(True)
        
        if success:
            self.progress_bar.setValue(100)
            self.status_label.setText("转换成功!")
            
            pdf_path = self.pdf_path_edit.text()
            output_dir = self.word_path_edit.text()
            pdf_filename = os.path.basename(pdf_path)
            word_filename = os.path.splitext(pdf_filename)[0] + '.docx'
            word_path = os.path.join(output_dir, word_filename)
            
            QMessageBox.information(
                self, 
                "成功", 
                f"PDF已成功转换为Word文档:\n{word_path}"
            )
        else:
            self.progress_bar.setValue(0)
            self.status_label.setText("转换失败")
            QMessageBox.critical(
                self, 
                "错误", 
                f"转换过程中发生错误:\n{error_msg}\n\n请确保您的系统上安装了Microsoft Word，并且可以打开PDF文件。"
            )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFToWordConverter()
    window.show()
    sys.exit(app.exec()) 