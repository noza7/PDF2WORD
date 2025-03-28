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
    superscript_fixed_signal = pyqtSignal(int)  # 新增信号，传递修复的上标数量
    paragraph_fixed_signal = pyqtSignal(int)  # 新增信号，传递修复的段落数量
    
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
                    superscript_count = self.fix_superscript_numbers(self.word_path)
                    self.superscript_fixed_signal.emit(superscript_count)
                except Exception as e:
                    print(f"修复上标数字失败: {str(e)}")
                
                # 修复编号段落问题
                try:
                    paragraph_count = self.fix_numbered_paragraphs(self.word_path)
                    self.paragraph_fixed_signal.emit(paragraph_count)
                except Exception as e:
                    print(f"修复编号段落失败: {str(e)}")
                
                self.finished_signal.emit(True, "")
                return
            else:
                raise Exception("转换后的文件无效或为空")
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.finished_signal.emit(False, str(e))
    
    def convert_with_word(self):
        """使用Word打开PDF并另存为Word文档，完全隐藏弹窗"""
        if platform.system() != "Windows":
            raise Exception("Word COM自动化仅在Windows系统上可用")
        
        try:
            import win32com.client
            import pythoncom
            import win32gui
            import win32con
            import win32process
            import win32api
            import time
            import threading
            import os
            
            # 初始化COM
            pythoncom.CoInitialize()
            
            # 创建Word应用实例
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            # 禁用所有警告和弹窗
            word.DisplayAlerts = False  # 禁用大多数警告弹窗
            word.Options.ConfirmConversions = False  # 禁用格式转换确认
            
            # 设置Word为静默模式
            try:
                word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            except:
                pass  # 如果此属性不可用，忽略错误
            
            # 发送进度信号
            self.progress_signal.emit(20)
            
            # 创建一个函数来处理弹窗
            def handle_dialogs():
                try:
                    # 循环检查并处理弹窗
                    for _ in range(30):  # 增加尝试次数
                        # 查找常见的Word对话框标题
                        dialog_titles = ["Microsoft Word", "警告", "Warning", "转换", "Conversion", 
                                        "安全警告", "Security Warning", "兼容性检查器", "Compatibility Checker"]
                        
                        for title in dialog_titles:
                            # 查找对话框
                            hwnd = win32gui.FindWindow(None, title)
                            if hwnd != 0 and win32gui.IsWindowVisible(hwnd):
                                print(f"找到对话框: {title}")
                                # 尝试点击"确定"或"是"按钮
                                buttons = ["确定", "OK", "是", "Yes", "继续", "Continue"]
                                for button_text in buttons:
                                    try:
                                        button = win32gui.FindWindowEx(hwnd, 0, "Button", button_text)
                                        if button != 0:
                                            win32gui.PostMessage(button, win32con.BM_CLICK, 0, 0)
                                            print(f"点击了按钮: {button_text}")
                                            break
                                    except:
                                        pass
                                
                                # 如果找不到按钮，尝试发送回车键
                                if True:  # 总是尝试发送回车键作为备选方案
                                    try:
                                        win32gui.SetForegroundWindow(hwnd)
                                        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                                        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
                                        print(f"发送回车键到对话框: {title}")
                                    except:
                                        pass
                        
                        # 查找任何可能的对话框（通过类名）
                        try:
                            def enum_windows_callback(hwnd, _):
                                if win32gui.IsWindowVisible(hwnd):
                                    classname = win32gui.GetClassName(hwnd)
                                    if classname == "#32770":  # 标准对话框类名
                                        title = win32gui.GetWindowText(hwnd)
                                        print(f"找到对话框: {title} (类名: {classname})")
                                        try:
                                            win32gui.SetForegroundWindow(hwnd)
                                            win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                                            win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
                                        except:
                                            pass
                                return True
                            
                            win32gui.EnumWindows(enum_windows_callback, None)
                        except:
                            pass
                        
                        time.sleep(0.5)
                except Exception as e:
                    print(f"处理对话框线程出错: {str(e)}")
            
            # 启动处理对话框的线程
            dialog_thread = threading.Thread(target=handle_dialogs)
            dialog_thread.daemon = True
            dialog_thread.start()
            
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
            import traceback
            traceback.print_exc()
            raise Exception(f"使用Word转换失败: {str(e)}")
    
    def fix_superscript_numbers(self, docx_path):
        """修复Word文档中的所有上标，使其格式与相邻字符一致"""
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
            
            # 记录处理的上标数量
            superscript_count = 0
            
            # 使用查找替换功能处理所有上标
            word.Selection.Find.ClearFormatting()
            word.Selection.Find.Font.Superscript = True  # 查找上标格式
            word.Selection.Find.Text = ""  # 空字符串表示任何文本
            
            # 循环查找所有上标并处理
            found = word.Selection.Find.Execute(
                FindText="",
                MatchCase=False,
                MatchWholeWord=False,
                MatchWildcards=False,
                MatchSoundsLike=False,
                MatchAllWordForms=False,
                Forward=True,
                Wrap=1,  # wdFindContinue
                Format=True
            )
            
            while found:
                # 记录找到的上标内容
                superscript_text = word.Selection.Text
                
                # 获取当前选择的范围
                current_range = word.Selection.Range
                
                # 尝试获取前一个字符的格式
                try:
                    # 创建一个新的范围，位于当前选择之前
                    prev_range = doc.Range(current_range.Start - 1, current_range.Start)
                    
                    # 如果前一个字符存在且不是空格或特殊字符
                    if prev_range.Text and prev_range.Text.strip():
                        # 移除上标格式
                        word.Selection.Font.Superscript = False
                        
                        # 应用前一个字符的格式
                        word.Selection.Font.Name = prev_range.Font.Name
                        word.Selection.Font.Size = prev_range.Font.Size
                        word.Selection.Font.Bold = prev_range.Font.Bold
                        word.Selection.Font.Italic = prev_range.Font.Italic
                        word.Selection.Font.Underline = prev_range.Font.Underline
                        word.Selection.Font.Color = prev_range.Font.Color
                    else:
                        # 如果前一个字符不存在或是空格，尝试获取后一个字符的格式
                        next_range = doc.Range(current_range.End, current_range.End + 1)
                        
                        if next_range.Text and next_range.Text.strip():
                            # 移除上标格式
                            word.Selection.Font.Superscript = False
                            
                            # 应用后一个字符的格式
                            word.Selection.Font.Name = next_range.Font.Name
                            word.Selection.Font.Size = next_range.Font.Size
                            word.Selection.Font.Bold = next_range.Font.Bold
                            word.Selection.Font.Italic = next_range.Font.Italic
                            word.Selection.Font.Underline = next_range.Font.Underline
                            word.Selection.Font.Color = next_range.Font.Color
                        else:
                            # 如果前后都没有可参考的字符，只移除上标格式
                            word.Selection.Font.Superscript = False
                except:
                    # 如果获取相邻字符格式失败，只移除上标格式
                    word.Selection.Font.Superscript = False
                
                # 计数
                superscript_count += 1
                
                # 继续查找下一个
                found = word.Selection.Find.Execute(
                    FindText="",
                    MatchCase=False,
                    MatchWholeWord=False,
                    MatchWildcards=False,
                    MatchSoundsLike=False,
                    MatchAllWordForms=False,
                    Forward=True,
                    Wrap=1,  # wdFindContinue
                    Format=True
                )
            
            # 保存文档
            doc.Save()
            
            # 关闭文档
            doc.Close()
            
            # 打印处理结果
            print(f"已修复 {superscript_count} 处上标格式")
            
            # 返回处理的上标数量
            return superscript_count
            
        finally:
            # 退出Word应用
            word.Quit()
            
            # 释放COM资源
            pythoncom.CoUninitialize()

    def fix_numbered_paragraphs(self, docx_path):
        """识别并修复编号段落，确保每个编号项都是单独的段落并保持格式一致，去掉首行缩进"""
        import win32com.client
        import pythoncom
        import re
        
        # 初始化COM
        pythoncom.CoInitialize()
        
        try:
            # 创建Word应用实例
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            # 打开Word文档
            doc = word.Documents.Open(docx_path)
            
            # 记录处理的段落数量
            fixed_count = 0
            
            # 定义编号模式的正则表达式
            # 匹配中文数字编号：一、二、三、...（包括前面可能有句号的情况）
            chinese_number_pattern = r'([。．.]\s*)?[一二三四五六七八九十百千万亿]+[、，,．.．:：]'
            # 匹配阿拉伯数字编号：1. 2. 3. ...（包括前面可能有句号的情况）
            arabic_number_pattern = r'([。．.]\s*)?[0-9]+[、，,．.．:：]'
            # 匹配带括号的编号：(1) （2） ...（包括前面可能有句号的情况）
            parenthesis_pattern = r'([。．.]\s*)?[(（][0-9]+[)）]'
            
            # 组合所有模式
            number_patterns = [chinese_number_pattern, arabic_number_pattern, parenthesis_pattern]
            
            # 遍历文档中的所有段落
            for i in range(1, doc.Paragraphs.Count + 1):
                para = doc.Paragraphs(i)
                text = para.Range.Text.strip()
                
                # 如果段落为空，跳过
                if not text:
                    continue
                
                # 保存当前段落的格式信息，以便应用到新段落
                para.Range.Select()
                current_format = {
                    'LeftIndent': word.Selection.ParagraphFormat.LeftIndent,
                    'RightIndent': word.Selection.ParagraphFormat.RightIndent,
                    'FirstLineIndent': word.Selection.ParagraphFormat.FirstLineIndent,
                    'Alignment': word.Selection.ParagraphFormat.Alignment,
                    'LineSpacing': word.Selection.ParagraphFormat.LineSpacing,
                    'SpaceBefore': word.Selection.ParagraphFormat.SpaceBefore,
                    'SpaceAfter': word.Selection.ParagraphFormat.SpaceAfter,
                    'Font': {
                        'Name': word.Selection.Font.Name,
                        'Size': word.Selection.Font.Size,
                        'Bold': word.Selection.Font.Bold,
                        'Italic': word.Selection.Font.Italic,
                        'Underline': word.Selection.Font.Underline,
                        'Color': word.Selection.Font.Color
                    }
                }
                
                # 检查段落是否以编号开头，如果是，去掉首行缩进
                is_numbered_paragraph = False
                for pattern in number_patterns:
                    if re.match(pattern, text):
                        is_numbered_paragraph = True
                        break
                
                if is_numbered_paragraph:
                    # 选择段落
                    para.Range.Select()
                    
                    # 保持左缩进不变，但去掉首行缩进
                    # 计算字符宽度（假设平均字符宽度为字体大小的一半）
                    char_width = word.Selection.Font.Size / 2
                    
                    # 去掉约2个字符的首行缩进
                    word.Selection.ParagraphFormat.FirstLineIndent = 0
                    
                    # 计数
                    fixed_count += 1
                
                # 检查段落中是否包含编号模式
                for pattern in number_patterns:
                    # 在段落中查找编号模式
                    matches = list(re.finditer(pattern, text))
                    
                    # 如果没有匹配，继续下一个模式
                    if not matches:
                        continue
                    
                    # 跳过第一个匹配（如果它在段落开头）
                    start_index = 0
                    if matches[0].start() <= 2:  # 允许段落开头有少量空格
                        start_index = 1
                    
                    # 处理其余匹配
                    for match_index in range(start_index, len(matches)):
                        match = matches[match_index]
                        
                        # 检查匹配前是否有句号
                        has_period_before = bool(re.search(r'[。．.]', match.group(0)))
                        
                        # 找到段落中间的编号，需要在此处分段
                        position = para.Range.Start + match.start()
                        
                        # 如果编号前有句号，在句号后分段
                        if has_period_before:
                            # 找到句号的位置
                            period_pos = match.group(0).find('。')
                            if period_pos == -1:
                                period_pos = match.group(0).find('．')
                            if period_pos == -1:
                                period_pos = match.group(0).find('.')
                            
                            if period_pos != -1:
                                # 调整分段位置到句号后
                                position = para.Range.Start + match.start() + period_pos + 1
                        
                        # 创建一个范围到分段位置
                        split_range = doc.Range(position, position)
                        
                        # 插入段落标记（回车符）
                        split_range.InsertParagraphAfter()
                        
                        # 获取新插入的段落
                        new_para_index = i + 1  # 新段落的索引
                        if new_para_index <= doc.Paragraphs.Count:
                            new_para = doc.Paragraphs(new_para_index)
                            
                            # 选择新段落
                            new_para.Range.Select()
                            
                            # 应用与原段落相同的格式
                            word.Selection.ParagraphFormat.LeftIndent = current_format['LeftIndent']
                            word.Selection.ParagraphFormat.RightIndent = current_format['RightIndent']
                            # 去掉首行缩进
                            word.Selection.ParagraphFormat.FirstLineIndent = 0
                            word.Selection.ParagraphFormat.Alignment = current_format['Alignment']
                            word.Selection.ParagraphFormat.LineSpacing = current_format['LineSpacing']
                            word.Selection.ParagraphFormat.SpaceBefore = current_format['SpaceBefore']
                            word.Selection.ParagraphFormat.SpaceAfter = current_format['SpaceAfter']
                            
                            # 应用字体格式
                            word.Selection.Font.Name = current_format['Font']['Name']
                            word.Selection.Font.Size = current_format['Font']['Size']
                            word.Selection.Font.Bold = current_format['Font']['Bold']
                            word.Selection.Font.Italic = current_format['Font']['Italic']
                            word.Selection.Font.Underline = current_format['Font']['Underline']
                            word.Selection.Font.Color = current_format['Font']['Color']
                        
                        # 计数
                        fixed_count += 1
                        
                        # 由于插入了新段落，需要重新获取当前段落的文本
                        text = para.Range.Text.strip()
                        
                        # 重新开始查找，因为文本已经改变
                        break
            
            # 保存文档
            doc.Save()
            
            # 关闭文档
            doc.Close()
            
            # 打印处理结果
            print(f"已修复 {fixed_count} 处编号段落")
            
            # 返回处理的段落数量
            return fixed_count
            
        finally:
            # 退出Word应用
            word.Quit()
            
            # 释放COM资源
            pythoncom.CoUninitialize()


class PDFToWordConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Noza出品")
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
        self.conversion_thread.superscript_fixed_signal.connect(self.update_superscript_info)
        self.conversion_thread.paragraph_fixed_signal.connect(self.update_paragraph_info)
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
    
    def update_superscript_info(self, count):
        """更新上标修复信息"""
        if count > 0:
            self.status_label.setText(f"转换成功! 已修复 {count} 处上标格式")
        else:
            self.status_label.setText("转换成功! 未发现上标格式")

    def update_paragraph_info(self, count):
        """更新段落修复信息"""
        if count > 0:
            current_text = self.status_label.text()
            self.status_label.setText(f"{current_text}, 已修复 {count} 处编号段落")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFToWordConverter()
    window.show()
    sys.exit(app.exec()) 