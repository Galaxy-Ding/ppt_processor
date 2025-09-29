import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QFileDialog,
                             QPushButton, QLabel, QTextEdit, QCheckBox,
                             QComboBox, QMessageBox, QLineEdit)
from PyQt5 import uic
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPainter, QColor
from PyQt5.QtGui import QPainter, QColor
from utils.logger import LoggerFactory, LOG_LEVELS, LEVEL_MAP
from config.loader import ConfigLoader
from PyQt5.QtCore import QThread, pyqtSignal
from extractors.extrator_发包规范 import ExtractorExcel
# 初始化处理器
from core.packing_file_发包规范 import PackingFileProcessor

# 后台处理线程
class ProcessThread(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    log_signal = pyqtSignal(str)
    auto_info_signal = pyqtSignal(str, str)  # 新增，传递自动匹配的名称和类型
    def __init__(self, processor, selected_dir, data_list=None, manual_proj_name_value=None, manual_proj_action_value=None):
        super().__init__()
        self.processor = processor
        self.selected_dir = selected_dir
        self.data_list = data_list
        self.manual_proj_name_value = manual_proj_name_value
        self.manual_proj_action_value = manual_proj_action_value
    def emit_log(self, msg):
        self.log_signal.emit(msg)
    def run(self):
        try:
            self.processor.set_log_callback(self.emit_log)
            results = self.processor.process_generate_reports(
                self.selected_dir, self.data_list,
                self.manual_proj_name_value, self.manual_proj_action_value
            )
            # 假设只处理一个PPT，取第一个结果
            for pptx_path, output_path in results.items():
                # 这里可以从processor里获取匹配结果
                matched_name = self.processor.last_matched_name if hasattr(self.processor, "last_matched_name") else ""
                matched_action = self.processor.last_matched_action if hasattr(self.processor, "last_matched_action") else ""
                self.auto_info_signal.emit(matched_name, matched_action)
                break
            self.finished.emit(results)
        except Exception as e:
            self.error.emit(str(e))


class DemoMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # 加载UI
        uic.loadUi("ui/发包规范_window.ui", self)
        self.setWindowTitle("发包规范一键生成工具")

        # 初始化配置和日志
        self.configs = ConfigLoader()
        self.logger = LoggerFactory.create_logger("DemoUI")
        self.processor = None
        self.selected_dirs = []

        # 绑定控件
        self._bind_widgets()
        self._connect_signals()
        # 美化界面
        self.apply_style()
        self.processor = PackingFileProcessor(self.configs)
        # 新增：初始化手动输入的值
        self.manual_proj_name_value = ""
        # 设置 combobox 默认值为第0个
        self.manual_proj_action.setCurrentIndex(0)
        self.manual_proj_action_value = self.manual_proj_action.currentText()
        # 设置 read_project_status 为只读
        self.read_project_status.setEnabled(False)
        self.data_list = self._read_from_local_excel()

    def apply_style(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QWidget {
                font-family: 'Microsoft YaHei', Arial, sans-serif;
            }
            QPushButton#selectDirBtn {
                background-color: #4a86e8;
                color: white;
                border: none;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton#selectDirBtn:hover {
                background-color: #3a76d8;
            }
            QPushButton#selectDirBtn:pressed {
                background-color: #2a66c8;
            }
            QPushButton#generateBtn {
                background-color: #4caf50;
                color: white;
                border: none;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton#generateBtn:hover {
                background-color: #3d9a40;
            }
            QPushButton#generateBtn:pressed {
                background-color: #2e8b30;
            }
            QTextEdit {
                background-color: white;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 8px;
                font-family: 'Consolas', 'Courier New', monospace;
            }
            QComboBox {
                padding: 5px;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                background-color: white;
            }
            QLabel {
                color: #333333;
            }
        """)
        font = QFont("Microsoft YaHei", 9)
        QApplication.setFont(font)

    def setup_logger(self, log_config: dict = {}):
        """设置日志记录器"""
        if log_config:
            log_path = log_config.get('path', '../logs')
            os.makedirs(log_path, exist_ok=True)

            return LoggerFactory.create_logger(
                "PPTProcessorGUI",
                log_level=log_config.get('level', 'INFO'),
                log_file=os.path.join(log_path, 'ppt_processor.log'),
                retention_days=log_config.get('retention_days', 30)
            )


    def _bind_widgets(self):
        # 这些名称需与demo.ui中的objectName一致
        self.select_dir_btn = self.findChild(type(self.select_dir_btn), "select_dir_btn")
        self.selected_dir_label = self.findChild(type(self.selected_dir_label), "selected_dir_label")
        self.generate_btn = self.findChild(type(self.generate_btn), "generate_btn")
        self.log_level_combo = self.findChild(QComboBox, "log_level_combo")
        self.log_display = self.findChild(QTextEdit, "log_display")
        self.status_label = self.findChild(type(self.status_label), "status_label")
        self.manual_proj_name = self.findChild(QLineEdit, "manual_txt_proj_name")
        self.manual_proj_action = self.findChild(QComboBox, "manual_proj_action_combo")
        self.auto_proj_name = self.findChild(QLabel, "auto_proj_name_label")
        self.auto_proj_action = self.findChild(QLabel, "auto_proj_action_label")
        self.read_project_status = self.findChild(QCheckBox, "chBox_read_ProjectStatus") # 状态圆如用自定义控件可用

    def _connect_signals(self):
        self.select_dir_btn.clicked.connect(self.select_directories)
        self.generate_btn.clicked.connect(self.generate_output)
        self.log_level_combo.currentTextChanged.connect(self.change_log_level)
        # 新增：手动输入信号与槽函数绑定
        self.manual_proj_name.editingFinished.connect(self.on_manual_proj_name_changed)
        self.manual_proj_action.currentTextChanged.connect(self.on_manual_proj_action_changed)

    # 新增：手动输入槽函数
    def on_manual_proj_name_changed(self):
        value = self.manual_proj_name.text()
        self.manual_proj_name_value = value
        self.logger.debug(f"手动工程名字输入: {value}")
        self.append_log(f"手动工程名字输入: {value}")

    def on_manual_proj_action_changed(self, value):
        self.manual_proj_action_value = value
        self.logger.debug(f"手动工程类型选择: {value}")
        self.append_log(f"手动工程类型选择: {value}")

    def select_directories(self):
        """选择多个目录，显示所有已选目录"""
        default_dir = os.getcwd()
        if isinstance(self.configs.config, dict) and 'default_dirs' in self.configs.config:
            if isinstance(self.configs.config['default_dirs'], list) and len(self.configs.config['default_dirs']) > 0:
                default_dir = os.path.abspath(self.configs.config['default_dirs'][0])
                if not os.path.exists(default_dir):
                    default_dir = os.getcwd()
                    self.logger.warning(f"配置的默认目录不存在，使用当前目录: {default_dir}")

        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        folderName = QFileDialog.getExistingDirectory(self, "选择文件夹", default_dir, options=options)
        if folderName:
            if folderName not in self.selected_dirs:
                self.selected_dirs.append(folderName)
            # 更新label和log显示所有已选目录
            dirs_str = '\n'.join(self.selected_dirs)
            self.selected_dir_label.setText("已选目录数: {}".format(len(self.selected_dirs)))
            self.log_display.append(f"当前已选择的目录：\n{dirs_str}")
        else:
            if not self.selected_dirs:
                self.selected_dir_label.setText("未选择目录")

    def change_log_level(self, level):
        new_level = LEVEL_MAP[level]
        self.logger.setLevel(new_level)
        for logger in LoggerFactory._loggers.values():
            logger.setLevel(new_level)
            logger.level = LOG_LEVELS[level]
            for handler in logger.handlers:
                handler.setLevel(new_level)
        self.append_log(f"日志级别已切换为: {level}")
        self.processor.change_log_level(level)


    def generate_output(self):
        if not self.selected_dirs:
            QMessageBox.warning(self, "提示", "请先选择目录！")
            self.logger.warning("请先选择目录")
            return
        self.append_log("开始处理...")
        try:
            self.status_label.setText("处理中...")
            self.generate_btn.setEnabled(False)
            # 拷贝队列，避免原始列表被破坏
            self._pending_dirs = list(self.selected_dirs)
            self._all_results = {}
            self._process_next_dir()
        except Exception as e:
            self.append_log(f"处理异常: {e}")
            self.logger.error(f"处理异常: {e}")
            # QMessageBox.critical(self, "错误", str(e))
            self.status_label.setText("异常")

    def update_auto_info(self, name, action):
        self.auto_proj_name.setText(name)
        self.auto_proj_action.setText(action)

    def _process_next_dir(self):
        if not self._pending_dirs:
            # 全部处理完成
            self.generate_btn.setEnabled(True)
            self.status_label.setText("处理完成")
            self.logger.info("所有目录处理完成")
            # 可在此处汇总所有结果 self._all_results
            return
        item = self._pending_dirs.pop(0)
        if not os.path.exists(item):
            self.logger.error(f"选择的目录不存在: {item}")
            QMessageBox.critical(self, "错误", f"选择的目录不存在: {item}")
            self.generate_btn.setEnabled(True)
            self.status_label.setText("就绪")
            return
        # 启动后台线程
        self.process_thread = ProcessThread(self.processor, item, self.data_list,
                                            self.manual_proj_name_value, self.manual_proj_action_value)
        self.process_thread.finished.connect(self._on_single_process_finished)
        self.process_thread.error.connect(self._on_single_process_error)
        self.process_thread.log_signal.connect(self.append_log)
        self.process_thread.auto_info_signal.connect(self.update_auto_info)  # 新增
        self.process_thread.start()

    def _on_single_process_finished(self, results):
        # 合并结果
        if isinstance(results, dict):
            self._all_results.update(results)
            for dir_path, output_files in results.items():
                self.logger.info(f"目录 {dir_path} 处理完成，生成 {output_files}")
        self._process_next_dir()

    def _on_single_process_error(self, error_msg):
        self.logger.error(f"处理过程中出错: {error_msg}")
        self.log_display.append(f"[ERROR] 处理过程中出错: {error_msg}")
        scrollbar = self.log_display.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
        self.generate_btn.setEnabled(True)
        self.status_label.setText("处理失败")

    def append_log(self, msg):
        self.log_display.append(msg)
        scrollbar = self.log_display.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def _read_from_local_excel(self):
        """
        读取本地Excel，异常保护，日志根据等级显示和记录
        """
        try:
            extractor = ExtractorExcel(self.configs)
            data_list = extractor.extract()  # 返回 [dict1, dict2, ...]
            self.logger.info(f"成功读取Excel数据，共{len(data_list)}行")
            self.append_log(f"成功读取Excel数据，共{len(data_list)}行")
            self.read_project_status.setChecked(True)  # 读取成功，设为True
            return data_list
        except Exception as e:
            msg = f"读取Excel失败: {e}"
            if self.logger:
                self.logger.error(msg)
            self.append_log(msg)
            self.read_project_status.setChecked(False)  # 读取失败，设为False
            return []

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DemoMainWindow()
    window.show()
    sys.exit(app.exec_())