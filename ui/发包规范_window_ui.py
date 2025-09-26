from PyQt5.QtCore import QThread, pyqtSignal
import sys
import os
import re
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QLabel, QTextEdit, 
                            QFileDialog, QComboBox, QMessageBox, QLineEdit,
                             QSizePolicy)
from PyQt5.QtCore import Qt
from utils.logger import LoggerFactory
from config.loader import ConfigLoader
# 初始化处理器
from core.packing_file_发包规范 import PackingFileProcessor
from PyQt5.QtGui import QPainter, QColor




# 后台处理线程
class ProcessThread(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    log_signal = pyqtSignal(str)
    def __init__(self, processor, selected_dir):
        super().__init__()
        self.processor = processor
        self.selected_dir = selected_dir
    def emit_log(self, msg):
        self.log_signal.emit(msg)
    def run(self):
        try:
            self.processor.set_log_callback(self.emit_log)
            results = self.processor.process_generate_reports(self.selected_dir)
            self.finished.emit(results)
        except Exception as e:
            self.error.emit(str(e))

from PyQt5.QtGui import QPainter, QColor

class StatusCircle(QWidget):
    """绿色/红色状态圆"""
    def __init__(self, color="green", parent=None):
        super().__init__(parent)
        self._color = color
        self.setFixedSize(24, 24)

    def set_color(self, color):
        self._color = color
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        if self._color == "green":
            painter.setBrush(QColor(0, 200, 0))
        else:
            painter.setBrush(QColor(200, 0, 0))
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(0, 0, 24, 24)

class PPTProcessorGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("发包规范一键生成工具")
        self.setGeometry(100, 100, 800, 600)
        
        # 初始化配置
        self.configs = ConfigLoader()
        self.logger = self.setup_logger(self.configs.config["logs"])
        self.selected_dirs = []

        # 创建UI
        self.init_ui()

        self.processor = PackingFileProcessor(
            self.configs
        )

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

    def init_ui(self):
        """根据新UI布局调整"""
        main_widget = QWidget(self)
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(15, 15, 15, 15)

        # 顶部控制区域
        top_control_widget = QWidget()
        top_layout = QHBoxLayout(top_control_widget)
        top_layout.setContentsMargins(0, 0, 0, 0)

        # 选择目录按钮
        self.select_dir_btn = QPushButton("选择目录")
        self.select_dir_btn.setFixedSize(160, 80)
        self.select_dir_btn.setObjectName("selectDirBtn")
        self.select_dir_btn.setStyleSheet("font-size:22px; background:#4A90E2; color:white; border-radius:20px; padding:20px;")
        self.select_dir_btn.clicked.connect(self.select_directories)

        # 目录标签
        self.selected_dir_label = QLabel("未选择")
        self.selected_dir_label.setStyleSheet("font-size:18px; color:gray;")
        self.selected_dir_label.setFixedHeight(30)

        # 日志级别标签和下拉框
        log_level_label = QLabel("日志级别：")
        self.log_level_combo = QComboBox()
        self.log_level_combo.addItems(["INFO", "DEBUG", "WARNING", "ERROR"])
        self.log_level_combo.setCurrentText("INFO")
        self.log_level_combo.setFixedSize(120, 30)
        self.log_level_combo.currentTextChanged.connect(self.change_log_level)

        # 一键生成按钮
        self.generate_btn = QPushButton("一键生成")
        self.generate_btn.setFixedSize(160, 80)
        self.generate_btn.setObjectName("generateBtn")
        self.generate_btn.setStyleSheet("font-size:22px; background:#4CAF50; color:white; border-radius:20px; padding:20px;")
        self.generate_btn.clicked.connect(self.generate_output)

        # 顶部布局
        top_layout.addWidget(self.select_dir_btn)
        top_layout.addWidget(self.selected_dir_label)
        top_layout.addStretch()
        top_layout.addWidget(log_level_label)
        top_layout.addWidget(self.log_level_combo)
        top_layout.addWidget(self.generate_btn)

        # 工程信息区域（左右沾满）
        info_layout = QHBoxLayout()
        info_layout.setSpacing(30)
        info_layout.setContentsMargins(0, 0, 0, 0)

        # 左侧工程信息
        left_info_widget = QWidget()
        left_info_widget.setStyleSheet("QWidget { border:2px solid #aaa; border-radius:10px; }")
        left_info_layout = QVBoxLayout(left_info_widget)
        left_info_layout.setContentsMargins(15, 15, 15, 15)
        self.lbl_proj_name = QLabel("工程名字：")
        self.txt_proj_name = QLineEdit()
        self.lbl_proj_type = QLabel("工程类型")
        self.cmb_proj_type = QComboBox()
        self.cmb_proj_type.addItems(["小改造", "大改造", "新建"])
        left_info_layout.addWidget(self.lbl_proj_name)
        left_info_layout.addWidget(self.txt_proj_name)
        left_info_layout.addWidget(self.lbl_proj_type)
        left_info_layout.addWidget(self.cmb_proj_type)
        left_info_layout.addStretch()
        # 让左侧沾满
        left_info_widget.setMinimumWidth(0)
        left_info_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        # 右侧工程信息
        right_info_widget = QWidget()
        right_info_widget.setStyleSheet("QWidget { border:2px solid #aaa; border-radius:10px; }")
        right_info_layout = QVBoxLayout(right_info_widget)
        right_info_layout.setContentsMargins(15, 15, 15, 15)
        self.lbl_proj_name_r = QLabel("工程名字：")
        self.txt_proj_name_r = QLineEdit()
        self.lbl_proj_type_r = QLabel("工程类型")
        self.cmb_proj_type_r = QComboBox()
        self.cmb_proj_type_r.addItems(["小改造", "大改造", "新建"])
        right_info_layout.addWidget(self.lbl_proj_name_r)
        right_info_layout.addWidget(self.txt_proj_name_r)
        right_info_layout.addWidget(self.lbl_proj_type_r)
        right_info_layout.addWidget(self.cmb_proj_type_r)
        right_info_layout.addStretch()
        # 让右侧沾满
        right_info_widget.setMinimumWidth(0)
        right_info_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        # 状态圆，绝对定位到右侧框左上角
        status_circle_container = QWidget(right_info_widget)
        status_circle_container.setFixedSize(30, 30)
        self.status_circle = StatusCircle("green", status_circle_container)
        self.status_circle.move(3, 3)  # 微调到右侧框左上角

        # 用布局包裹右侧框和圆
        right_info_outer = QWidget()
        right_info_outer_layout = QVBoxLayout(right_info_outer)
        right_info_outer_layout.setContentsMargins(0, 0, 0, 0)
        right_info_outer_layout.setSpacing(0)
        right_info_outer_layout.addWidget(status_circle_container, alignment=Qt.AlignLeft | Qt.AlignTop)
        right_info_outer_layout.addWidget(right_info_widget)

        info_layout.addWidget(left_info_widget)
        info_layout.addWidget(right_info_outer)

        # 消息打印区域
        msg_label = QLabel("消息打印：")
        self.log_display = QTextEdit()
        self.log_display.setReadOnly(True)

        # 添加到主布局
        main_layout.addWidget(top_control_widget)
        main_layout.addLayout(info_layout)
        main_layout.addWidget(msg_label)
        main_layout.addWidget(self.log_display)

        # 状态栏
        status_layout = QHBoxLayout()
        status_layout.addStretch(1)
        self.status_label = QLabel("就绪")
        self.status_label.setStyleSheet("font-size:9.5pt; color:#888; margin-top:2px;")
        status_layout.addWidget(self.status_label)
        main_layout.addLayout(status_layout)

        self.apply_style()

    def set_status_circle(self, color: str):
        """设置状态圆颜色，color='green'或'red'"""
        self.status_circle.set_color(color)

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
        from PyQt5.QtGui import QFont
        font = QFont("Microsoft YaHei", 9)
        QApplication.setFont(font)

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
        """更改日志级别"""
        self.logger.setLevel(level)
        self.logger.info(f"日志级别已更改为: {level}")
        self.processor.change_log_level(level)

    def generate_output(self):
        """一键生成输出（顺序队列处理，防卡顿）"""
        if not self.selected_dirs:
            self.logger.warning("请先选择目录")
            return
        self.logger.info("开始处理...")
        self.status_label.setText("处理中...")
        self.generate_btn.setEnabled(False)
        # 拷贝队列，避免原始列表被破坏
        self._pending_dirs = list(self.selected_dirs)
        self._all_results = {}
        self._process_next_dir()

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
        self.process_thread = ProcessThread(self.processor, item)
        self.process_thread.finished.connect(self._on_single_process_finished)
        self.process_thread.error.connect(self._on_single_process_error)
        self.process_thread.log_signal.connect(self.append_log)
        self.process_thread.start()

    def _on_single_process_finished(self, results):
        # 合并结果
        if isinstance(results, dict):
            self._all_results.update(results)
            for dir_path, output_files in results.items():
                self.logger.info(f"目录 {dir_path} 处理完成，生成 {len(output_files)} 个文件")
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


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PPTProcessorGUI()
    window.show()
    sys.exit(app.exec_())