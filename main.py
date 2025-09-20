import sys
from PyQt5.QtWidgets import QApplication
from ui.main_window import MainWindow

def main():
    """主函数"""
    app = QApplication(sys.argv)
    
    # 创建主窗口
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()