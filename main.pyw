import sys
import datetime
import requests
import winsound
import json
import winreg
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QHBoxLayout, QPushButton, QVBoxLayout,
    QStackedWidget, QListWidget, QListWidgetItem, QLineEdit, QTextEdit,
    QFrame, QSizePolicy, QGridLayout, QProgressBar, QInputDialog,
    QMessageBox, QSpinBox, QTableWidgetItem, QTabWidget, QTableWidget, QButtonGroup,
    QRadioButton, QCheckBox, QComboBox
)
from PyQt5.QtCore import Qt, QTimer, QSize, QThread, pyqtSignal, QPropertyAnimation
from PyQt5.QtGui import QFont, QColor, QPalette, QPainter, QBrush, QIcon
import random
import yaml
import os
import keyboard  # 用于模拟按键
import subprocess
import tendo.singleton
import base64
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.backends import default_backend

# 防止程序多开
single = tendo.singleton.SingleInstance()

class UpdateChecker(QThread):
    update_found = pyqtSignal(str, str)
    no_update = pyqtSignal()
    error_occurred = pyqtSignal(str)

    def __init__(self, current_version):
        super().__init__()
        self.current_version = current_version

    def run(self):
        try:
            # 获取最新版本号
            version_url = "http://spj2025.top:19540/apps/version/pyclasstool.txt"
            response = requests.get(version_url)
            if response.status_code != 200:
                self.error_occurred.emit(f"无法获取版本信息: HTTP {response.status_code}")
                return
            
            # 使用GBK编码解码响应内容
            latest_version = response.content.decode('gbk').strip()
            
            # 检查是否有新版本
            if latest_version != self.current_version:
                # 获取更新日志
                changelog_url = "http://spj2025.top:19540/apps/version/changelog/pyclasstool.txt"
                response = requests.get(changelog_url)
                if response.status_code != 200:
                    self.error_occurred.emit(f"无法获取更新日志: HTTP {response.status_code}")
                    return
                
                # 使用GBK编码解码更新日志
                changelog = response.content.decode('gbk')
                self.update_found.emit(latest_version, changelog)
            else:
                self.no_update.emit()
                
        except Exception as e:
            self.error_occurred.emit(f"检查更新时出错: {str(e)}")

class SettingsWindow(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.setWindowTitle("刷屏君 课堂工具 v1.0.1")
        self.setWindowIcon(QIcon("icon.ico"))  # 如果有图标文件的话
        self.setGeometry(100, 100, 800, 600)
        self.setMinimumSize(800, 600)
        
        # 强制设置页面使用浅色模式
        self.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                color: #000000;
            }
            QListWidget {
                background-color: #f0f0f0;
                color: #000000;
            }
            QListWidget::item:selected {
                background-color: #e0e0e0;
            }
            QLabel {
                color: #000000;
            }
            QCheckBox {
                color: #000000;
            }
            QPushButton {
                background-color: #e0e0e0;
                color: #000000;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 5px;
            }
            QPushButton:hover {
                background-color: #d0d0d0;
            }
            QTextEdit {
                background-color: #ffffff;
                color: #000000;
                border: 1px solid #d0d0d0;
            }
            QProgressBar {
                background-color: #ffffff;
                border: 1px solid #d0d0d0;
                border-radius: 2px;
            }
            QProgressBar::chunk {
                background-color: #0078d7;
            }
            QTabWidget::pane {
                border: 1px solid #d0d0d0;
                background-color: #f0f0f0;
            }
            QTabBar::tab {
                background-color: #f0f0f0;
                color: #000000;
                padding: 8px;
                border: 1px solid #d0d0d0;
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background-color: #ffffff;
            }
            QTableWidget {
                background-color: #ffffff;
                color: #000000;
                gridline-color: #d0d0d0;
            }
            QHeaderView::section {
                background-color: #e0e0e0;
                color: #000000;
                border: 1px solid #d0d0d0;
            }
            QLineEdit {
                background-color: #ffffff;
                color: #000000;
                border: 1px solid #d0d0d0;
            }
        """)
        
        # 创建主布局
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # 创建左侧导航栏
        self.nav_frame = QFrame()
        self.nav_frame.setFixedWidth(200)
        self.nav_frame.setStyleSheet("background-color: #f0f0f0;")
        nav_layout = QVBoxLayout(self.nav_frame)
        nav_layout.setContentsMargins(0, 10, 0, 10)
        nav_layout.setSpacing(0)
        
        # 导航列表
        self.nav_list = QListWidget()
        self.nav_list.setStyleSheet("""
            QListWidget {
                border: none;
                background-color: transparent;
            }
            QListWidget::item {
                padding: 10px 15px;
                border-bottom: 1px solid #e0e0e0;
            }
            QListWidget::item:selected {
                background-color: #e6e6e6;
                color: #000;
            }
        """)
        self.nav_list.setFixedWidth(200)
        
        # 添加导航项
        nav_items = ["临时改课", "考试功能", "检查更新", "电源管理", "程序管理"]
        for item in nav_items:
            list_item = QListWidgetItem(item)
            list_item.setSizeHint(QSize(200, 40))
            self.nav_list.addItem(list_item)
        
        # 添加设置项
        settings_item = QListWidgetItem("设置")
        settings_item.setSizeHint(QSize(200, 40))
        self.nav_list.addItem(settings_item)
        
        nav_layout.addWidget(self.nav_list)
        nav_layout.addStretch()
        
        # 创建右侧内容区域
        self.content_area = QStackedWidget()
        
        # 添加临时改课页面
        self.temp_course_page = self.create_temp_course_page()
        self.content_area.addWidget(self.temp_course_page)
        
        # 添加考试功能页面
        self.exam_tab = self.create_exam_setting_tab()
        self.content_area.addWidget(self.exam_tab)
        
        # 添加检查更新页面
        self.update_page = self.create_update_page()
        self.content_area.addWidget(self.update_page)
        
        # 添加电源管理页面
        self.power_page = self.create_power_management_page()
        self.content_area.addWidget(self.power_page)
        
        # 添加程序管理页面
        self.app_page = self.create_app_management_page()
        self.content_area.addWidget(self.app_page)
        
        # 添加设置页面
        self.settings_page = self.create_settings_page()
        self.content_area.addWidget(self.settings_page)
        
        # 连接导航选择事件
        self.nav_list.currentRowChanged.connect(self.content_area.setCurrentIndex)
        self.nav_list.currentRowChanged.connect(self.on_nav_changed)  # 添加导航变化处理
        
        # 添加到主布局
        main_layout.addWidget(self.nav_frame)
        main_layout.addWidget(self.content_area)
        
        # 设置初始选择 - 默认显示临时改课页面
        self.nav_list.setCurrentRow(0)
    
    def on_nav_changed(self, index):
        """导航变化时处理"""
        # 如果切换到检查更新页面
        if index == 2:  # 检查更新页面索引
            self.start_update_check()
        
        # 如果切换到设置页面
        elif index == 5:  # 设置页面索引
            self.verify_settings_access()
    
    def verify_settings_access(self):
        """验证设置页面访问权限"""
        # 检查是否设置了密码
        if 'password' in self.main_window.config:
            # 弹出密码输入对话框
            password, ok = QInputDialog.getText(
                self, 
                "密码验证", 
                "请输入管理密码以访问设置:", 
                QLineEdit.Password
            )
            
            if not ok:
                # 用户取消，返回上一个页面
                self.nav_list.setCurrentRow(0)  # 返回临时改课页面
                return
                
            # 验证密码
            if not self.main_window.verify_password(password):
                QMessageBox.critical(self, "密码错误", "输入的管理密码不正确")
                self.nav_list.setCurrentRow(0)  # 返回临时改课页面
    
    def create_temp_course_page(self):
        """创建临时改课页面"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # 标题
        title = QLabel("临时改课")
        title_font = QFont("Microsoft YaHei UI", 18, QFont.Bold)
        title.setFont(title_font)
        layout.addWidget(title)
        
        # 分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("background-color: #e0e0e0;")
        layout.addWidget(separator)
        
        # 快速选择区域
        quick_select_label = QLabel("快速选择:")
        quick_select_label.setFont(QFont("Microsoft YaHei UI", 12))
        layout.addWidget(quick_select_label)
        
        # 星期按钮
        weekdays = ["星期一", "星期二", "星期三", "星期四", "星期五"]
        button_layout = QGridLayout()
        button_layout.setHorizontalSpacing(10)
        button_layout.setVerticalSpacing(10)
        
        for i, day in enumerate(weekdays):
            btn = QPushButton(day)
            btn.setFixedHeight(40)
            btn.setFont(QFont("Microsoft YaHei UI", 12))
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #f0f0f0;
                    border: 1px solid #d0d0d0;
                    border-radius: 4px;
                }
                QPushButton:hover {
                    background-color: #e0e0e0;
                }
            """)
            # 修复闭包问题：使用默认参数捕获当前day值
            btn.clicked.connect(lambda _, d=day: self.select_day_courses(d))
            button_layout.addWidget(btn, i // 3, i % 3)
        
        layout.addLayout(button_layout)
        
        # 自定义课表区域
        custom_label = QLabel("自定义课表:")
        custom_label.setFont(QFont("Microsoft YaHei UI", 12))
        layout.addWidget(custom_label)
        
        # 课表输入框
        self.course_input = QTextEdit()
        self.course_input.setFont(QFont("Microsoft YaHei UI", 12))
        self.course_input.setPlaceholderText("在此输入自定义课表，课程之间用空格分隔")
        self.course_input.setStyleSheet("""
            QTextEdit {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 8px;
            }
        """)
        self.course_input.setFixedHeight(100)
        
        # 添加此行：在打开页面时自动填充当前课表内容
        self.course_input.setPlainText(self.main_window.schedule_label.text())
        
        layout.addWidget(self.course_input)
        
        # 应用按钮
        apply_btn = QPushButton("应用自定义课表")
        apply_btn.setFixedHeight(40)
        apply_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        apply_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d7;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0066b4;
            }
        """)
        apply_btn.clicked.connect(self.apply_custom_courses)
        layout.addWidget(apply_btn)
        
        # 添加伸缩空间
        layout.addStretch()
        
        return page
    
    def create_exam_setting_tab(self):
        """创建考试设置页面"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)
        
        # 标题
        title = QLabel("考试功能")
        title_font = QFont("Microsoft YaHei UI", 18, QFont.Bold)
        title.setFont(title_font)
        layout.addWidget(title)
        
        # 分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("background-color: #e0e0e0;")
        layout.addWidget(separator)
        
        # 说明标签
        info_label = QLabel("考试功能由独立的考试助手程序处理。请在此设置考试信息，保存后考试助手会自动处理。")
        info_label.setFont(QFont("Microsoft YaHei UI", 10))
        layout.addWidget(info_label)
        
        # 创建考试表格
        self.exam_table = QTableWidget()
        self.exam_table.setColumnCount(5)  # 科目、开始时间、结束时间、关机、操作
        self.exam_table.setHorizontalHeaderLabels(["考试科目", "开始时间", "结束时间", "考完关机", "操作"])
        self.exam_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
            }
            QHeaderView::section {
                background-color: #e0e0e0;
                padding: 4px;
                border: 1px solid #d0d0d0;
            }
        """)
        
        # 设置表格属性
        self.exam_table.horizontalHeader().setStretchLastSection(True)
        self.exam_table.setEditTriggers(QTableWidget.AllEditTriggers)
        self.exam_table.setSelectionBehavior(QTableWidget.SelectRows)
        
        # 加载现有考试
        self.load_exams()
        
        layout.addWidget(self.exam_table, 1)
        
        # 添加考试区域
        add_layout = QGridLayout()
        add_layout.setHorizontalSpacing(10)
        add_layout.setVerticalSpacing(10)
        
        # 科目输入框
        subject_label = QLabel("科目:")
        subject_label.setFont(QFont("Microsoft YaHei UI", 12))
        add_layout.addWidget(subject_label, 0, 0)
        
        self.subject_input = QLineEdit()
        self.subject_input.setFont(QFont("Microsoft YaHei UI", 12))
        self.subject_input.setPlaceholderText("输入考试科目")
        self.subject_input.setStyleSheet("""
            QLineEdit {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 5px;
            }
        """)
        add_layout.addWidget(self.subject_input, 0, 1, 1, 2)  # 跨两列
        
        # 开始时间输入框
        start_label = QLabel("开始时间:")
        start_label.setFont(QFont("Microsoft YaHei UI", 12))
        add_layout.addWidget(start_label, 1, 0)
        
        # 开始时间输入框和按钮容器
        start_container = QHBoxLayout()
        start_container.setSpacing(5)
        
        self.start_input = QLineEdit()
        self.start_input.setFont(QFont("Microsoft YaHei UI", 12))
        self.start_input.setPlaceholderText("格式: YYYY/MM/dd HH:mm")
        self.start_input.setStyleSheet("""
            QLineEdit {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 5px;
            }
        """)
        start_container.addWidget(self.start_input)
        
        # 获取当前时间按钮（开始时间）
        start_now_btn = QPushButton("当前时间")
        start_now_btn.setFixedWidth(80)
        start_now_btn.setFont(QFont("Microsoft YaHei UI", 10))
        start_now_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 5px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        start_now_btn.clicked.connect(lambda: self.copy_current_time("start"))
        start_container.addWidget(start_now_btn)
        
        add_layout.addLayout(start_container, 1, 1)
        
        # 结束时间输入框
        end_label = QLabel("结束时间:")
        end_label.setFont(QFont("Microsoft YaHei UI", 12))
        add_layout.addWidget(end_label, 2, 0)
        
        # 结束时间输入框和按钮容器
        end_container = QHBoxLayout()
        end_container.setSpacing(5)
        
        self.end_input = QLineEdit()
        self.end_input.setFont(QFont("Microsoft YaHei UI", 12))
        self.end_input.setPlaceholderText("格式: YYYY/MM/dd HH:mm")
        self.end_input.setStyleSheet("""
            QLineEdit {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 5px;
            }
        """)
        end_container.addWidget(self.end_input)
        
        # 获取当前时间按钮（结束时间）
        end_now_btn = QPushButton("当前时间")
        end_now_btn.setFixedWidth(80)
        end_now_btn.setFont(QFont("Microsoft YaHei UI", 10))
        end_now_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 5px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        end_now_btn.clicked.connect(lambda: self.copy_current_time("end"))
        end_container.addWidget(end_now_btn)
        
        add_layout.addLayout(end_container, 2, 1)
        
        # 关机复选框
        self.shutdown_check = QCheckBox("考完关机")
        self.shutdown_check.setFont(QFont("Microsoft YaHei UI", 12))
        add_layout.addWidget(self.shutdown_check, 3, 1)
        
        # 添加按钮
        add_btn = QPushButton("添加考试")
        add_btn.setFixedHeight(40)
        add_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        add_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d7;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0066b4;
            }
        """)
        add_btn.clicked.connect(self.add_exam)
        add_layout.addWidget(add_btn, 4, 1)
        
        layout.addLayout(add_layout)
        
        # 保存按钮
        save_btn = QPushButton("保存考试设置")
        save_btn.setFixedHeight(40)
        save_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        save_btn.setStyleSheet("""
            QPushButton {
                background-color: #107c10;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0e6e0e;
            }
        """)
        save_btn.clicked.connect(self.save_exams)
        layout.addWidget(save_btn)
        
        return tab

    def create_update_page(self):
        """创建检查更新页面（添加自动更新设置）"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # 标题
        title = QLabel("检查更新")
        title_font = QFont("Microsoft YaHei UI", 18, QFont.Bold)
        title.setFont(title_font)
        layout.addWidget(title)
        
        # 分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("background-color: #e0e0e0;")
        layout.addWidget(separator)
        
        # 自动更新设置
        auto_update_layout = QHBoxLayout()
        auto_update_layout.setSpacing(10)
        
        # 自动更新复选框
        self.auto_update_check = QCheckBox("启用自动更新")
        self.auto_update_check.setFont(QFont("Microsoft YaHei UI", 12))
        self.auto_update_check.stateChanged.connect(self.toggle_auto_update)
        auto_update_layout.addWidget(self.auto_update_check)
        
        # 加载当前设置
        auto_update_enabled = self.main_window.config.get('autoUpdate', False)
        self.auto_update_check.setChecked(auto_update_enabled)
        
        # 添加伸缩空间
        auto_update_layout.addStretch()
        
        layout.addLayout(auto_update_layout)
        
        # 状态标签
        self.update_status = QLabel("正在检查更新.")
        self.update_status.setFont(QFont("Microsoft YaHei UI", 14))
        self.update_status.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.update_status)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # 不确定模式
        self.progress_bar.setTextVisible(False)
        layout.addWidget(self.progress_bar)
        
        # 更新结果容器
        self.update_result_container = QWidget()
        self.update_result_container.setVisible(False)
        result_layout = QVBoxLayout(self.update_result_container)
        result_layout.setContentsMargins(0, 0, 0, 0)
        
        # 更新标题
        self.update_title = QLabel()
        self.update_title.setFont(QFont("Microsoft YaHei UI", 16, QFont.Bold))
        self.update_title.setStyleSheet("color: #0078d7;")
        result_layout.addWidget(self.update_title)
        
        # 更新日志区域
        changelog_label = QLabel("更新内容:")
        changelog_label.setFont(QFont("Microsoft YaHei UI", 12))
        result_layout.addWidget(changelog_label)
        
        self.changelog_text = QTextEdit()
        self.changelog_text.setReadOnly(True)
        self.changelog_text.setFont(QFont("Microsoft YaHei UI", 11))
        self.changelog_text.setStyleSheet("""
            QTextEdit {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 8px;
                background-color: #f8f8f8;
            }
        """)
        result_layout.addWidget(self.changelog_text, 1)
        
        # 更新按钮
        self.update_btn = QPushButton("立即更新")
        self.update_btn.setFixedHeight(40)
        self.update_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        self.update_btn.setStyleSheet("""
            QPushButton {
                background-color: #107c10;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0e6e0e;
            }
        """)
        self.update_btn.clicked.connect(self.start_update)
        result_layout.addWidget(self.update_btn)
        
        layout.addWidget(self.update_result_container, 1)
        
        # 添加伸缩空间
        layout.addStretch()
        
        return page

    def toggle_auto_update(self, state):
        """切换自动更新设置"""
        auto_update = state == Qt.Checked
        self.main_window.config['autoUpdate'] = auto_update
        self.main_window.save_config()
    
    def create_power_management_page(self):
        """创建电源管理页面"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # 标题
        title = QLabel("电源管理")
        title_font = QFont("Microsoft YaHei UI", 18, QFont.Bold)
        title.setFont(title_font)
        layout.addWidget(title)
        
        # 分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("background-color: #e0e0e0;")
        layout.addWidget(separator)
        
        # 警告提示
        warning_label = QLabel("警告：这些操作将影响您的计算机状态，请谨慎操作！")
        warning_label.setFont(QFont("Microsoft YaHei UI", 12))
        warning_label.setStyleSheet("color: #ff0000;")
        layout.addWidget(warning_label)
        
        # 按钮布局
        button_layout = QGridLayout()
        button_layout.setHorizontalSpacing(15)
        button_layout.setVerticalSpacing(15)
        
        # 关机按钮
        shutdown_btn = QPushButton("关机")
        shutdown_btn.setFixedHeight(50)
        shutdown_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        shutdown_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        shutdown_btn.clicked.connect(self.shutdown_computer)
        button_layout.addWidget(shutdown_btn, 0, 0)
        
        # 重启按钮
        restart_btn = QPushButton("重启")
        restart_btn.setFixedHeight(50)
        restart_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        restart_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        restart_btn.clicked.connect(self.confirm_restart)
        button_layout.addWidget(restart_btn, 0, 1)
        
        # 注销用户按钮
        logout_btn = QPushButton("注销用户")
        logout_btn.setFixedHeight(50)
        logout_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        logout_btn.setStyleSheet("""
            QPushButton {
                background-color: #2ecc71;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
        """)
        logout_btn.clicked.connect(self.confirm_logout)
        button_layout.addWidget(logout_btn, 1, 0)
        
        # 强制关机按钮
        force_shutdown_btn = QPushButton("强制关机")
        force_shutdown_btn.setFixedHeight(50)
        force_shutdown_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        force_shutdown_btn.setStyleSheet("""
            QPushButton {
                background-color: #9b59b6;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #8e44ad;
            }
        """)
        force_shutdown_btn.clicked.connect(self.confirm_force_shutdown)
        button_layout.addWidget(force_shutdown_btn, 1, 1)
        
        layout.addLayout(button_layout)
        
        # 添加伸缩空间
        layout.addStretch()
        
        return page
    
    def create_settings_page(self):
        """创建设置页面（添加课表编辑、人员编辑、主题设置和管理密码）"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # 标题
        title = QLabel("设置")
        title_font = QFont("Microsoft YaHei UI", 18, QFont.Bold)
        title.setFont(title_font)
        layout.addWidget(title)
        
        # 分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("background-color: #e0e0e0;")
        layout.addWidget(separator)
        
        # 创建选项卡
        self.tab_widget = QTabWidget()
        layout.addWidget(self.tab_widget)
        
        # 添加课表编辑选项卡
        self.course_tab = self.create_course_edit_tab()
        self.tab_widget.addTab(self.course_tab, "课表编辑")
        
        # 添加人员编辑选项卡
        self.people_tab = self.create_people_edit_tab()
        self.tab_widget.addTab(self.people_tab, "人员编辑")

        # 添加分组管理选项卡
        self.group_tab = self.create_group_management_tab()
        self.tab_widget.addTab(self.group_tab, "分组管理")
        
        # 添加主题设置选项卡
        self.theme_tab = self.create_theme_setting_tab()
        self.tab_widget.addTab(self.theme_tab, "主题设置")
        
        # 添加管理密码选项卡
        self.password_tab = self.create_password_setting_tab()
        self.tab_widget.addTab(self.password_tab, "管理密码")

        # 添加自动关机选项卡
        self.shutdown_tab = self.create_auto_shutdown_tab()
        self.tab_widget.addTab(self.shutdown_tab, "自动关机")

        # 添加下课提醒选项卡
        self.reminder_tab = self.create_class_reminder_tab()
        self.tab_widget.addTab(self.reminder_tab, "下课提醒")
        
        # 添加伸缩空间
        layout.addStretch()
        
        return page
    
    def create_class_reminder_tab(self):
        """创建下课提醒选项卡"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)
        
        # 说明标签
        info_label = QLabel("设置下课提醒时间，格式如：8:20、9:30、10:40（多个时间用顿号分隔）")
        info_label.setFont(QFont("Microsoft YaHei UI", 10))
        layout.addWidget(info_label)
        
        # 提醒时间输入框
        self.reminder_input = QLineEdit()
        self.reminder_input.setFont(QFont("Microsoft YaHei UI", 12))
        self.reminder_input.setPlaceholderText("输入下课提醒时间")
        self.reminder_input.setStyleSheet("""
            QLineEdit {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 8px;
            }
        """)
        
        # 加载现有提醒时间
        reminder_times = self.main_window.config.get('class_reminders', "")
        self.reminder_input.setText(reminder_times)
        
        layout.addWidget(self.reminder_input)
        
        # 保存按钮
        save_btn = QPushButton("保存提醒设置")
        save_btn.setFixedHeight(40)
        save_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        save_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d7;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0066b4;
            }
        """)
        save_btn.clicked.connect(self.save_class_reminders)
        layout.addWidget(save_btn)
        
        # 状态标签
        self.reminder_status = QLabel("")
        self.reminder_status.setFont(QFont("Microsoft YaHei UI", 10))
        layout.addWidget(self.reminder_status)
        
        # 测试按钮
        test_btn = QPushButton("测试提醒效果")
        test_btn.setFixedHeight(40)
        test_btn.setFont(QFont("Microsoft YaHei UI", 12))
        test_btn.setStyleSheet("""
            QPushButton {
                background-color: #2ecc71;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
        """)
        test_btn.clicked.connect(self.test_class_reminder)
        layout.addWidget(test_btn)
        
        # 添加伸缩空间
        layout.addStretch()
        
        return tab
    
    def create_password_setting_tab(self):
        """创建管理密码设置选项卡"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)
        
        # 说明标签
        info_label = QLabel("设置管理密码，用于退出程序和访问设置页面时验证身份")
        info_label.setFont(QFont("Microsoft YaHei UI", 10))
        layout.addWidget(info_label)
        
        # 密码输入框
        password_label = QLabel("新密码:")
        password_label.setFont(QFont("Microsoft YaHei UI", 12))
        layout.addWidget(password_label)
        
        self.new_password_input = QLineEdit()
        self.new_password_input.setFont(QFont("Microsoft YaHei UI", 12))
        self.new_password_input.setPlaceholderText("输入新密码")
        self.new_password_input.setEchoMode(QLineEdit.Password)
        self.new_password_input.setStyleSheet("""
            QLineEdit {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 8px;
            }
        """)
        layout.addWidget(self.new_password_input)
        
        # 确认密码输入框
        confirm_label = QLabel("确认密码:")
        confirm_label.setFont(QFont("Microsoft YaHei UI", 12))
        layout.addWidget(confirm_label)
        
        self.confirm_password_input = QLineEdit()
        self.confirm_password_input.setFont(QFont("Microsoft YaHei UI", 12))
        self.confirm_password_input.setPlaceholderText("再次输入新密码")
        self.confirm_password_input.setEchoMode(QLineEdit.Password)
        self.confirm_password_input.setStyleSheet("""
            QLineEdit {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 8px;
            }
        """)
        layout.addWidget(self.confirm_password_input)
        
        # 密码状态标签
        self.password_status = QLabel("")
        self.password_status.setFont(QFont("Microsoft YaHei UI", 10))
        layout.addWidget(self.password_status)
        
        # 保存按钮
        save_btn = QPushButton("保存密码")
        save_btn.setFixedHeight(40)
        save_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        save_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d7;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0066b4;
            }
        """)
        save_btn.clicked.connect(self.save_password)
        layout.addWidget(save_btn)
        
        # 加载当前密码状态
        self.load_password_status()
        
        return tab
    
    def create_course_edit_tab(self):
        """创建课表编辑选项卡"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)
        
        # 说明标签
        info_label = QLabel("在此编辑每周课表，每行代表一天的课程，课程之间用空格分隔")
        info_label.setFont(QFont("Microsoft YaHei UI", 10))
        layout.addWidget(info_label)
        
        # 创建表格
        self.course_table = QTableWidget()
        self.course_table.setColumnCount(1)
        self.course_table.setHorizontalHeaderLabels(["课程安排"])
        self.course_table.verticalHeader().setVisible(True)
        self.course_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
            }
            QHeaderView::section {
                background-color: #e0e0e0;
                padding: 4px;
                border: 1px solid #d0d0d0;
            }
        """)
        
        # 设置行标题（星期几）
        weekdays = ["星期一", "星期二", "星期三", "星期四", "星期五"]
        self.course_table.setRowCount(len(weekdays))
        for i, day in enumerate(weekdays):
            self.course_table.setVerticalHeaderItem(i, QTableWidgetItem(day))
        
        # 加载现有课表
        for i, day in enumerate(weekdays):
            courses = self.main_window.schedule.get(day, "")
            item = QTableWidgetItem(courses)
            self.course_table.setItem(i, 0, item)
        
        # 设置表格属性
        self.course_table.horizontalHeader().setStretchLastSection(True)
        self.course_table.setEditTriggers(QTableWidget.AllEditTriggers)
        layout.addWidget(self.course_table, 1)
        
        # 保存按钮
        save_btn = QPushButton("保存课表")
        save_btn.setFixedHeight(40)
        save_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        save_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d7;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0066b4;
            }
        """)
        save_btn.clicked.connect(self.save_course_schedule)
        layout.addWidget(save_btn)
        
        return tab
    
    def create_people_edit_tab(self):
        """创建人员编辑选项卡"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)
        
        # 说明标签
        info_label = QLabel("在此编辑随机点人的人员列表，每行一个人名")
        info_label.setFont(QFont("Microsoft YaHei UI", 10))
        layout.addWidget(info_label)
        
        # 人员列表编辑框
        self.people_edit = QTextEdit()
        self.people_edit.setFont(QFont("Microsoft YaHei UI", 12))
        self.people_edit.setStyleSheet("""
            QTextEdit {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 8px;
            }
        """)
        
        # 加载现有人员列表
        people_list = self.main_window.config.get('peopleList', "")
        self.people_edit.setPlainText(people_list.replace('、', '\n'))
        
        layout.addWidget(self.people_edit, 1)
        
        # 保存按钮
        save_btn = QPushButton("保存人员列表")
        save_btn.setFixedHeight(40)
        save_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        save_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d7;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0066b4;
            }
        """)
        save_btn.clicked.connect(self.save_people_list)
        layout.addWidget(save_btn)
        
        return tab
    
    def create_theme_setting_tab(self):
        """创建主题设置选项卡（只控制主窗口）"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)
        
        # 主题切换
        theme_label = QLabel("主题模式:")
        theme_label.setFont(QFont("Microsoft YaHei UI", 12))
        layout.addWidget(theme_label)
        
        # 主题选择按钮组
        theme_group = QButtonGroup(tab)
        
        # 浅色模式按钮
        light_theme_btn = QRadioButton("浅色模式")
        light_theme_btn.setFont(QFont("Microsoft YaHei UI", 12))
        layout.addWidget(light_theme_btn)
        theme_group.addButton(light_theme_btn)
        
        # 深色模式按钮
        dark_theme_btn = QRadioButton("深色模式")
        dark_theme_btn.setFont(QFont("Microsoft YaHei UI", 12))
        layout.addWidget(dark_theme_btn)
        theme_group.addButton(dark_theme_btn)
        
        # 设置当前选择
        if self.main_window.dark_mode:
            dark_theme_btn.setChecked(True)
        else:
            light_theme_btn.setChecked(True)
        
        # 保存按钮
        save_btn = QPushButton("应用主题")
        save_btn.setFixedHeight(40)
        save_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        save_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d7;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0066b4;
            }
        """)
        save_btn.clicked.connect(lambda: self.apply_theme(theme_group.checkedButton()))
        layout.addWidget(save_btn)
        
        # 添加伸缩空间
        layout.addStretch()
        
        return tab

    def create_auto_shutdown_tab(self):
        """创建自动关机选项卡"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)
        
        # 说明标签
        info_label = QLabel("设置晚自习后自动关机功能")
        info_label.setFont(QFont("Microsoft YaHei UI", 12))
        layout.addWidget(info_label)
        
        # 启用自动关机复选框
        self.auto_shutdown_enabled = QCheckBox("启用自动关机")
        self.auto_shutdown_enabled.setFont(QFont("Microsoft YaHei UI", 12))
        layout.addWidget(self.auto_shutdown_enabled)
        
        # 时间设置区域
        time_layout = QHBoxLayout()
        time_layout.setSpacing(10)
        
        # 小时选择
        hour_label = QLabel("时:")
        hour_label.setFont(QFont("Microsoft YaHei UI", 12))
        time_layout.addWidget(hour_label)
        
        self.hour_input = QSpinBox()
        self.hour_input.setFont(QFont("Microsoft YaHei UI", 12))
        self.hour_input.setRange(0, 23)
        self.hour_input.setValue(22)  # 默认22点
        self.hour_input.setSuffix("时")
        self.hour_input.setStyleSheet("""
            QSpinBox {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 5px;
            }
        """)
        time_layout.addWidget(self.hour_input)
        
        # 分钟选择
        minute_label = QLabel("分:")
        minute_label.setFont(QFont("Microsoft YaHei UI", 12))
        time_layout.addWidget(minute_label)
        
        self.minute_input = QSpinBox()
        self.minute_input.setFont(QFont("Microsoft YaHei UI", 12))
        self.minute_input.setRange(0, 59)
        self.minute_input.setValue(0)  # 默认0分
        self.minute_input.setSuffix("分")
        self.minute_input.setStyleSheet("""
            QSpinBox {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 5px;
            }
        """)
        time_layout.addWidget(self.minute_input)
        
        layout.addLayout(time_layout)
        
        # 状态标签
        self.shutdown_status = QLabel("")
        self.shutdown_status.setFont(QFont("Microsoft YaHei UI", 10))
        layout.addWidget(self.shutdown_status)
        
        # 保存按钮
        save_btn = QPushButton("保存设置")
        save_btn.setFixedHeight(40)
        save_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        save_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d7;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0066b4;
            }
        """)
        save_btn.clicked.connect(self.save_auto_shutdown_settings)
        layout.addWidget(save_btn)
        
        # 加载当前设置
        self.load_auto_shutdown_settings()
        
        return tab
    
    def create_group_management_tab(self):
        """创建分组管理选项卡"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)
        
        # 说明标签
        info_label = QLabel("在此管理分组信息，每组一行，格式为：组名:成员1、成员2、成员3")
        info_label.setFont(QFont("Microsoft YaHei UI", 10))
        layout.addWidget(info_label)
        
        # 分组编辑框
        self.group_edit = QTextEdit()
        self.group_edit.setFont(QFont("Microsoft YaHei UI", 12))
        self.group_edit.setStyleSheet("""
            QTextEdit {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 8px;
            }
        """)
        
        # 加载现有分组
        groups = self.main_window.config.get('groups', {})
        group_text = ""
        for group_name, members in groups.items():
            group_text += f"{group_name}:{members}\n"
        self.group_edit.setPlainText(group_text.strip())
        
        layout.addWidget(self.group_edit, 1)
        
        # 保存按钮
        save_btn = QPushButton("保存分组设置")
        save_btn.setFixedHeight(40)
        save_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        save_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d7;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0066b4;
            }
        """)
        save_btn.clicked.connect(self.save_group_settings)
        layout.addWidget(save_btn)
        
        return tab
    
    def copy_current_time(self, target):
        """获取当前时间并复制到剪贴板，同时填充到输入框"""
        try:
            # 获取当前时间
            now = datetime.datetime.now()
            
            # 格式化为两种格式
            # 1. 用于显示在输入框的格式: YYYY/MM/dd HH:mm
            input_format = now.strftime("%Y/%m/%d %H:%M")
            
            # 填充到输入框
            if target == "start":
                self.start_input.setText(input_format)
            elif target == "end":
                self.end_input.setText(input_format)
                
        except Exception as e:
            QMessageBox.critical(self, "错误", f"获取当前时间失败: {str(e)}")

    def load_exams(self):
        """从contest.json加载考试数据"""
        try:
            if os.path.exists("contest.json"):
                with open("contest.json", "r", encoding="utf-8") as f:
                    exams = json.load(f)
            else:
                exams = []
        except:
            exams = []
        
        # 设置表格行数
        self.exam_table.setRowCount(len(exams))
        
        # 填充表格
        for i, exam in enumerate(exams):
            # 科目
            subject_item = QTableWidgetItem(exam.get("subject", ""))
            self.exam_table.setItem(i, 0, subject_item)
            
            # 开始时间
            start_item = QTableWidgetItem(self.reverse_format_time(exam.get("st", "")))
            self.exam_table.setItem(i, 1, start_item)
            
            # 结束时间
            end_item = QTableWidgetItem(self.reverse_format_time(exam.get("ed", "")))
            self.exam_table.setItem(i, 2, end_item)
            
            # 关机复选框
            shutdown_item = QTableWidgetItem()
            shutdown_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            shutdown_item.setCheckState(Qt.Checked if exam.get("autoshutdown", False) else Qt.Unchecked)
            self.exam_table.setItem(i, 3, shutdown_item)
            
            # 删除按钮
            delete_btn = QPushButton("-")
            delete_btn.setFixedSize(30, 30)
            delete_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
            delete_btn.setStyleSheet("""
                QPushButton {
                    background-color: #e74c3c;
                    color: white;
                    border: none;
                    border-radius: 15px;
                }
                QPushButton:hover {
                    background-color: #c0392b;
                }
            """)
            delete_btn.clicked.connect(lambda _, row=i: self.delete_exam(row))
            self.exam_table.setCellWidget(i, 4, delete_btn)

    def reverse_format_time(self, time_str):
        """将'2025年9月13日14时5分'格式转换为'2025/09/13 14:05'"""
        try:
            # 解析新格式
            dt = datetime.strptime(time_str, "%Y年%m月%d日%H时%M分")
            
            # 格式化为原始格式
            return dt.strftime("%Y/%m/%d %H:%M")
        except:
            return time_str  # 如果解析失败，返回原始字符串

    def delete_exam(self, row):
        """删除考试"""
        self.exam_table.removeRow(row)
    
    def add_exam(self):
        """添加考试"""
        # 获取输入值
        subject = self.subject_input.text().strip()
        start_time = self.start_input.text().strip()
        end_time = self.end_input.text().strip()
        shutdown = self.shutdown_check.isChecked()
        
        # 验证输入
        if not subject:
            QMessageBox.warning(self, "错误", "请输入考试科目")
            return
            
        if not start_time or not end_time:
            QMessageBox.warning(self, "错误", "请输入开始和结束时间")
            return
            
        # 验证时间格式
        try:
            datetime.datetime.strptime(start_time, "%Y/%m/%d %H:%M")
            datetime.datetime.strptime(end_time, "%Y/%m/%d %H:%M")
        except ValueError:
            QMessageBox.warning(self, "错误", "时间格式不正确，应为 YYYY/MM/dd HH:mm")
            return
        
        # 添加新行
        row = self.exam_table.rowCount()
        self.exam_table.insertRow(row)
        
        # 科目
        subject_item = QTableWidgetItem(subject)
        self.exam_table.setItem(row, 0, subject_item)
        
        # 开始时间
        start_item = QTableWidgetItem(start_time)
        self.exam_table.setItem(row, 1, start_item)
        
        # 结束时间
        end_item = QTableWidgetItem(end_time)
        self.exam_table.setItem(row, 2, end_item)
        
        # 关机复选框
        shutdown_item = QTableWidgetItem()
        shutdown_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
        shutdown_item.setCheckState(Qt.Checked if shutdown else Qt.Unchecked)
        self.exam_table.setItem(row, 3, shutdown_item)
        
        # 删除按钮
        delete_btn = QPushButton("-")
        delete_btn.setFixedSize(30, 30)
        delete_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        delete_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                border: none;
                border-radius: 15px;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        delete_btn.clicked.connect(lambda _, row=row: self.delete_exam(row))
        self.exam_table.setCellWidget(row, 4, delete_btn)
        
        # 清空输入框
        self.subject_input.clear()
        self.start_input.clear()
        self.end_input.clear()
        self.shutdown_check.setChecked(False)
    
    def save_exams(self):
        """保存考试设置为JSON格式"""
        exams = []
        
        # 从表格中提取考试数据
        for row in range(self.exam_table.rowCount()):
            subject = self.exam_table.item(row, 0).text()
            start_time = self.exam_table.item(row, 1).text()
            end_time = self.exam_table.item(row, 2).text()
            shutdown = self.exam_table.item(row, 3).checkState() == Qt.Checked
            
            # 转换为新格式
            exams.append({
                "subject": subject,
                "st": self.format_time(start_time),
                "ed": self.format_time(end_time),
                "autoshutdown": shutdown
            })
        
        # 保存到JSON文件
        with open("contest.json", "w", encoding="utf-8") as f:
            json.dump(exams, f, ensure_ascii=False, indent=4)
        
        QMessageBox.information(self, "保存成功", "考试设置已保存")

    def format_time(self, time_str):
        """将时间格式化为'2025年9月13日14时5分'格式"""
        # 解析原始时间
        dt = datetime.datetime.strptime(time_str, "%Y/%m/%d %H:%M")
        
        # 格式化新时间
        # 使用格式化字符串确保没有前导零
        year = dt.strftime("%Y")
        month = dt.strftime("%m")  # 没有前导零的月份
        day = dt.strftime("%d")    # 没有前导零的日期
        hour = dt.strftime("%H")   # 没有前导零的小时
        minute = dt.strftime("%M") # 没有前导零的分钟
        
        return f"{year}年{month}月{day}日{hour}时{minute}分"

    def save_class_reminders(self):
        """保存下课提醒设置"""
        reminder_times = self.reminder_input.text().strip()
        
        # 验证格式
        if reminder_times:
            times = reminder_times.split('、')
            for time_str in times:
                try:
                    # 验证时间格式
                    datetime.datetime.strptime(time_str, '%H:%M')
                except ValueError:
                    self.reminder_status.setText(f"时间格式错误: {time_str} (应为HH:MM格式)")
                    self.reminder_status.setStyleSheet("color: red;")
                    return
        
        # 保存到配置
        self.main_window.config['class_reminders'] = reminder_times
        self.main_window.save_config()
        
        # 更新主窗口的提醒时间
        self.main_window.class_reminders = times if reminder_times else []
        
        # 更新状态
        self.reminder_status.setText("提醒设置已保存")
        self.reminder_status.setStyleSheet("color: green;")
    
    def test_class_reminder(self):
        """测试提醒效果"""
        self.main_window.show_class_reminder()
    
    def save_group_settings(self):
        """保存分组设置"""
        # 获取分组文本
        group_text = self.group_edit.toPlainText().strip()
        if not group_text:
            QMessageBox.warning(self, "错误", "分组信息不能为空")
            return
        
        # 解析分组
        groups = {}
        for line in group_text.split('\n'):
            if ':' in line:
                group_name, members = line.split(':', 1)
                group_name = group_name.strip()
                members = members.strip()
                if group_name and members:
                    groups[group_name] = members
        
        if not groups:
            QMessageBox.warning(self, "错误", "未找到有效的分组信息")
            return
        
        # 保存到配置
        self.main_window.config['groups'] = groups
        self.main_window.save_config()
        
        QMessageBox.information(self, "保存成功", "分组设置已保存")

    
    def load_auto_shutdown_settings(self):
        """加载自动关机设置"""
        # 获取配置
        auto_shutdown_enabled = self.main_window.config.get('auto_shutdown_enabled', False)
        auto_shutdown_time = self.main_window.config.get('auto_shutdown_time', "22:00")
        
        # 设置界面控件
        self.auto_shutdown_enabled.setChecked(auto_shutdown_enabled)
        
        # 解析时间
        if auto_shutdown_time:
            try:
                hour, minute = map(int, auto_shutdown_time.split(':'))
                self.hour_input.setValue(hour)
                self.minute_input.setValue(minute)
            except:
                pass
        
        # 更新状态标签
        self.update_shutdown_status()
    
    def update_shutdown_status(self):
        """更新自动关机状态标签"""
        enabled = self.auto_shutdown_enabled.isChecked()
        hour = self.hour_input.value()
        minute = self.minute_input.value()
        
        if enabled:
            self.shutdown_status.setText(f"已启用自动关机，将在每天 {hour:02d}:{minute:02d} 关机")
            self.shutdown_status.setStyleSheet("color: green;")
        else:
            self.shutdown_status.setText("自动关机已禁用")
            self.shutdown_status.setStyleSheet("color: red;")
    
    def save_auto_shutdown_settings(self):
        """保存自动关机设置"""
        # 获取设置值
        enabled = self.auto_shutdown_enabled.isChecked()
        hour = self.hour_input.value()
        minute = self.minute_input.value()
        shutdown_time = f"{hour:02d}:{minute:02d}"
        
        # 保存到配置
        self.main_window.config['auto_shutdown_enabled'] = enabled
        self.main_window.config['auto_shutdown_time'] = shutdown_time
        self.main_window.save_config()
        
        # 更新主窗口的自动关机设置
        self.main_window.auto_shutdown_enabled = enabled
        self.main_window.auto_shutdown_time = shutdown_time
        
        # 更新状态标签
        self.update_shutdown_status()
        
        QMessageBox.information(self, "保存成功", "自动关机设置已保存")
    
    def apply_theme(self, selected_button):
        """应用主题（只改变主窗口）"""
        if selected_button.text() == "深色模式":
            self.main_window.dark_mode = True
        else:
            self.main_window.dark_mode = False
        
        # 更新配置
        self.main_window.config['darkmode'] = self.main_window.dark_mode
        self.main_window.save_config()
        
        # 应用新主题到主窗口
        self.main_window.apply_theme()
        self.main_window.update()
        
        QMessageBox.information(self, "主题已应用", "主题设置已保存并应用")
    
    def save_course_schedule(self):
        """保存课表"""
        # 获取课表数据
        weekdays = ["星期一", "星期二", "星期三", "星期四", "星期五"]
        schedule = {}
        
        for i in range(self.course_table.rowCount()):
            day = weekdays[i]
            item = self.course_table.item(i, 0)
            if item and item.text().strip():
                schedule[day] = item.text().strip()
        
        # 更新配置
        self.main_window.schedule = schedule
        self.main_window.config['schedule'] = schedule
        self.main_window.save_config()
        
        # 更新主窗口显示
        self.main_window.update_all_content()
        self.main_window.adjust_window_size()
        
        QMessageBox.information(self, "保存成功", "课表已保存")
    
    def save_people_list(self):
        """保存人员列表"""
        # 获取人员列表
        people_text = self.people_edit.toPlainText().strip()
        people_list = people_text.replace('\n', '、')
        
        # 更新配置
        self.main_window.config['peopleList'] = people_list
        self.main_window.save_config()
        
        QMessageBox.information(self, "保存成功", "人员列表已保存")
    
    def create_app_management_page(self):
        """创建程序管理页面"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # 标题
        title = QLabel("程序管理")
        title_font = QFont("Microsoft YaHei UI", 18, QFont.Bold)
        title.setFont(title_font)
        layout.addWidget(title)
        
        # 分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("background-color: #e0e0e0;")
        layout.addWidget(separator)
        
        # 按钮布局
        button_layout = QVBoxLayout()
        button_layout.setSpacing(15)
        
        # 重启程序按钮
        restart_app_btn = QPushButton("重启程序")
        restart_app_btn.setFixedHeight(50)
        restart_app_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        restart_app_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        restart_app_btn.clicked.connect(self.restart_application)
        button_layout.addWidget(restart_app_btn)
        
        # 退出程序按钮
        exit_app_btn = QPushButton("退出程序")
        exit_app_btn.setFixedHeight(50)
        exit_app_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        exit_app_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        exit_app_btn.clicked.connect(self.exit_application)
        button_layout.addWidget(exit_app_btn)
        
        auto_start_btn = QPushButton("添加开机自启")
        auto_start_btn.setFixedHeight(50)
        auto_start_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        auto_start_btn.setStyleSheet("""
            QPushButton {
                background-color: #2ecc71;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
        """)
        auto_start_btn.clicked.connect(self.add_auto_start)
        button_layout.addWidget(auto_start_btn)

        # 添加更新日志按钮
        change_log_btn = QPushButton("更新日志")
        change_log_btn.setFixedHeight(50)
        change_log_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        change_log_btn.setStyleSheet("""
            QPushButton {
                background-color: #ffd45d;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #e6b400;
            }
        """)
        change_log_btn.clicked.connect(self.show_changeLog)
        button_layout.addWidget(change_log_btn)

        # 添加关于按钮
        about_btn = QPushButton("关于")
        about_btn.setFixedHeight(50)
        about_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        about_btn.setStyleSheet("""
            QPushButton {
                background-color: #9b59b6;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #8e44ad;
            }
        """)
        about_btn.clicked.connect(self.show_about_info)
        button_layout.addWidget(about_btn)
        
        layout.addLayout(button_layout)
        
        # 添加伸缩空间
        layout.addStretch()
        
        return page
    
    def add_auto_start(self):
        """添加开机自启功能"""
        try:
            # 检查是否已经添加过
            if self.is_auto_start_enabled():
                QMessageBox.information(self, "已添加", "开机自启已经设置过，无需重复添加")
                return
            
            # 二次确认
            reply = QMessageBox.question(
                self, 
                "确认添加开机自启", 
                "确定要添加开机自启功能吗？\n\n程序将在每次系统启动时自动运行",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            
            if reply != QMessageBox.Yes:
                return
            
            # 获取当前脚本路径
            script_path = os.path.abspath(sys.argv[0])
            
            # 如果是.py文件，使用pythonw运行
            if script_path.endswith(".py") or script_path.endswith(".pyw"):
                # 获取Python解释器路径
                python_exe = sys.executable
                if not python_exe:
                    QMessageBox.critical(self, "错误", "无法获取Python解释器路径")
                    return
                    
                # 使用pythonw运行，避免显示控制台窗口
                if python_exe.endswith("python.exe"):
                    pythonw_exe = python_exe.replace("python.exe", "pythonw.exe")
                else:
                    pythonw_exe = python_exe
                    
                # 检查pythonw.exe是否存在
                if not os.path.exists(pythonw_exe):
                    QMessageBox.critical(self, "错误", f"找不到pythonw.exe: {pythonw_exe}")
                    return
                    
                command = f'"{pythonw_exe}" "{script_path}"'
            else:
                # 如果是可执行文件，直接运行
                command = f'"{script_path}"'
            
            key = winreg.HKEY_CURRENT_USER
            subkey = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
            
            try:
                # 打开注册表项
                with winreg.OpenKey(key, subkey, 0, winreg.KEY_WRITE) as reg_key:
                    # 设置注册表值
                    winreg.SetValueEx(reg_key, "SPJClassTool", 0, winreg.REG_SZ, command)
                
                QMessageBox.information(self, "成功", "开机自启已添加成功")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"添加开机自启失败: {str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"添加开机自启时出错: {str(e)}")

    def is_auto_start_enabled(self):
        """检查是否已经添加了开机自启"""
        try:
            import winreg
            key = winreg.HKEY_CURRENT_USER
            subkey = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
            
            # 打开注册表项
            with winreg.OpenKey(key, subkey, 0, winreg.KEY_READ) as reg_key:
                try:
                    # 尝试获取值
                    value, _ = winreg.QueryValueEx(reg_key, "SPJClassTool")
                    return True
                except FileNotFoundError:
                    # 如果找不到值，表示未添加
                    return False
        except:
            return False

    def show_about_info(self):
        """显示关于信息"""
        about_text = """
            <h2>刷屏君 课堂工具 v1.0.1</h2>
            <p>这是一个专为课堂设计的实用工具，帮助教师管理课程、学生和时间。</p>
            <p><b>主要功能：</b></p>
            <ul>
                <li>临时改课 - 快速调整和显示课程安排</li>
                <li>考试功能 - 配置考试</li>
                <li>检查更新 - 自动检查并安装新版本</li>
                <li>电源管理 - 控制计算机电源状态</li>
                <li>程序管理 - 重启或退出程序</li>
            </ul>
            <p><b>开发者：</b> 刷屏君个人</p>
            <p><b>联系方式：</b> BXJZDXMM@vip.qq.com</p>
            <p><b>官方网站：</b> <a href="http://spj2025.top">http://spj2025.top</a></p>
            <p><b>版权信息：</b> © 2025 刷屏君团队 保留所有权利</p>
        """
        
        QMessageBox.information(
            self,
            "关于",
            about_text,
            QMessageBox.Ok
        )
    
    def show_changeLog(self):
        """显示关于信息"""
        about_text = """
            <h2>刷屏君 课堂工具 v1.0.1 更新日志</h2>
            <p><b>v1.0.1</b></p>
            <ul>
                <li>修复未配置课程表时程序异常崩溃的问题</li>
                <li>修复保存密码时程序异常崩溃的问题</li>
            </ul>
            <p><b>v1.0.0</b></p>
            <ul>
                <li>初代 pyclasstool 问世</li>
            </ul>
        """
        
        QMessageBox.information(
            self,
            "更新日志",
            about_text,
            QMessageBox.Ok
        )
    
    def apply_theme_to_settings(self):
        """应用主题到设置窗口"""
        if self.main_window.dark_mode:
            # 深色模式
            self.setStyleSheet("""
                QWidget {
                    background-color: #2d2d2d;
                    color: #f0f0f0;
                }
                QTabWidget::pane {
                    border: 1px solid #505050;
                    background-color: #3c3c3c;
                }
                QTabBar::tab {
                    background-color: #3c3c3c;
                    color: #f0f0f0;
                    padding: 8px;
                    border: 1px solid #505050;
                    border-bottom: none;
                    border-top-left-radius: 4px;
                    border-top-right-radius: 4px;
                }
                QTabBar::tab:selected {
                    background-color: #505050;
                }
                QTableWidget {
                    background-color: #3c3c3c;
                    color: #f0f0f0;
                    gridline-color: #505050;
                }
                QHeaderView::section {
                    background-color: #505050;
                    color: #f0f0f0;
                    border: 1px solid #606060;
                }
                QTextEdit {
                    background-color: #3c3c3c;
                    color: #f0f0f0;
                }
                QRadioButton {
                    color: #f0f0f0;
                }
            """)
        else:
            # 浅色模式
            self.setStyleSheet("""
                QWidget {
                    background-color: #ffffff;
                    color: #000000;
                }
                QTabWidget::pane {
                    border: 1px solid #d0d0d0;
                    background-color: #f0f0f0;
                }
                QTabBar::tab {
                    background-color: #f0f0f0;
                    color: #000000;
                    padding: 8px;
                    border: 1px solid #d0d0d0;
                    border-bottom: none;
                    border-top-left-radius: 4px;
                    border-top-right-radius: 4px;
                }
                QTabBar::tab:selected {
                    background-color: #ffffff;
                }
                QTableWidget {
                    background-color: #ffffff;
                    color: #000000;
                    gridline-color: #d0d0d0;
                }
                QHeaderView::section {
                    background-color: #e0e0e0;
                    color: #000000;
                    border: 1px solid #d0d0d0;
                }
                QTextEdit {
                    background-color: #ffffff;
                    color: #000000;
                }
                QRadioButton {
                    color: #000000;
                }
            """)

    def restart_application(self):
        """重启应用程序"""
        # 关闭设置窗口
        self.close()
        
        # 重启程序
        current_script = os.path.join(os.path.dirname(__file__), "restart.pyw")
        subprocess.Popen(["pythonw", current_script])
        
        # 退出当前程序
        QApplication.quit()
    
    def exit_application(self):
        """退出应用程序"""
        # 检查是否设置了密码
        if 'password' in self.main_window.config:
            # 弹出密码输入对话框
            from PyQt5.QtWidgets import QInputDialog
            password, ok = QInputDialog.getText(
                self, 
                "密码验证", 
                "请输入管理密码:", 
                QLineEdit.Password
            )
            
            if not ok:
                return  # 用户取消
                
            # 验证密码
            if not self.main_window.verify_password(password):
                QMessageBox.critical(self, "密码错误", "输入的管理密码不正确")
                return
        
        # 退出程序
        QApplication.quit()
    
    def load_password_status(self):
        """加载密码状态"""
        if 'password' in self.main_window.config:
            self.password_status.setText("已设置管理密码")
            self.password_status.setStyleSheet("color: green;")
        else:
            self.password_status.setText("未设置管理密码")
            self.password_status.setStyleSheet("color: red;")
    
    def save_password(self):
        """保存密码"""
        password = self.new_password_input.text().strip()
        confirm_password = self.confirm_password_input.text().strip()
        
        if not password:
            self.password_status.setText("密码不能为空")
            self.password_status.setStyleSheet("color: red;")
            return
        
        if password != confirm_password:
            self.password_status.setText("两次输入的密码不一致")
            self.password_status.setStyleSheet("color: red;")
            return
        
        # 加密密码
        encrypted_password = self.encrypt_password(password)
        
        # 保存到配置
        self.main_window.config['password'] = encrypted_password
        self.main_window.save_config()
        
        # 更新状态
        self.password_status.setText("密码保存成功")
        self.password_status.setStyleSheet("color: green;")
        
        # 清空输入框
        self.new_password_input.clear()
        self.confirm_password_input.clear()
    
    def encrypt_password(self, password):
        """加密密码"""
        # 生成盐值
        salt = os.urandom(16)
        
        # 使用PBKDF2HMAC算法生成密钥
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=salt,
            iterations=100000,
            backend=default_backend()
        )
        key = base64.urlsafe_b64encode(kdf.derive(password.encode()))
        
        # 使用Fernet加密密码
        f = Fernet(key)
        encrypted_password = f.encrypt(password.encode())
        
        # 返回加密后的密码和盐值（以冒号分隔）
        return f"{base64.b64encode(salt).decode()}:{encrypted_password.decode()}"
    
    def shutdown_computer(self):
        """关机计算机"""
        try:
            # 调用系统关机程序
            subprocess.Popen(["C:\\Windows\\System32\\slidetoshutdown.exe"])
        except Exception as e:
            print(f"关机失败: {e}")

    def confirm_restart(self):
        """确认重启操作"""
        reply = QMessageBox.question(
            self, 
            "确认重启", 
            "确实要重启计算机吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.restart_computer()

    def restart_computer(self):
        """重启计算机"""
        try:
            # 调用系统重启命令
            subprocess.Popen(["shutdown", "/r", "/t", "0"])
        except Exception as e:
            print(f"重启失败: {e}")

    def confirm_logout(self):
        """确认注销操作"""
        reply = QMessageBox.question(
            self, 
            "确认注销", 
            "确实要注销当前用户吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.logout_user()

    def logout_user(self):
        """注销当前用户"""
        try:
            # 调用系统注销命令
            subprocess.Popen(["logoff"])
        except Exception as e:
            print(f"注销失败: {e}")

    def confirm_force_shutdown(self):
        """确认强制关机操作"""
        reply = QMessageBox.question(
            self, 
            "确认强制关机", 
            "确实要强制关机吗？此操作可能导致数据丢失！",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.force_shutdown()

    def force_shutdown(self):
        """强制关机"""
        try:
            # 获取当前目录
            current_dir = os.path.dirname(os.path.abspath(__file__))
            # 调用强制关机程序
            subprocess.Popen([os.path.join(current_dir, "ForceShutdown.exe")])
        except Exception as e:
            print(f"强制关机失败: {e}")

    def select_day_courses(self, day):
        """选择指定星期的课程"""
        # 获取主窗口的课表配置
        courses = self.main_window.schedule.get(day, [])
        
        # 更新主窗口的显示
        self.main_window.schedule_label.setText(courses)
        
        # 更新自定义课表输入框
        self.course_input.setPlainText(courses)
        
        # 调整主窗口大小
        self.main_window.adjust_window_size()
    
    def apply_custom_courses(self):
        """应用自定义课表"""
        # 获取输入的课表
        custom_courses = self.course_input.toPlainText().strip()
        
        if custom_courses:
            # 更新主窗口的课表显示
            self.main_window.schedule_label.setText(custom_courses)
            
            # 更新配置（保存为字符串）
            day_cn = self.main_window.weekday_label.text()
            self.main_window.schedule[day_cn] = custom_courses
            self.main_window.config['schedule'] = self.main_window.schedule
            self.main_window.save_config()
            
            # 调整主窗口大小
            self.main_window.adjust_window_size()
    
    def start_update_check(self):
        """开始检查更新"""
        # 显示加载状态
        self.update_status.setText("正在检查更新.")
        self.progress_bar.setVisible(True)
        self.update_result_container.setVisible(False)
        
        # 启动加载动画
        self.dot_count = 1
        self.update_timer = QTimer(self)
        self.update_timer.timeout.connect(self.update_dots)
        self.update_timer.start(500)  # 每500毫秒更新一次
        
        # 启动更新检查线程
        self.update_checker = UpdateChecker(self.main_window.config.get('version', '1.0'))
        self.update_checker.update_found.connect(self.show_update)
        self.update_checker.no_update.connect(self.show_no_update)
        self.update_checker.error_occurred.connect(self.show_update_error)
        self.update_checker.start()
    
    def update_dots(self):
        """更新加载动画的点"""
        self.dot_count = (self.dot_count % 3) + 1
        dots = "." * self.dot_count
        self.update_status.setText(f"正在检查更新{dots}")
    
    def show_update(self, version, changelog):
        """显示更新信息"""
        self.update_timer.stop()
        self.progress_bar.setVisible(False)
        self.update_result_container.setVisible(True)
        
        self.update_title.setText(f"发现新版本 {version}")
        self.changelog_text.setPlainText(changelog)
    
    def show_no_update(self):
        """显示无更新信息"""
        self.update_timer.stop()
        self.progress_bar.setVisible(False)
        self.update_status.setText("已是最新版本")
    
    def show_update_error(self, error):
        """显示更新错误信息"""
        self.update_timer.stop()
        self.progress_bar.setVisible(False)
        self.update_status.setText(f"更新检查失败: {error}")
    
    def start_update(self):
        """启动更新程序"""
        # 关闭设置窗口
        self.close()
        
        # 启动更新程序
        update_script = os.path.join(os.path.dirname(__file__), "update.pyw")
        subprocess.Popen(["python", update_script])
        exit()

class RandomPickerWindow(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.setWindowTitle("随机点人")
        self.setGeometry(100, 100, 800, 600)
        self.setMinimumSize(800, 600)
        
        # 创建主布局
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # 创建左侧导航栏
        self.nav_frame = QFrame()
        self.nav_frame.setFixedWidth(200)
        self.nav_frame.setStyleSheet("background-color: #f0f0f0;")
        nav_layout = QVBoxLayout(self.nav_frame)
        nav_layout.setContentsMargins(0, 10, 0, 10)
        nav_layout.setSpacing(0)
        
        # 导航列表
        self.nav_list = QListWidget()
        self.nav_list.setStyleSheet("""
            QListWidget {
                border: none;
                background-color: transparent;
            }
            QListWidget::item {
                padding: 10px 15px;
                border-bottom: 1px solid #e0e0e0;
            }
            QListWidget::item:selected {
                background-color: #e6e6e6;
                color: #000;
            }
        """)
        self.nav_list.setFixedWidth(200)
        
        # 添加导航项
        nav_items = ["随机点人", "分组点人"]
        for item in nav_items:
            list_item = QListWidgetItem(item)
            list_item.setSizeHint(QSize(200, 40))
            self.nav_list.addItem(list_item)
        
        nav_layout.addWidget(self.nav_list)
        nav_layout.addStretch()
        
        # 创建右侧内容区域
        self.content_area = QStackedWidget()
        
        # 添加随机点人页面
        self.random_pick_page = self.create_random_pick_page()
        self.content_area.addWidget(self.random_pick_page)
        
        # 添加分组点人页面
        self.group_pick_page = self.create_group_pick_page()
        self.content_area.addWidget(self.group_pick_page)
        
        # 连接导航选择事件
        self.nav_list.currentRowChanged.connect(self.content_area.setCurrentIndex)
        
        # 添加到主布局
        main_layout.addWidget(self.nav_frame)
        main_layout.addWidget(self.content_area)
        
        # 设置初始选择 - 默认显示随机点人页面
        self.nav_list.setCurrentRow(0)
    
    def create_random_pick_page(self):
        """创建随机点人页面"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # 标题
        title = QLabel("随机点人")
        title_font = QFont("Microsoft YaHei UI", 18, QFont.Bold)
        title.setFont(title_font)
        layout.addWidget(title)
        
        # 分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("background-color: #e0e0e0;")
        layout.addWidget(separator)
        
        # 人数选择区域
        count_layout = QHBoxLayout()
        count_layout.setSpacing(10)
        
        count_label = QLabel("抽取人数:")
        count_label.setFont(QFont("Microsoft YaHei UI", 12))
        count_layout.addWidget(count_label)
        
        # 人数输入框
        self.count_input = QSpinBox()
        self.count_input.setFont(QFont("Microsoft YaHei UI", 12))
        self.count_input.setMinimum(1)
        self.count_input.setMaximum(20)  # 最大抽取20人
        self.count_input.setValue(1)
        self.count_input.setStyleSheet("""
            QSpinBox {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 5px;
            }
        """)
        count_layout.addWidget(self.count_input)
        
        # 添加伸缩空间
        count_layout.addStretch()
        
        layout.addLayout(count_layout)
        
        # 抽取按钮
        pick_btn = QPushButton("抽取")
        pick_btn.setFixedHeight(40)
        pick_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        pick_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d7;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0066b4;
            }
        """)
        pick_btn.clicked.connect(self.pick_random_people)
        layout.addWidget(pick_btn)
        
        # 结果标签
        result_label = QLabel("抽取结果:")
        result_label.setFont(QFont("Microsoft YaHei UI", 12))
        layout.addWidget(result_label)
        
        # 结果文本框
        self.random_result_text = QTextEdit()
        self.random_result_text.setReadOnly(True)
        self.random_result_text.setFont(QFont("Microsoft YaHei UI", 12))
        self.random_result_text.setStyleSheet("""
            QTextEdit {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 8px;
                background-color: #f8f8f8;
            }
        """)
        self.random_result_text.setFixedHeight(100)
        layout.addWidget(self.random_result_text)
        
        # 放大结果按钮
        enlarge_btn = QPushButton("放大结果")
        enlarge_btn.setFixedHeight(40)
        enlarge_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        enlarge_btn.setStyleSheet("""
            QPushButton {
                background-color: #107c10;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0e6e0e;
            }
        """)
        enlarge_btn.clicked.connect(self.enlarge_random_result)
        layout.addWidget(enlarge_btn)
        
        # 添加伸缩空间
        layout.addStretch()
        
        return page
    
    def pick_random_people(self):
        """随机抽取人员（抽取后移除，抽完后重置）"""
        # 获取人员列表
        people_list = self.main_window.config.get('peopleList', "")
        if not people_list:
            QMessageBox.warning(self, "错误", "人员列表为空，请在设置中添加人员")
            return
        
        # 分割人员列表
        people = people_list.split('、')
        if not people:
            QMessageBox.warning(self, "错误", "人员列表格式不正确")
            return
        
        # 获取抽取人数
        count = self.count_input.value()
        if count > len(people):
            QMessageBox.warning(self, "错误", f"抽取人数不能超过总人数 {len(people)}")
            return
        
        # 从主窗口获取当前可用人员列表
        if not hasattr(self.main_window, 'available_people'):
            # 第一次抽取，初始化可用人员列表
            self.main_window.available_people = people.copy()
        
        # 如果可用人员不足，重置为原始人员列表
        if len(self.main_window.available_people) < count:
            self.main_window.available_people = people.copy()
        
        # 随机抽取
        selected = []
        for _ in range(count):
            if not self.main_window.available_people:
                # 如果抽完了，重置列表
                self.main_window.available_people = people.copy()
            
            # 随机选择一个人
            person = random.choice(self.main_window.available_people)
            selected.append(person)
            
            # 从可用人员中移除
            self.main_window.available_people.remove(person)
        
        # 显示结果
        self.random_result_text.setPlainText("、".join(selected))
    
    def enlarge_random_result(self):
        """放大显示随机点人结果"""
        result = self.random_result_text.toPlainText().strip()
        if not result:
            QMessageBox.warning(self, "错误", "没有抽取结果")
            return
        
        # 创建放大窗口
        self.enlarge_window = QWidget()
        self.enlarge_window.setWindowFlags(
            Qt.FramelessWindowHint | 
            Qt.WindowStaysOnTopHint
        )
        self.enlarge_window.setAttribute(Qt.WA_TranslucentBackground)
        self.enlarge_window.setFixedSize(1920, 300)
        
        # 设置布局
        layout = QVBoxLayout(self.enlarge_window)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # 结果标签
        result_label = QLabel(result)
        result_label.setAlignment(Qt.AlignCenter)
        result_label.setFont(QFont("Microsoft YaHei UI", 34, QFont.Bold))
        result_label.setStyleSheet("color: #000000; background-color: rgba(255, 255, 255, 200);")
        result_label.setWordWrap(True)  # 启用自动换行
        layout.addWidget(result_label, 1)
        
        # 返回按钮
        return_btn = QPushButton("点击此处返回")
        return_btn.setFixedHeight(60)
        return_btn.setFont(QFont("Microsoft YaHei UI", 24, QFont.Bold))
        return_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d7;
                color: white;
                border: none;
                border-radius: 0px;
            }
            QPushButton:hover {
                background-color: #0066b4;
            }
        """)
        return_btn.clicked.connect(self.enlarge_window.close)
        layout.addWidget(return_btn)
        
        # 居中显示窗口
        screen_geometry = QApplication.desktop().screenGeometry()
        x = (screen_geometry.width() - 1920) // 2
        y = (screen_geometry.height() - 300) // 2
        self.enlarge_window.move(x, y)
        
        # 显示窗口
        self.enlarge_window.show()
        
        # 设置30秒后自动关闭
        QTimer.singleShot(30000, self.enlarge_window.close)
    
    def create_group_pick_page(self):
        """创建分组点人页面"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # 标题
        title = QLabel("分组点人")
        title_font = QFont("Microsoft YaHei UI", 18, QFont.Bold)
        title.setFont(title_font)
        layout.addWidget(title)
        
        # 分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("background-color: #e0e0e0;")
        layout.addWidget(separator)
        
        # 组选择区域
        group_layout = QHBoxLayout()
        group_layout.setSpacing(10)
        
        group_label = QLabel("选择组:")
        group_label.setFont(QFont("Microsoft YaHei UI", 12))
        group_layout.addWidget(group_label)
        
        # 组选择下拉框
        self.group_combo = QComboBox()
        self.group_combo.setFont(QFont("Microsoft YaHei UI", 12))
        self.group_combo.setStyleSheet("""
            QComboBox {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 5px;
            }
        """)
        self.group_combo.setMinimumWidth(150)
        group_layout.addWidget(self.group_combo, 1)
        
        # 添加伸缩空间
        group_layout.addStretch()
        
        layout.addLayout(group_layout)
        
        # 组成员标签 - 添加自动换行和高度设置
        self.group_members_label = QLabel("该组成员: ")
        self.group_members_label.setFont(QFont("Microsoft YaHei UI", 12))
        self.group_members_label.setWordWrap(True)  # 启用自动换行
        self.group_members_label.setFixedHeight(60)  # 设置高度为40像素（两行）
        self.group_members_label.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        layout.addWidget(self.group_members_label)
        
        # 抽取按钮
        pick_btn = QPushButton("抽取成员")
        pick_btn.setFixedHeight(40)
        pick_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        pick_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d7;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0066b4;
            }
        """)
        pick_btn.clicked.connect(self.pick_group_member)
        layout.addWidget(pick_btn)
        
        # 结果标签
        result_label = QLabel("抽取结果:")
        result_label.setFont(QFont("Microsoft YaHei UI", 12))
        layout.addWidget(result_label)
        
        # 结果文本框
        self.group_result_text = QTextEdit()
        self.group_result_text.setReadOnly(True)
        self.group_result_text.setFont(QFont("Microsoft YaHei UI", 12))
        self.group_result_text.setStyleSheet("""
            QTextEdit {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 8px;
                background-color: #f8f8f8;
            }
        """)
        self.group_result_text.setFixedHeight(60)
        layout.addWidget(self.group_result_text)
        
        # 放大结果按钮
        enlarge_btn = QPushButton("放大结果")
        enlarge_btn.setFixedHeight(40)
        enlarge_btn.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        enlarge_btn.setStyleSheet("""
            QPushButton {
                background-color: #107c10;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0e6e0e;
            }
        """)
        enlarge_btn.clicked.connect(self.enlarge_group_result)
        layout.addWidget(enlarge_btn)
        
        # 添加伸缩空间
        layout.addStretch()
        
        # 加载分组数据
        self.load_groups()
        
        # 连接组选择变化事件
        self.group_combo.currentIndexChanged.connect(self.update_group_members)
        
        return page
    
    def load_groups(self):
        """加载分组数据"""
        # 获取分组配置
        groups = self.main_window.config.get('groups', {})
        
        # 清空下拉框
        self.group_combo.clear()
        
        # 添加分组
        for group_name in groups.keys():
            self.group_combo.addItem(group_name)
        
        # 如果有分组，更新成员显示
        if groups:
            self.update_group_members(0)
    
    def update_group_members(self, index):
        """更新组成员显示"""
        if index < 0:
            return
            
        # 获取分组配置
        groups = self.main_window.config.get('groups', {})
        
        # 获取当前选择的组名
        group_name = self.group_combo.currentText()
        
        # 获取组成员
        members = groups.get(group_name, "")
        
        # 更新显示
        self.group_members_label.setText(f"该组成员: {members}")
    
    def pick_group_member(self):
        """从组中随机抽取一个成员"""
        # 获取分组配置
        groups = self.main_window.config.get('groups', {})
        
        # 获取当前选择的组名
        group_name = self.group_combo.currentText()
        if not group_name:
            QMessageBox.warning(self, "错误", "请先选择一个组")
            return
        
        # 获取组成员
        members_str = groups.get(group_name, "")
        if not members_str:
            QMessageBox.warning(self, "错误", "该组没有成员")
            return
        
        # 分割成员
        members = members_str.split('、')
        if not members:
            QMessageBox.warning(self, "错误", "该组没有成员")
            return
        
        # 初始化可用成员列表
        if not hasattr(self, 'available_members'):
            self.available_members = {}
        
        # 初始化当前组的可用成员
        if group_name not in self.available_members:
            self.available_members[group_name] = members.copy()
        
        # 如果当前组没有可用成员，重置列表
        if not self.available_members[group_name]:
            self.available_members[group_name] = members.copy()
        
        # 随机抽取一个成员
        person = random.choice(self.available_members[group_name])
        
        # 从可用成员中移除
        self.available_members[group_name].remove(person)
        
        # 显示结果
        self.group_result_text.setPlainText(person)
    
    def enlarge_group_result(self):
        """放大显示分组点人结果"""
        result = self.group_result_text.toPlainText().strip()
        if not result:
            QMessageBox.warning(self, "错误", "没有抽取结果")
            return
        
        # 创建放大窗口
        self.enlarge_window = QWidget()
        self.enlarge_window.setWindowFlags(
            Qt.FramelessWindowHint | 
            Qt.WindowStaysOnTopHint
        )
        self.enlarge_window.setAttribute(Qt.WA_TranslucentBackground)
        self.enlarge_window.setFixedSize(1920, 300)
        
        # 设置布局
        layout = QVBoxLayout(self.enlarge_window)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # 结果标签
        result_label = QLabel(result)
        result_label.setAlignment(Qt.AlignCenter)
        result_label.setFont(QFont("Microsoft YaHei UI", 34, QFont.Bold))
        result_label.setStyleSheet("color: #000000; background-color: rgba(255, 255, 255, 200);")
        result_label.setWordWrap(True)  # 启用自动换行
        layout.addWidget(result_label, 1)
        
        # 返回按钮
        return_btn = QPushButton("点击此处返回")
        return_btn.setFixedHeight(60)
        return_btn.setFont(QFont("Microsoft YaHei UI", 24, QFont.Bold))
        return_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d7;
                color: white;
                border: none;
                border-radius: 0px;
            }
            QPushButton:hover {
                background-color: #0066b4;
            }
        """)
        return_btn.clicked.connect(self.enlarge_window.close)
        layout.addWidget(return_btn)
        
        # 居中显示窗口
        screen_geometry = QApplication.desktop().screenGeometry()
        x = (screen_geometry.width() - 1920) // 2
        y = (screen_geometry.height() - 300) // 2
        self.enlarge_window.move(x, y)
        
        # 显示窗口
        self.enlarge_window.show()
        
        # 设置30秒后自动关闭
        QTimer.singleShot(30000, self.enlarge_window.close)

class ScheduleWindow(QWidget):
    def __init__(self):
        super().__init__()
        
        # 初始化配置
        self.load_config()
        
        # 初始化自动关机设置
        self.auto_shutdown_enabled = self.config.get('auto_shutdown_enabled', False)
        self.auto_shutdown_time = self.config.get('auto_shutdown_time', "22:00")
        self.auto_shutdown_notified = False  # 标记是否已经发送过通知
        
        # 设置定时器每分钟检查一次自动关机
        self.shutdown_timer = QTimer(self)
        self.shutdown_timer.timeout.connect(self.check_auto_shutdown)
        self.shutdown_timer.start(1000)  # 每1秒检查一次

        # 设置窗口属性
        self.setWindowFlags(
            Qt.FramelessWindowHint | 
            Qt.WindowStaysOnTopHint | 
            Qt.Tool |
            Qt.WindowDoesNotAcceptFocus  # 防止窗口获得焦点
        )
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedHeight(40)  # 固定高度为40像素
        self.setMinimumWidth(300)  # 最小宽度300像素
        
        # 设置透明度为95%
        self.setWindowOpacity(0.95)
        
        # 创建主布局
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(10, 0, 10, 0)  # 左右各10像素内边距
        
        # 创建左侧内容区域
        content_layout = QHBoxLayout()
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(10)  # 组件间距10像素
        
        # 创建三个独立的组件
        self.weekday_label = QLabel()
        self.schedule_label = QLabel()
        self.time_label = QLabel()
        
        # 设置字体
        font = QFont("Microsoft YaHei UI", 16)
        font.setBold(True)
        self.weekday_label.setFont(font)
        self.schedule_label.setFont(font)
        self.time_label.setFont(font)
        
        # 设置对齐方式
        self.weekday_label.setAlignment(Qt.AlignCenter)
        self.schedule_label.setAlignment(Qt.AlignCenter)
        self.time_label.setAlignment(Qt.AlignCenter)
        
        # 添加组件到内容布局
        content_layout.addWidget(self.weekday_label)
        content_layout.addWidget(self.schedule_label)
        content_layout.addWidget(self.time_label)
        
        # 添加内容布局到主布局
        main_layout.addLayout(content_layout)
        
        # 添加可伸缩空间使内容居中
        main_layout.addStretch(1)
        
        # 创建按钮布局
        buttons_layout = QVBoxLayout()
        buttons_layout.setContentsMargins(0, 0, 0, 0)
        buttons_layout.setSpacing(0)  # 按钮之间无间距
        
        # 随机点人按钮
        self.random_pick_btn = QPushButton("随机点人")
        self.random_pick_btn.setFixedHeight(20)  # 高度为20像素
        self.random_pick_btn.setFont(QFont("Microsoft YaHei UI", 10))  # 字体大小调整为10
        self.random_pick_btn.clicked.connect(self.random_pick_student)
        buttons_layout.addWidget(self.random_pick_btn)
        
        # 截图和更多按钮的容器
        bottom_buttons_layout = QHBoxLayout()
        bottom_buttons_layout.setContentsMargins(0, 0, 0, 0)
        bottom_buttons_layout.setSpacing(5)  # 按钮之间5间距
        
        # 截图按钮
        self.screenshot_btn = QPushButton("截图")
        self.screenshot_btn.setFixedHeight(20)  # 高度为20像素
        self.screenshot_btn.setFont(QFont("Microsoft YaHei UI", 10))  # 字体大小调整为10
        self.screenshot_btn.clicked.connect(self.simulate_screenshot)
        bottom_buttons_layout.addWidget(self.screenshot_btn)
        
        # 更多按钮
        self.more_btn = QPushButton("更多")
        self.more_btn.setFixedHeight(20)  # 高度为20像素
        self.more_btn.setFont(QFont("Microsoft YaHei UI", 10))  # 字体大小调整为10
        self.more_btn.clicked.connect(self.show_settings)
        bottom_buttons_layout.addWidget(self.more_btn)
        
        buttons_layout.addLayout(bottom_buttons_layout)
        
        # 添加按钮布局到主布局
        main_layout.addLayout(buttons_layout)
        
        # 设置样式
        self.apply_theme()
        
        # 设置初始大小和位置
        self.resize(300, 40)
        self.center_window()
        
        # 设置定时器每秒更新时间
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)
        
        # 初始化内容
        self.update_all_content()
        
        # 添加此行：在初始化后立即调整窗口大小
        QTimer.singleShot(100, self.adjust_window_size)
        
        # 设置窗口
        self.settings_window = None

        # 添加双击事件处理
        self.weekday_label.mouseDoubleClickEvent = self.toggle_window_size
        self.time_label.mouseDoubleClickEvent = self.show_fullscreen_time
        
        # 添加焦点事件处理
        self.focusInEvent = self.expand_window

        # 加载下课提醒设置
        self.class_reminders = self.config.get('class_reminders', "").split('、')
        self.class_reminders = [time.strip() for time in self.class_reminders if time.strip()]
        self.last_reminder_time = None  # 记录上次触发提醒的时间
        
        # 创建提醒覆盖层
        self.reminder_overlay = QLabel(self)
        self.reminder_overlay.setAlignment(Qt.AlignCenter)
        self.reminder_overlay.setFont(QFont("Microsoft YaHei UI", 16, QFont.Bold))
        self.reminder_overlay.setText(f"下课")
        self.reminder_overlay.setStyleSheet("background-color: #2ecc71; color: white;")
        self.reminder_overlay.hide()
        
        # 设置定时器每分钟检查提醒时间
        self.reminder_timer = QTimer(self)
        self.reminder_timer.timeout.connect(self.check_class_reminder)
        self.reminder_timer.start(1000)  # 1s检查一次
        
        # 加载考试设置
        self.exams = self.config.get('exams', [])

        self.contest_helper = None
        self.start_contest_helper()

        # 添加窗口状态标志
        self.is_collapsed = False

    def start_contest_helper(self):
        """启动考试助手程序"""
        try:
            # 获取当前目录
            current_dir = os.path.dirname(os.path.abspath(__file__))
            helper_path = os.path.join(current_dir, "contest_helper.exe")
            
            if os.path.exists(helper_path):
                self.contest_helper = subprocess.Popen([helper_path])
        except Exception as e:
            print(f"启动考试助手失败: {e}")
    
    def closeEvent(self, event):
        """关闭窗口时停止考试助手"""
        try:
            if self.contest_helper:
                self.contest_helper.terinate()
                self.contest_helper.wait(timeout=5)
        except Exception as e:
            print(f"停止考试助手失败: {e}")
        
        super().closeEvent(event)

    def check_class_reminder(self):
        """检查是否需要显示下课提醒"""
        current_time = datetime.datetime.now().strftime("%H:%M")
        
        # 如果当前时间在提醒列表中，并且与上次提醒时间不同
        if current_time in self.class_reminders and current_time != self.last_reminder_time:
            self.show_class_reminder()
            self.last_reminder_time = current_time  # 记录本次提醒时间
    
    def show_class_reminder(self):
        """显示下课提醒"""
        # 播放提示音
        self.play_reminder_sound()
        
        # 显示覆盖层
        self.reminder_overlay.setGeometry(0, 0, self.width(), self.height())
        self.reminder_overlay.raise_()
        self.reminder_overlay.show()
        
        # 添加动画效果
        self.animate_reminder()
        
        # 5秒后隐藏覆盖层
        QTimer.singleShot(5000, self.hide_class_reminder)
    
    def play_reminder_sound(self):
        """播放提醒音"""
        sound_path = os.path.join(os.path.dirname(__file__), "finish_class.wav")
        if os.path.exists(sound_path):
            try:
                # 使用系统默认播放器播放声音
                winsound.PlaySound(sound_path, winsound.SND_FILENAME | winsound.SND_ASYNC)
            except Exception as e:
                print(f"播放提示音失败: {e}")
    
    def animate_reminder(self):
        """添加提醒动画效果"""
        # 创建动画 - 淡入效果
        self.animation = QPropertyAnimation(self.reminder_overlay, b"windowOpacity")
        self.animation.setDuration(1000)  # 1秒
        self.animation.setStartValue(0.0)
        self.animation.setEndValue(1.0)
        self.animation.start()
    
    def hide_class_reminder(self):
        """隐藏下课提醒"""
        # 添加淡出动画
        self.animation = QPropertyAnimation(self.reminder_overlay, b"windowOpacity")
        self.animation.setDuration(1000)  # 1秒
        self.animation.setStartValue(1.0)
        self.animation.setEndValue(0.0)
        self.animation.finished.connect(self.reminder_overlay.hide)
        self.animation.start()
    
    def resizeEvent(self, event):
        """窗口大小变化时调整覆盖层大小"""
        super().resizeEvent(event)
        if self.reminder_overlay.isVisible():
            self.reminder_overlay.setGeometry(0, 0, self.width(), self.height())

    def toggle_window_size(self, event):
        """双击星期几标签时切换窗口大小"""
        if self.is_collapsed:
            self.expand_window()
        else:
            self.collapse_window()
    
    def collapse_window(self):
        """收缩窗口到屏幕顶部"""
        # 保存原始大小
        if not hasattr(self, 'original_size'):
            self.original_size = self.size()
        
        # 设置新大小（宽度不变，高度5像素）
        self.setFixedHeight(5)
        
        # 移动到屏幕顶部居中
        screen_geometry = QApplication.desktop().screenGeometry()
        x = (screen_geometry.width() - self.width()) // 2
        self.move(x, 0)
        
        # 设置标志
        self.is_collapsed = True
        
        # 设置窗口不接受鼠标事件（穿透）
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
    
    def expand_window(self, event=None):
        """展开窗口到正常大小"""
        if not self.is_collapsed:
            return
            
        # 恢复原始大小
        if hasattr(self, 'original_size'):
            self.setFixedHeight(self.original_size.height())
        
        # 移动到屏幕顶部居中
        screen_geometry = QApplication.desktop().screenGeometry()
        x = (screen_geometry.width() - self.width()) // 2
        self.move(x, 0)
        
        # 设置标志
        self.is_collapsed = False
        
        # 恢复鼠标事件
        self.setAttribute(Qt.WA_TransparentForMouseEvents, False)
    
    def show_fullscreen_time(self, event):
        """双击时间标签时显示全屏时间窗口"""
        # 创建全屏时间窗口
        self.time_window = QWidget()
        self.time_window.setWindowFlags(
            Qt.FramelessWindowHint | 
            Qt.WindowStaysOnTopHint
        )
        self.time_window.setStyleSheet("background-color: black;")
        self.time_window.setGeometry(0, 0, 1920, 1080)
        
        # 创建主布局
        layout = QVBoxLayout(self.time_window)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # 时间标签
        self.fullscreen_time_label = QLabel()
        self.fullscreen_time_label.setAlignment(Qt.AlignCenter)
        self.fullscreen_time_label.setFont(QFont("Microsoft YaHei UI", 200, QFont.Bold))
        self.fullscreen_time_label.setStyleSheet("color: white;")
        layout.addWidget(self.fullscreen_time_label, 1)
        
        # 退出按钮
        exit_btn = QPushButton("退出时间大屏")
        exit_btn.setFixedHeight(60)
        exit_btn.setFont(QFont("Microsoft YaHei UI", 24, QFont.Bold))
        exit_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                border: none;
                border-radius: 0px;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        exit_btn.clicked.connect(self.time_window.close)
        layout.addWidget(exit_btn, alignment=Qt.AlignLeft | Qt.AlignBottom)
        
        # 设置定时器每秒更新时间
        self.time_timer = QTimer(self.time_window)
        self.time_timer.timeout.connect(self.update_fullscreen_time)
        self.time_timer.start(1000)
        
        # 初始更新时间
        self.update_fullscreen_time()
        
        # 显示窗口
        self.time_window.showFullScreen()
    
    def update_fullscreen_time(self):
        """更新全屏时间显示"""
        current_time = datetime.datetime.now().strftime("%H:%M:%S")
        self.fullscreen_time_label.setText(current_time)
    
    def mousePressEvent(self, event):
        """鼠标点击事件处理"""
        if self.is_collapsed:
            self.expand_window()
        super().mousePressEvent(event)
    
    def focusInEvent(self, event):
        """获得焦点事件处理"""
        if self.is_collapsed:
            self.expand_window()
        super().focusInEvent(event)
    
    def check_auto_shutdown(self):
        """检查是否需要自动关机"""
        if not self.auto_shutdown_enabled:
            return
        
        # 获取当前时间
        now = datetime.datetime.now()
        current_time = now.strftime("%H:%M")
        
        # 解析关机时间
        try:
            shutdown_hour, shutdown_minute = map(int, self.auto_shutdown_time.split(':'))
            shutdown_time = datetime.time(shutdown_hour, shutdown_minute)
            
            # 计算提前10分钟的时间
            early_time = (datetime.datetime.combine(now.date(), shutdown_time) - 
                         datetime.timedelta(minutes=10)).time()
            early_time_str = early_time.strftime("%H:%M")
            
            # 检查是否是提前10分钟
            if current_time == early_time_str and not self.auto_shutdown_notified:
                # 弹出警告窗口
                self.show_shutdown_warning(shutdown_hour, shutdown_minute)
                self.auto_shutdown_notified = True
            
            # 检查是否到达关机时间
            if current_time == self.auto_shutdown_time:
                # 执行关机
                self.perform_auto_shutdown()
                
        except Exception as e:
            print(f"自动关机检查错误: {e}")
    
    def show_shutdown_warning(self, hour, minute):
        """显示关机警告（居中显示）"""
        warning_text = f"晚自习已下，将在10分钟后 ({hour:02d}:{minute:02d}) 自动关机"
        
        # 创建警告窗口
        warning_box = QMessageBox(self)
        warning_box.setWindowTitle("自动关机警告")
        warning_box.setText(warning_text)
        warning_box.setIcon(QMessageBox.Warning)
        warning_box.setStandardButtons(QMessageBox.Ok)
        
        # 在显示前调整位置
        warning_box.show()
        
        # 获取屏幕尺寸
        screen_geometry = QApplication.desktop().screenGeometry()
        
        # 获取窗口尺寸
        window_geometry = warning_box.frameGeometry()
        
        # 计算居中位置
        x = (screen_geometry.width() - window_geometry.width()) // 2
        y = (screen_geometry.height() - window_geometry.height()) // 2
        
        # 移动窗口到屏幕中央
        warning_box.move(x, y)
        
        # 显示模态对话框
        warning_box.exec_()
    
    def perform_auto_shutdown(self):
        """执行自动关机"""
        try:
            # 调用系统关机命令
            subprocess.Popen(["shutdown", "/f","/s", "/t", "0"])
        except Exception as e:
            print(f"自动关机失败: {e}")

    def load_config(self):
        """从config.yaml加载配置"""
        config_path = os.path.join(os.path.dirname(__file__), "config.yaml")
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                self.config = yaml.safe_load(f)
                self.schedule = self.config.get('schedule', {})
                self.dark_mode = self.config.get('darkmode', False)
                self.auto_shutdown_enabled = self.config.get('auto_shutdown_enabled', False)
                self.auto_shutdown_time = self.config.get('auto_shutdown_time', "22:00")
        except:
            self.config = {}
            self.schedule = {}
            self.dark_mode = False
            self.auto_shutdown_enabled = False
            self.auto_shutdown_time = "19:00"
    
    def save_config(self):
        """保存配置到文件"""
        config_path = os.path.join(os.path.dirname(__file__), "config.yaml")
        with open(config_path, 'w', encoding='utf-8') as f:
            yaml.dump(self.config, f)
    
    def apply_theme(self):
        """应用主题颜色"""
        palette = self.palette()
        
        if self.dark_mode:
            # 深色模式
            self.bg_color = QColor(30, 30, 30, 127)  # 50%透明度
            self.text_color = QColor(220, 220, 220)
            self.btn_bg_color = QColor(60, 60, 60)
            self.btn_text_color = QColor(220, 220, 220)  # 深色模式按钮文字颜色
        else:
            # 浅色模式
            self.bg_color = QColor(240, 240, 240, 127)  # 50%透明度
            self.text_color = QColor(40, 40, 40)
            self.btn_bg_color = QColor(200, 200, 200)
            self.btn_text_color = QColor(40, 40, 40)  # 浅色模式按钮文字颜色
        
        # 设置文字颜色
        palette.setColor(QPalette.WindowText, self.text_color)
        self.weekday_label.setPalette(palette)
        self.schedule_label.setPalette(palette)
        self.time_label.setPalette(palette)
        
        # 设置星期几为灰色
        weekday_palette = self.weekday_label.palette()
        weekday_palette.setColor(QPalette.WindowText, QColor(128, 128, 128))
        self.weekday_label.setPalette(weekday_palette)
        
        # 设置时间标签为非粗体
        time_font = self.time_label.font()
        time_font.setBold(False)
        self.time_label.setFont(time_font)
        
        # 设置按钮样式
        btn_style = f"""
            QPushButton {{
                background-color: {self.btn_bg_color.name()};
                color: {self.btn_text_color.name()};
                border: none;
                border-radius: 0px;
                padding: 0px;
            }}
            QPushButton:hover {{
                background-color: {self.btn_bg_color.lighter(120).name()};
            }}
        """
        self.random_pick_btn.setStyleSheet(btn_style)
        self.screenshot_btn.setStyleSheet(btn_style)
        self.more_btn.setStyleSheet(btn_style)
    
    def paintEvent(self, event):
        """绘制半透明背景"""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setBrush(QBrush(self.bg_color))
        painter.setPen(Qt.NoPen)
        painter.drawRect(self.rect())
    
    def center_window(self):
        """将窗口居中于屏幕顶部"""
        screen_geometry = QApplication.desktop().screenGeometry()
        x = (screen_geometry.width() - self.width()) // 2
        self.move(x, 0)
    
    def get_day_schedule(self):
        """获取今天的课程安排"""
        today = datetime.datetime.now().strftime("%A")
        chinese_days = {
            "Monday": "星期一",
            "Tuesday": "星期二",
            "Wednesday": "星期三",
            "Thursday": "星期四",
            "Friday": "星期五",
            "Saturday": "星期六",
            "Sunday": "星期日"
        }
        day_cn = chinese_days.get(today, today)
        
        # 周末处理
        if today in ["Saturday", "Sunday"]:
            return day_cn, "暂无课程"
        
        # 获取课程
        courses = self.schedule.get(day_cn, [])

        if not courses:
           return day_cn, "暂无课程"
        
        return day_cn, courses
    
    def update_all_content(self):
        """更新所有内容（只在日期变化或课表变化时调用）"""
        day_cn, schedule_text = self.get_day_schedule()
        
        # 更新星期几和课表内容
        self.weekday_label.setText(day_cn)
        self.schedule_label.setText(schedule_text)
        
        # 更新时间
        self.update_time()
    
    def update_time(self):
        """只更新时间内容（不调整窗口大小）"""
        current_time = datetime.datetime.now().strftime("%H:%M")
        self.time_label.setText(current_time)
    
    def adjust_window_size(self):
        """调整窗口大小（只在内容变化时调用）"""
        # 确保标签大小正确更新
        self.weekday_label.adjustSize()
        self.schedule_label.adjustSize()
        self.time_label.adjustSize()
        
        # 计算窗口总宽度
        total_width = (
            self.weekday_label.width() +  # 星期几标签宽度
            self.schedule_label.width() +  # 课表标签宽度
            self.time_label.width() +     # 时间标签宽度
            # self.random_pick_btn.width() + # 随机点人按钮宽度
            self.screenshot_btn.width() + # 截图按钮宽度
            self.more_btn.width() +       # 更多按钮宽度
            40 +  # 左右边距各10像素，共20像素
            15     # 三个间距各10像素，共30像素
        )
        
        # 确保最小宽度为300像素
        new_width = max(total_width, 300)
        self.setFixedWidth(new_width)  # 固定宽度，禁止调整大小
        self.center_window()
    
    def random_pick_student(self):
        """随机点名学生"""
        # 创建随机点人窗口
        self.random_picker_window = RandomPickerWindow(self)
        
        # 居中显示窗口
        screen_geometry = QApplication.desktop().screenGeometry()
        x = (screen_geometry.width() - self.random_picker_window.width()) // 2
        y = (screen_geometry.height() - self.random_picker_window.height()) // 2
        self.random_picker_window.move(x, y)
        
        self.random_picker_window.show()
    
    def simulate_screenshot(self):
        """模拟按下 Ctrl+Alt+A 截图快捷键，并隐藏窗口1秒"""
        # 隐藏窗口
        self.hide()
        
        # 使用单次定时器在1秒后恢复窗口
        QTimer.singleShot(1000, self.show)
        
        # 模拟截图快捷键
        try:
            # 使用小延迟确保窗口已隐藏
            QTimer.singleShot(100, self.send_screenshot_keys)
        except Exception as e:
            print(f"模拟截图快捷键失败: {e}")
    
    def send_screenshot_keys(self):
        """实际发送截图快捷键"""
        try:
            # 模拟按键
            keyboard.press('ctrl')
            keyboard.press('alt')
            keyboard.press('a')
            keyboard.release('a')
            keyboard.release('alt')
            keyboard.release('ctrl')
        except Exception as e:
            print(f"发送截图快捷键失败: {e}")
    
    def show_settings(self):
        """显示设置窗口"""
        if not self.settings_window:
            self.settings_window = SettingsWindow(self)
        
        # 居中显示设置窗口
        screen_geometry = QApplication.desktop().screenGeometry()
        x = (screen_geometry.width() - self.settings_window.width()) // 2
        y = (screen_geometry.height() - self.settings_window.height()) // 2
        self.settings_window.move(x, y)
        
        # 默认显示临时改课页面（不再强制切换到检查更新页面）
        self.settings_window.show()
    
    def verify_password(self, password):
        """验证密码"""
        if 'password' not in self.config:
            return True  # 没有设置密码，直接通过
        
        # 获取加密的密码和盐值
        encrypted_data = self.config['password']
        salt_b64, encrypted_password = encrypted_data.split(':')
        salt = base64.b64decode(salt_b64.encode())
        
        # 使用PBKDF2HMAC算法生成密钥
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=salt,
            iterations=100000,
            backend=default_backend()
        )
        key = base64.urlsafe_b64encode(kdf.derive(password.encode()))
        
        # 使用Fernet解密密码
        f = Fernet(key)
        try:
            decrypted_password = f.decrypt(encrypted_password.encode()).decode()
            return decrypted_password == password
        except:
            return False
    
    def exit_application(self):
        """退出应用程序"""
        # 检查是否设置了密码
        if 'password' in self.config:
            # 弹出密码输入对话框
            password, ok = QInputDialog.getText(
                self, 
                "密码验证", 
                "请输入管理密码:", 
                QLineEdit.Password
            )
            
            if not ok:
                return  # 用户取消
                
            # 验证密码
            if not self.verify_password(password):
                QMessageBox.critical(self, "密码错误", "输入的管理密码不正确")
                return
        
        # 退出程序
        subprocess.run(["taskkill","/f","/im","contest_helper.exe"])
        QApplication.quit()
    
    def closeEvent(self, event):
        """禁用关闭事件"""
        event.ignore()  # 忽略关闭事件
    
    def keyPressEvent(self, event):
        """禁用键盘关闭功能"""
        # 忽略Alt+F4组合键
        if event.key() == Qt.Key_F4 and event.modifiers() == Qt.AltModifier:
            event.ignore()
        # 忽略Esc键
        elif event.key() == Qt.Key_Escape:
            event.ignore()
        else:
            super().keyPressEvent(event)

def check_for_updates():
    """在程序启动时检查更新（根据配置）"""
    try:
        # 加载配置文件
        config_path = os.path.join(os.path.dirname(__file__), "config.yaml")
        if not os.path.exists(config_path):
            return
            
        with open(config_path, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
        
        # 检查是否启用自动更新
        auto_update = config.get('autoUpdate', False)
        if not auto_update:
            return
            
        # 获取当前版本
        current_version = config.get('version', '1.0')
        
        # 获取最新版本
        version_url = "http://spj2025.top:19540/apps/version/pyclasstool.txt"
        response = requests.get(version_url)
        if response.status_code != 200:
            return
            
        latest_version = response.content.decode('gbk').strip()
        
        # 检查是否有新版本
        if latest_version != current_version:
            # 启动更新程序
            update_script = os.path.join(os.path.dirname(__file__), "update.pyw")
            if os.path.exists(update_script):
                subprocess.Popen(["python", update_script])
                # 退出当前程序
                sys.exit(0)
    except Exception as e:
        print(f"启动更新检查失败: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    check_for_updates()
    window = ScheduleWindow()
    window.show()
    sys.exit(app.exec_())
