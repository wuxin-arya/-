#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Copyright (c) 2025 lzy. All rights reserved.
"""
Project: 智能考勤系统
Description: TTS点名，保护嗓子
             使用方式：添加班级，导入学生名单（姓名，学号）
                     开始考勤，选择出勤、旷课、请假
                     data目录下students文件为学生名单，attendance文件为考勤细节，stats文件为统计
"""


import os
import pandas as pd
import pyttsx3
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QListWidget, QFileDialog, QMessageBox,
    QInputDialog
)
from PyQt5.QtCore import QThread, pyqtSignal
from datetime import datetime, time
import re


# ================= 数据管理模块 =================
class DataManager:
    def __init__(self):
        self.data_dir = "data"
        # 强制创建数据目录（修复权限问题）
        os.makedirs(self.data_dir, exist_ok=True, mode=0o777)

        # 初始化空白索引文件（如果不存在）
        self.classes_path = os.path.join(self.data_dir, "classes.xlsx")
        if not os.path.exists(self.classes_path):
            pd.DataFrame(columns=["班级名称", "学生名单文件", "考勤文件", "统计文件"]
                         ).to_excel(self.classes_path, index=False)

    def load_classes(self):
        """修复数据加载类型问题[6](@ref)"""
        if os.path.exists(self.classes_path):
            df = pd.read_excel(self.classes_path)
            # 强制转换字段类型为字符串
            df["班级名称"] = df["班级名称"].astype(str)
            return df.to_dict('records')
        return []

    def save_class(self, class_name):
        """修复文件创建逻辑（关键修正点）"""
        # 清洗文件名（处理特殊字符）
        base_name = re.sub(r'[\\/*?:"<>|]', '', class_name.strip())
        base_name = base_name.replace(" ", "_")  # 空格转下划线

        # 定义文件路径（增加路径存在性检查）
        files = {
            "students": os.path.join(self.data_dir, f"{base_name}_students.xlsx"),
            "attendance": os.path.join(self.data_dir, f"{base_name}_attendance.xlsx"),
            "stats": os.path.join(self.data_dir, f"{base_name}_stats.xlsx")
        }

        # 强制创建文件（包含表头）
        try:
            for file_type, path in files.items():
                os.makedirs(os.path.dirname(path), exist_ok=True, mode=0o777)
                if not os.path.exists(path):
                    pd.DataFrame(columns=self._get_headers(file_type)).to_excel(path, index=False)
        except Exception as e:
            raise RuntimeError(f"文件创建失败: {str(e)}")


        # 更新索引文件（增加重复检查）
        df = pd.read_excel(self.classes_path)
        if class_name in df["班级名称"].values:
            raise ValueError("班级名称已存在")

        new_row = pd.DataFrame([{
            "班级名称": class_name,
            "学生名单文件": files["students"],
            "考勤文件": files["attendance"],
            "统计文件": files["stats"]
        }])
        updated_df = pd.concat([df, new_row], ignore_index=True)
        updated_df.to_excel(self.classes_path, index=False)
        return files

    def _get_headers(self, file_type):
        """统一管理文件表头"""
        headers = {
            "students": ["学号", "姓名"],
            "attendance": ["学号", "姓名", "状态", "日期", "时间"],
            "stats": ["学号", "姓名", "出勤", "旷课", "请假"]
        }
        return headers[file_type]


# ================= 语音播报模块 =================
class VoiceThread(QThread):
    finished = pyqtSignal()

    def __init__(self, name):
        super().__init__()
        self.name = name

    def run(self):
        """独立语音引擎实例避免冲突[1](@ref)"""
        try:
            engine = pyttsx3.init()
            engine.setProperty('rate', 150)  # 设置语速[3](@ref)
            engine.setProperty('volume', 0.8)  # 设置音量[2](@ref)
            engine.say(f"{self.name}")
            engine.runAndWait()
        except Exception as e:
            print(f"语音异常: {str(e)}")
        finally:
            self.finished.emit()


# ================= 主界面模块 =================
class AttendanceSystem(QMainWindow):
    def __init__(self):
        super().__init__()
        self.data_manager = DataManager()
        self.classes = self.data_manager.load_classes()
        self.current_index = 0
        self.current_class = None
        self.current_files = {}
        self.students = []
        self.initUI()
        self.data_dir = "data"

    def initUI(self):
        self.setWindowTitle('智能考勤系统')
        self.setGeometry(300, 300, 800, 600)

        # 控件初始化
        self.class_list = QListWidget()
        self.class_list.addItems([c["班级名称"] for c in self.classes])

        btn_new = QPushButton('新建班级')
        btn_delete = QPushButton('删除班级')
        btn_import = QPushButton('导入名单')
        btn_start = QPushButton('开始考勤')
        btn_replay = QPushButton('重复播报')
        # btn_export = QPushButton('导出报表')

        self.status_label = QLabel('当前班级：无')
        self.student_label = QLabel('当前学生：无')

        btn_attend = QPushButton('出勤')
        btn_absent = QPushButton('旷课')
        btn_leave = QPushButton('请假')

        # 布局设置
        left_panel = QVBoxLayout()
        left_panel.addWidget(btn_new)
        left_panel.addWidget(btn_delete)
        left_panel.addWidget(btn_import)
        left_panel.addWidget(self.class_list)

        right_panel = QVBoxLayout()
        right_panel.addWidget(self.status_label)
        right_panel.addWidget(self.student_label)
        right_panel.addWidget(btn_start)
        right_panel.addWidget(btn_replay)
        right_panel.addWidget(btn_attend)
        right_panel.addWidget(btn_absent)
        right_panel.addWidget(btn_leave)
        # right_panel.addWidget(btn_export)

        main_layout = QHBoxLayout()
        main_layout.addLayout(left_panel, 1)
        main_layout.addLayout(right_panel, 2)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        # 信号连接
        btn_new.clicked.connect(self.create_class)
        btn_import.clicked.connect(self.import_students)
        btn_delete.clicked.connect(self.delete_class)
        btn_start.clicked.connect(self.start_attendance)
        btn_replay.clicked.connect(self.replay_name)  # 修复重复播报按钮
        btn_attend.clicked.connect(lambda: self.record_status('出勤'))
        btn_absent.clicked.connect(lambda: self.record_status('旷课'))
        btn_leave.clicked.connect(lambda: self.record_status('请假'))
        # btn_export.clicked.connect(self.export_report)
        self.class_list.itemClicked.connect(self.select_class)

    # ================= 核心功能方法 =================
    def create_class(self):
        class_name, ok = QInputDialog.getText(self, '新建班级', '输入班级名称：')
        if not ok or not class_name:
            return

        try:
            # 调用修正后的保存方法（关键新增）
            self.data_manager.save_class(class_name)

            # 更新界面
            self.class_list.addItem(class_name)
            self.classes = self.data_manager.load_classes()
            QMessageBox.information(self, "成功", f"班级 [{class_name}] 创建成功")

        except ValueError as e:
            QMessageBox.warning(self, "警告", str(e))
        except Exception as e:
            QMessageBox.critical(self, "错误", f"创建失败: {str(e)}")

    def select_class(self, item):
        self.current_class = item.text()
        for cls in self.classes:
            if cls["班级名称"] == self.current_class:
                self.current_files = {
                    "students": cls["学生名单文件"],
                    "attendance": cls["考勤文件"],
                    "stats": cls["统计文件"]
                }
                break
        self.status_label.setText(f'当前班级：{self.current_class}')

    def delete_class(self):
        try:
            if not self.current_class:
                QMessageBox.warning(self, "警告", "请先选择要删除的班级")
                return

            # 二次确认对话框
            confirm = QMessageBox.question(
                self, "确认删除",
                f"确定要永久删除班级【{self.current_class}】及其所有数据吗？",
                QMessageBox.Yes | QMessageBox.No
            )
            if confirm != QMessageBox.Yes:
                return

            # 创建备份目录（按日期归档）
            backup_dir = os.path.join(
                self.data_manager.data_dir,
                "deleted_classes",
                datetime.now().strftime("%Y%m%d")
            )
            os.makedirs(backup_dir, exist_ok=True)

            # 备份并删除文件（带异常处理）
            deleted_files = []
            for file_type in ["students", "attendance", "stats"]:
                src = self.current_files.get(file_type, "")
                if src and os.path.exists(src):
                    # 生成带时间戳的备份文件名
                    timestamp = datetime.now().strftime("%H%M%S")
                    dst = os.path.join(backup_dir, f"{timestamp}_{os.path.basename(src)}")
                    os.rename(src, dst)
                    deleted_files.append(dst)

            # 更新班级索引文件
            class_index_path = self.data_manager.classes_path
            if os.path.exists(class_index_path):
                df = pd.read_excel(class_index_path)
                # 精确匹配当前班级（考虑前后空格问题）
                current_class = str(self.current_class).strip()
                df["班级名称"] = df["班级名称"].astype(str).str.strip()
                df = df[df["班级名称"] != current_class]
                df.to_excel(class_index_path, index=False)

            # 更新界面显示
            current_row = self.class_list.currentRow()
            if current_row >= 0:
                self.class_list.takeItem(current_row)
                # 重置当前选择状态
                self.current_class = None
                self.current_files = {}
                self.students = []

            # 显示操作结果
            result_msg = f"已删除班级：{self.current_class}\n"
            result_msg += f"备份文件保存在：\n{backup_dir}"
            QMessageBox.information(self, "删除成功", result_msg)

        except PermissionError as e:
            QMessageBox.critical(self, "权限错误",
                                 f"文件被其他程序占用\n{e.filename}\n请关闭Excel或其他正在使用文件的程序")
        except Exception as e:
            QMessageBox.critical(self, "删除失败",
                                 f"操作未完成\n错误类型：{type(e).__name__}\n错误详情：{str(e)}")

    def start_attendance(self):
        """启动考勤流程"""
        if not self.current_class:
            QMessageBox.warning(self, "警告", "请先选择班级")
            return

        try:
            self.students = pd.read_excel(self.current_files["students"]).to_dict('records')
            self.current_index = 0
            self.update_student_display()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"数据加载失败：{str(e)}")

    def replay_name(self):
        """修复重复播报功能[1](@ref)"""
        if 0 <= self.current_index < len(self.students):
            current_student = self.students[self.current_index]["姓名"]
            self.voice_thread = VoiceThread(current_student)  # 播报当前学生
            self.voice_thread.start()

    def record_status(self, status):
        if not self.current_files.get("attendance"):
            QMessageBox.critical(self, "错误", "未选择有效班级")
            return

        try:
            # 记录考勤
            student = self.students[self.current_index]
            timestamp = pd.Timestamp.now()

            record = {
                "学号": student["学号"],
                "姓名": student["姓名"],
                "状态": status,
                "日期": timestamp.strftime('%Y-%m-%d'),
                "时间": timestamp.strftime('%H:%M:%S')
            }

            # 更新考勤文件
            # 读取现有数据
            existing = pd.read_excel(self.current_files["attendance"])
            # 合并新记录
            updated = pd.concat([existing, pd.DataFrame([record])], ignore_index=True)
            # 覆盖写入
            updated.to_excel(self.current_files["attendance"], index=False)

            # 更新统计
            stats_df = pd.read_excel(self.current_files["stats"])
            stats_df.loc[stats_df['学号'] == student["学号"], status] += 1
            stats_df.to_excel(self.current_files["stats"], index=False)

            # 切换下个学生
            self.current_index += 1
            if self.current_index < len(self.students):
                self.update_student_display()
            else:
                QMessageBox.information(self, "完成", "本班考勤已完成")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"记录失败：{str(e)}")

    def update_student_display(self):
        """更新学生显示并播报"""
        current_student = self.students[self.current_index]["姓名"]
        self.student_label.setText(f'当前学生：{current_student}')
        self.replay_name()  # 自动播报新学生


    def import_students(self):
        """改进版学生名单导入方法（支持异常处理和格式校验）"""
        try:
            # 获取文件路径
            file_path, _ = QFileDialog.getOpenFileName(
                self, "选择学生名单", "",
                "Excel文件 (*.xlsx *.xls);;CSV文件 (*.csv)",
                options=QFileDialog.DontUseNativeDialog
            )
            if not file_path:
                return

            # 读取前验证文件存在性
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"文件不存在: {file_path}")

            # 动态处理不同文件格式
            if file_path.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file_path, engine='openpyxl')
            elif file_path.endswith('.csv'):
                df = pd.read_csv(file_path, encoding='utf-8-sig')
            else:
                raise ValueError("不支持的格式")

            # 关键字段校验
            required_columns = {'学号', '姓名'}
            if not required_columns.issubset(df.columns):
                missing = required_columns - set(df.columns)
                raise ValueError(f"缺少必要字段: {', '.join(missing)}")

            # 数据清洗
            df = df[['学号', '姓名']].dropna().drop_duplicates()
            if df.empty:
                raise ValueError("数据为空或格式错误")

            # 保存到班级文件
            if not self.current_class:
                raise RuntimeError("请先选择或创建班级")

            base_name = self.current_class.strip().replace(' ', '_')
            target_path = os.path.join(
                self.data_manager.data_dir,
                f"{base_name}_students.xlsx"
            )

            # 写入前创建目录
            os.makedirs(os.path.dirname(target_path), exist_ok=True)
            df.to_excel(target_path, index=False)

            # 更新索引文件
            classes_df = pd.read_excel(self.data_manager.classes_path)
            mask = classes_df['班级名称'] == self.current_class
            classes_df.loc[mask, '学生名单文件'] = target_path
            classes_df.to_excel(self.data_manager.classes_path, index=False)

            QMessageBox.information(self, "成功",
                                    f"已导入{len(df)}条学生记录\n保存路径: {target_path}")

        except Exception as e:
            error_msg = f"导入失败: {str(e)}\n建议检查："
            if isinstance(e, FileNotFoundError):
                error_msg += "\n1. 文件是否被移动或删除\n2. 路径是否包含特殊字符"
            elif isinstance(e, ValueError):
                error_msg += "\n1. 文件是否包含学号、姓名列\n2. 数据是否存在空行"
            elif isinstance(e, PermissionError):
                error_msg += "\n1. 文件是否被其他程序打开\n2. 是否有写入权限"

            QMessageBox.critical(self, "错误", error_msg)

        files = {
            "students": os.path.join(self.data_dir, f"{base_name}_students.xlsx"),
            "attendance": os.path.join(self.data_dir, f"{base_name}_attendance.xlsx"),
            "stats": os.path.join(self.data_dir, f"{base_name}_stats.xlsx")
        }

        # DataManager.save_class()中初始化统计文件时：
        stats_df = pd.DataFrame(columns=["学号", "姓名", "出勤", "旷课", "请假"])

        # 遍历学生名单填充初始值（修复append问题）
        students_df = pd.read_excel(files['students'])
        for _, student in students_df.iterrows():
            new_row = pd.DataFrame([{
                "学号": student["学号"],
                "姓名": student["姓名"],
                "出勤": 0,
                "旷课": 0,
                "请假": 0
            }])
            stats_df = pd.concat([stats_df, new_row], ignore_index=True)

        stats_df.to_excel(files["stats"], index=False)


if __name__ == '__main__':
    app = QApplication([])
    window = AttendanceSystem()
    window.show()
    app.exec_()
