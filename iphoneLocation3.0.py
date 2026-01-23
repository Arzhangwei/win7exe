#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations  # 解决Python 3.8的类型注解兼容性问题

"""
读取 peopleList.csv → 生成步步高 USB 电话通讯录专用 CSV
优化功能：支持客户经理列，输出格式可配置
打包：pyinstaller -F -w bbk_csv_tool.py
用于GitHub action编译 win7 exe
"""
import os
import sys
import csv
import pandas as pd
from phone import Phone
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout,
                             QPushButton, QTextEdit, QMessageBox,
                             QLabel, QHBoxLayout, QCheckBox, QGroupBox,
                             QGridLayout)
from PyQt5.QtCore import QCoreApplication, Qt

# ----------------------------------------------------------
# 路径工具：打包后 exe 同目录，开发时脚本目录
# ----------------------------------------------------------
def current_dir() -> str:
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.realpath(__file__))

def res_path(*names) -> str:
    full = os.path.join(current_dir(), *names)
    # 只创建目录，不创建文件
    if names:
        os.makedirs(os.path.dirname(full), exist_ok=True)
    return full

# ----------------------------------------------------------
# 文件常量
# ----------------------------------------------------------
CSV_FILE   = res_path('peopleList.csv')     # 输入
DAT_PATH   = res_path('phone.dat')          # 归属地库
VCARD_CSV  = res_path('名片.csv')           # 输出

BBK_HEADER = ["姓名", "移动电话", "办公电话", "家庭电话", "备注"]

# ----------------------------------------------------------
# 写 CSV 工具：GB18030 无 BOM
# ----------------------------------------------------------
def write_bbk_csv(path, rows: list[list[str]]):
    try:
        with open(path, 'w', newline='', encoding='gb18030') as f:
            writer = csv.writer(f)
            writer.writerow(BBK_HEADER)
            writer.writerows(rows)
        return True
    except Exception as e:
        raise Exception(f"写入文件失败: {str(e)}")

# ----------------------------------------------------------
# Qt 界面
# ----------------------------------------------------------
class MainWin(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('利辛农商行电话回访辅助（CSV→步步高）')
        self.resize(600, 450)
        self.is_processing = False  # 防止重复点击

        # 创建控件
        self.create_widgets()
        
        # 设置布局
        self.setup_layout()
        
        # 连接信号
        self.run_btn.clicked.connect(self.work)

        # 检查必要文件
        self.check_required_files()

    def create_widgets(self):
        """创建所有界面控件"""
        # 配置选项组
        self.config_group = QGroupBox('输出配置')
        
        # 配置复选框
        self.include_manager_cb = QCheckBox('包含客户经理姓名')
        self.include_manager_cb.setChecked(True)
        self.include_manager_cb.setToolTip('在姓名中显示客户经理姓名')
        
        self.include_phone_cb = QCheckBox('包含手机号码')
        self.include_phone_cb.setChecked(True)
        self.include_phone_cb.setToolTip('在姓名中显示手机号码')
        
        self.include_seq_cb = QCheckBox('包含序号')
        self.include_seq_cb.setChecked(True)
        self.include_seq_cb.setToolTip('在姓名前添加序号（001. 格式）')
        
        # 格式化示例标签
        self.format_example_label = QLabel()
        self.update_format_example()
        
        # 连接复选框信号，更新示例
        self.include_manager_cb.stateChanged.connect(self.update_format_example)
        self.include_phone_cb.stateChanged.connect(self.update_format_example)
        self.include_seq_cb.stateChanged.connect(self.update_format_example)
        
        # 处理按钮
        self.run_btn = QPushButton('开始处理')
        self.run_btn.setFixedHeight(40)
        self.run_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        
        # 日志文本框
        self.log = QTextEdit(readOnly=True)
        self.log.setPlaceholderText('处理日志将显示在这里...')
        self.log.setStyleSheet("""
            QTextEdit {
                background-color: #f5f5f5;
                border: 1px solid #ddd;
                border-radius: 3px;
                padding: 5px;
            }
        """)

    def setup_layout(self):
        """设置界面布局"""
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        
        # 配置组布局
        config_layout = QGridLayout()
        config_layout.addWidget(self.include_manager_cb, 0, 0)
        config_layout.addWidget(self.include_phone_cb, 0, 1)
        config_layout.addWidget(self.include_seq_cb, 1, 0)
        config_layout.addWidget(QLabel('格式示例:'), 1, 1)
        config_layout.addWidget(self.format_example_label, 1, 2)
        config_layout.setColumnStretch(2, 1)
        
        self.config_group.setLayout(config_layout)
        
        # 添加部件到主布局
        main_layout.addWidget(self.config_group)
        main_layout.addWidget(self.run_btn)
        main_layout.addWidget(QLabel('处理日志:'))
        main_layout.addWidget(self.log, 1)  # 1表示可拉伸

    def update_format_example(self):
        """更新格式示例"""
        include_manager = self.include_manager_cb.isChecked()
        include_phone = self.include_phone_cb.isChecked()
        include_seq = self.include_seq_cb.isChecked()
        
        # 构建示例文本
        example_parts = []
        
        if include_seq:
            example_parts.append("001.")
        
        example_parts.append("客户姓名")

        if include_phone:
            example_parts.append("_13800000000")
        
        if include_manager:
            example_parts.append("_经办:张三")
        
        
        
        example = "".join(example_parts) if example_parts else "客户姓名"
        
        # 担保人示例
        guarantor_example = example.replace("客户姓名", "客户姓名_担保")
        
        # 设置标签文本
        self.format_example_label.setText(f"客户: {example}<br>担保人: {guarantor_example}")
        self.format_example_label.setStyleSheet("color: #666; font-style: italic;")

    def check_required_files(self):
        """检查必要文件是否存在"""
        missing_files = []
        if not os.path.exists(DAT_PATH):
            missing_files.append(f'phone.dat - 归属地查询库')
        if not os.path.exists(CSV_FILE):
            missing_files.append(f'peopleList.csv - 数据源文件')
        
        if missing_files:
            msg = "缺少以下必要文件，请将它们放在程序目录下：\n\n"
            for f in missing_files:
                msg += f"• {f}\n"
            msg += f"\n程序目录：{current_dir()}"
            QMessageBox.warning(self, '文件缺失', msg)

    # -------------- 工具 --------------
    def need_prefix(self, mobile: str) -> bool:
        """检查是否需要添加长途前缀"""
        try:
            # 清理手机号
            mobile_clean = ''.join(filter(str.isdigit, str(mobile)))
            if not mobile_clean:
                return False
                
            info = Phone(dat_file=DAT_PATH).find(mobile_clean) or {}
            return info.get('area_code') != '0558' or info.get('city') != '亳州'
        except Exception:
            # 如果查询失败，默认不加前缀
            return False

    def format_name(self, name: str, mobile: str, role: str, manager: str = "", 
                    include_manager: bool = True, include_phone: bool = True, 
                    include_seq: bool = True, seq: int = 0) -> str:
        """格式化姓名，根据配置包含序号、角色、客户经理和手机号"""
        parts = []
        
        # 添加序号
        if include_seq:
            parts.append(f"{seq:03d}.")
        

        # 添加姓名和角色
        if role == '担保人':
            parts.append(f"{name}_担保")
        else:
            parts.append(name)

        # 添加手机号
        if include_phone:
            parts.append(f"_{mobile}")
        
        # 添加客户经理
        if include_manager and manager and manager.strip():
            parts.append(f"_经办:{manager}")
        
        
        
        return "".join(parts)

    # -------------- 主流程 --------------
    def work(self):
        """主处理函数"""
        if self.is_processing:
            return
            
        self.is_processing = True
        self.run_btn.setEnabled(False)
        self.run_btn.setText("处理中...")
        self.log.clear()
        QCoreApplication.processEvents()  # 更新界面
        
        try:
            if not os.path.exists(CSV_FILE):
                QMessageBox.critical(self, '错误', f'当前目录未找到：\n{CSV_FILE}')
                return
            
            # 获取配置
            include_manager = self.include_manager_cb.isChecked()
            include_phone = self.include_phone_cb.isChecked()
            include_seq = self.include_seq_cb.isChecked()
            
            try:
                # 尝试多种编码格式读取CSV
                df = self.read_csv_with_encodings()
                
                # 标准化列名（处理可能的空格或不同命名）
                df.columns = df.columns.str.strip()
                
                # 打印列名用于调试
                self.log.append(f"检测到CSV列名：{df.columns.tolist()}")
                
            except Exception as e:
                QMessageBox.critical(self, '读取CSV失败', 
                                   f'文件编码无法识别或格式错误。\n请确保文件使用UTF-8或GBK编码保存。\n\n错误详情:\n{str(e)}')
                return

            # 处理数据
            records = self.process_dataframe(df, include_manager)
            
            if not records:
                QMessageBox.information(self, '提示', '未找到任何有效手机号')
                return

            # 写步步高 CSV
            success = self.write_output_csv(records, include_manager, include_phone, include_seq)
            if not success:
                return

            # 显示日志
            self.display_log(records, include_manager, include_phone, include_seq)
            
            QMessageBox.information(self, '处理完成', 
                                  f'成功生成 {len(records)} 条记录\n输出文件：{VCARD_CSV}')
            
        except Exception as e:
            QMessageBox.critical(self, '处理异常', f'处理过程中出现异常：\n{str(e)}')
            self.log.append(f"错误: {str(e)}")
        finally:
            self.is_processing = False
            self.run_btn.setEnabled(True)
            self.run_btn.setText("开始处理")

    def read_csv_with_encodings(self):
        """尝试多种编码读取CSV文件"""
        encodings_to_try = ['utf-8', 'utf-8-sig', 'gbk', 'gb18030']
        df = None
        last_error = None

        for enc in encodings_to_try:
            try:
                df = pd.read_csv(CSV_FILE, dtype=str, encoding=enc).fillna('')
                self.log.append(f"成功使用 {enc} 编码读取文件")
                return df
            except UnicodeDecodeError as e:
                last_error = e
                continue
            except Exception as e:
                last_error = e
                continue

        if df is None:
            raise last_error if last_error else Exception("无法读取CSV文件")

    def process_dataframe(self, df: pd.DataFrame, include_manager: bool) -> list:
        """处理DataFrame，提取记录"""
        records = []
        
        # 可能的列名变体
        possible_columns = {
            '客户姓名': ['客户姓名', '客户名称', '姓名'],
            '客户经理': ['客户经理', '经理', '管户经理', '责任人'],
            '手机号码': ['手机号码', '手机号', '电话', '联系电话'],
            '担保人姓名': ['担保人姓名', '担保人', '共同借款人'],
            '担保人手机': ['手机号码.1', '担保人手机', '担保人电话', '联系电话.1']
        }
        
        # 映射实际列名
        actual_columns = {}
        for key, variants in possible_columns.items():
            for variant in variants:
                if variant in df.columns:
                    actual_columns[key] = variant
                    break
        
        self.log.append(f"列映射结果: {actual_columns}")
        
        for idx, line in df.iterrows():
            # 处理客户记录
            client_name_col = actual_columns.get('客户姓名')
            manager_col = actual_columns.get('客户经理')
            mobile_col = actual_columns.get('手机号码')
            
            if client_name_col and mobile_col:
                name = str(line.get(client_name_col, '')).strip()
                mobile = str(line.get(mobile_col, '')).strip()
                manager = str(line.get(manager_col, '')) if manager_col else ""
                
                if mobile and mobile != 'nan' and len(mobile) >= 7:  # 放宽长度限制
                    # 清理手机号，只保留数字
                    mobile_clean = ''.join(filter(str.isdigit, mobile))
                    if len(mobile_clean) >= 7:  # 至少7位数字
                        if self.need_prefix(mobile_clean):
                            mobile_clean = '0' + mobile_clean
                        records.append({
                            '原始姓名': name,
                            '手机号': mobile_clean,
                            '角色': '客户',
                            '客户经理': manager.strip() if manager else "",
                            '行号': idx + 1  # 添加行号便于调试
                        })
            
            # 处理担保人记录
            guarantor_name_col = actual_columns.get('担保人姓名')
            guarantor_mobile_col = actual_columns.get('担保人手机')
            
            if guarantor_name_col and guarantor_mobile_col:
                name = str(line.get(guarantor_name_col, '')).strip()
                mobile = str(line.get(guarantor_mobile_col, '')).strip()
                
                if mobile and mobile != 'nan' and len(mobile) >= 7:
                    # 清理手机号，只保留数字
                    mobile_clean = ''.join(filter(str.isdigit, mobile))
                    if len(mobile_clean) >= 7:
                        if self.need_prefix(mobile_clean):
                            mobile_clean = '0' + mobile_clean
                        # 担保人使用客户的客户经理信息
                        manager = str(line.get(manager_col, '')) if manager_col else ""
                        records.append({
                            '原始姓名': name,
                            '手机号': mobile_clean,
                            '角色': '担保人',
                            '客户经理': manager.strip() if manager else "",
                            '行号': idx + 1
                        })
        
        self.log.append(f"成功提取 {len(records)} 条有效记录")
        return records

    def write_output_csv(self, records: list, include_manager: bool, include_phone: bool, include_seq: bool) -> bool:
        """写入输出CSV文件"""
        try:
            # 准备步步高CSV数据
            bbk_rows = []
            for i, record in enumerate(records, 1):
                formatted_name = self.format_name(
                    record['原始姓名'],
                    record['手机号'],
                    record['角色'],
                    record['客户经理'],
                    include_manager,
                    include_phone,
                    include_seq,
                    i
                )
                bbk_rows.append([formatted_name, record['手机号'], '', '', ''])
            
            write_bbk_csv(VCARD_CSV, bbk_rows)
            self.log.append(f"成功写入文件: {VCARD_CSV}")
            return True
            
        except PermissionError:
            QMessageBox.warning(self, '文件被占用',
                              f'请先关闭正在打开的文件再试：\n{VCARD_CSV}')
            return False
        except Exception as e:
            QMessageBox.critical(self, '写CSV失败', str(e))
            self.log.append(f"写入失败: {str(e)}")
            return False

    def display_log(self, records: list, include_manager: bool, include_phone: bool, include_seq: bool):
        """显示处理日志"""
        text_lines = []
        
        for i, record in enumerate(records, 1):
            # 显示格式化后的完整信息
            formatted_name = self.format_name(
                record['原始姓名'],
                record['手机号'],
                record['角色'],
                record['客户经理'],
                include_manager,
                include_phone,
                include_seq,
                i
            )
            
            text_lines.append(f"{i:03d}. {formatted_name}")
        
        summary = f"""
{'='*60}
处理完成！
共生成 {len(records)} 条记录
输出文件：{VCARD_CSV}
配置选项：
  • 包含客户经理：{'是' if include_manager else '否'}
  • 包含手机号：{'是' if include_phone else '否'}
  • 包含序号：{'是' if include_seq else '否'}

详细记录：
{'='*60}
"""
        self.log.setPlainText(summary + '\n'.join(text_lines))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(True)  # 确保关闭窗口时退出
    
    # 设置应用程序样式
    app.setStyle('Fusion')
    
    w = MainWin()
    w.show()
    sys.exit(app.exec_())
