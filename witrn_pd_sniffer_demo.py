#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WITRN HID PD查看器 GUI 应用程序
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import time
from datetime import datetime
from typing import List, Dict, Any, Optional
import json
from witrnhid import WITRN_DEV
import sys


class DataItem:
    """数据项类，表示列表中的一行数据"""
    def __init__(self, index: int, timestamp: str, sop: str, ppr: str, pdr: str, msg_type: str, data: Any = None):
        self.index = index
        self.timestamp = timestamp
        self.sop = sop
        self.ppr = ppr
        self.pdr = pdr
        self.msg_type = msg_type
        self.data = data or {}


class WITRNGUI:
    """WITRN HID 数据查看器主类"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("WITRN HID 数据查看器")
        self.root.geometry("1300x800")
        # 锁定窗口大小，禁止用户调整（固定宽高）
        self.root.resizable(False, False)
        try:
            w, h = 1300, 800
            self.root.minsize(w, h)
            self.root.maxsize(w, h)
        except Exception:
            pass
        
        # 数据存储
        self.data_list: List[DataItem] = []
        self.current_selection: Optional[DataItem] = None
        
        # 控制状态
        self.is_paused = False
        self.data_collection_active = False
        
        # 创建界面
        self.create_widgets()
        
        # 启动数据刷新线程
        self.refresh_thread = threading.Thread(target=self.refresh_data_loop, daemon=True)
        self.refresh_thread.start()

        # 设备句柄和数据采集线程控制
        self.k2 = None
        self.data_thread_started = False
    
    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 左侧列表框架（使用 header_frame 显示标题与复选框并列）
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        # header：包含标题和过滤复选框并列
        header_frame = ttk.Frame(left_frame)
        header_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 5))
        title_label = ttk.Label(header_frame, text="数据列表")
        title_label.pack(side=tk.LEFT)
        
        # 创建Treeview（表格）
        columns = ('序号', '时间', 'SOP', 'PPR', 'PDR', 'Msg Type')
        self.tree = ttk.Treeview(left_frame, columns=columns, show='headings', height=20)
        
        # 设置列标题和宽度
        column_widths = {'序号': 30, '时间': 80, 'SOP': 90, 'PPR': 120, 'PDR': 100, 'Msg Type': 160}
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths[col], anchor=tk.CENTER)
        
        # 配置选择样式，确保选中行高亮显示
        style = ttk.Style()
        style.configure("Treeview", rowheight=25)
        
        # 设置选中行的样式
        try:
            # 尝试使用蓝色高亮
            style.map("Treeview", 
                     background=[('selected', '#0078d4')],  # 选中行背景色
                     foreground=[('selected', 'white')])    # 选中行文字颜色
            print("使用自定义蓝色高亮样式")
        except Exception as e:
            print(f"自定义样式设置失败: {e}")
            try:
                # 尝试使用系统默认高亮
                style.map("Treeview", 
                         background=[('selected', '')],  # 使用系统默认选中背景色
                         foreground=[('selected', '')])  # 使用系统默认选中文字颜色
                print("使用系统默认高亮样式")
            except Exception as e2:
                print(f"系统默认样式也失败: {e2}")
                # 如果都失败，至少设置基本样式
                style.configure("Treeview", rowheight=25)
        
        # 添加滚动条
        tree_scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scrollbar.set)
        
        # 布局
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 绑定选择事件
        self.tree.bind('<<TreeviewSelect>>', self.on_item_select)
        self.tree.bind('<Button-1>', self.on_item_click)  # 绑定鼠标点击事件
        
        # 右侧数据显示框架
        right_frame = ttk.LabelFrame(main_frame, text="数据显示", padding=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # 数据文本显示区域
        self.data_text = scrolledtext.ScrolledText(
            right_frame, 
            wrap=tk.WORD, 
            width=50, 
            height=25,
            font=('Consolas', 10),
            state=tk.DISABLED  # 初始为只读
        )
        self.data_text.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 按钮框架
        button_frame = ttk.Frame(right_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 控制按钮框架
        control_frame = ttk.Frame(button_frame)
        control_frame.pack(side=tk.LEFT, padx=(0, 20))
        
        # 开始按钮
        self.start_button = ttk.Button(
            control_frame, 
            text="开始收集", 
            command=self.start_collection
        )
        self.start_button.pack(side=tk.LEFT, padx=(0, 5))
        
        # 暂停按钮
        self.pause_button = ttk.Button(
            control_frame, 
            text="暂停", 
            command=self.pause_collection,
            state=tk.DISABLED
        )
        self.pause_button.pack(side=tk.LEFT, padx=(0, 5))

        # 重连按钮（若设备未连接，允许用户重试）
        self.reconnect_button = ttk.Button(
            control_frame,
            text="重连设备",
            command=self.reconnect_device
        )
        self.reconnect_button.pack(side=tk.LEFT, padx=(0, 5))
        
        # 数据操作按钮框架
        data_frame = ttk.Frame(button_frame)
        # 让该框架水平扩展，这样状态显示可以右对齐到该框架的末端
        data_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 复制按钮
        self.copy_button = ttk.Button(
            data_frame, 
            text="复制数据", 
            command=self.copy_data,
            state=tk.DISABLED
        )
        self.copy_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # 清空按钮
        self.clear_button = ttk.Button(
            data_frame, 
            text="清空列表", 
            command=self.clear_list
        )
        self.clear_button.pack(side=tk.LEFT)

        # 状态显示（与清空按钮同级，右对齐）
        # 放在 data_frame 内并靠右显示，保留 relief 以突出显示
        
        # 状态显示（与清空按钮同级，右对齐）
        # 放在 data_frame 内并靠右显示，保留 relief 以突出显示
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        # 使用容器使状态显示固定像素宽度（200px）
        # 给容器一个固定高度以匹配按钮行的高度
        status_container = ttk.Frame(data_frame, width=200, height=20)
        status_container.pack(side=tk.RIGHT, padx=(10, 0))
        # 不允许容器根据子控件自动调整大小，保持固定宽度
        status_container.pack_propagate(False)
        status_label = ttk.Label(status_container, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.E)
        status_label.pack(fill=tk.BOTH, expand=True)

        # 在 header 中放置复选框以与标题并列
        self.filter_goodcrc_var = tk.BooleanVar(value=False)
        self.filter_goodcrc_cb = ttk.Checkbutton(
            header_frame,
            text="屏蔽goodCRC",
            variable=self.filter_goodcrc_var,
            command=self.update_treeview
        )
        self.filter_goodcrc_cb.pack(side=tk.LEFT, padx=(10, 0))
    
    def add_data_item(self, sop: str, ppr: str, pdr: str, msg_type: str, data: Any = None, force: bool = False):
        """添加新的数据项到列表"""
        # 只有在数据收集激活且未暂停时才添加数据，除非强制添加
        if not force and (not self.data_collection_active or self.is_paused):
            return
            
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        index = len(self.data_list) + 1
        
        item = DataItem(index, timestamp, sop, ppr, pdr, msg_type, data)
        self.data_list.append(item)
        
        # 更新状态
        self.status_var.set(f"已读取 {len(self.data_list)} 条数据")

    def reconnect_device(self):
        """尝试重新连接设备（由 UI 按钮调用）。成功则启用开始按钮并启动数据线程。"""
        try:
            self.status_var.set("尝试连接设备...")
            self.k2 = witrnhid.WITRN_DEV()
            self.status_var.set("设备已连接")
            self.start_button.config(state=tk.NORMAL)
            # 连接成功后禁用重连按钮
            try:
                self.reconnect_button.config(state=tk.DISABLED)
            except Exception:
                pass
            # 确保暂停按钮恢复为默认文本并禁用（等待用户点击开始）
            try:
                self.pause_button.config(text="暂停", state=tk.DISABLED)
            except Exception:
                pass
            # 启动数据线程（如果尚未启动）
            self.start_data_thread_if_needed()
        except Exception as e:
            messagebox.showerror("重连失败", f"无法连接到K2:\n{e}")
            self.status_var.set("无法连接到K2")

    def start_data_thread_if_needed(self):
        """在设备可用时安全地启动数据采集线程一次。"""
        if getattr(self, 'data_thread_started', False):
            return
        if getattr(self, 'k2', None) is None:
            return

        # 启动收集线程
        t = threading.Thread(target=self._collect_data_loop, daemon=True)
        t.start()
        self.data_thread_started = True

    def _collect_data_loop(self):
        """内部数据收集循环；与旧的 collect_data 等价但作为实例方法使用 self.k2。"""
        while True:
            try:
                if self.k2 is None:
                    time.sleep(0.5)
                    continue
                self.k2.read_data()
                _, pkg = self.k2.auto_unpack()
                if pkg.field() == "pd":
                    sop = pkg["SOP*"].value()
                    ppr = pkg["Message Header"][3].value()
                    pdr = pkg["Message Header"][5].value()
                    msg_type = pkg["Message Header"]["Message Type"].value()
                    data = pkg
                    self.add_data_item(sop, ppr, pdr, msg_type, data)
            except Exception as e:
                # 发生异常，认为设备可能已断开或不可用
                print(f"数据采集异常，设备可能断开: {e}")
                # 将设备句柄置空，停止当前数据收集状态
                try:
                    self.k2 = None
                    self.data_collection_active = False
                except Exception:
                    pass

                # 在线程安全地更新 UI：启用重连按钮，禁用开始/暂停按钮，更新状态
                def _on_disconnect():
                    try:
                        self.status_var.set("设备断开")
                        self.start_button.config(state=tk.DISABLED)
                        self.pause_button.config(state=tk.DISABLED)
                        self.reconnect_button.config(state=tk.NORMAL)
                        # 弹窗提示用户设备断开
                        try:
                            messagebox.showwarning("设备断开", "检测到设备已断开，请重连或检查连接。")
                        except Exception:
                            pass
                    except Exception:
                        pass

                try:
                    self.root.after(0, _on_disconnect)
                except Exception:
                    pass

                # 休眠后继续等待重连
                time.sleep(0.5)
    
    def update_treeview(self):
        """更新Treeview显示"""
        # 保存当前选中的项目（按原始索引）
        current_selection = self.tree.selection()
        selected_index = None
        if current_selection:
            selected_item = self.tree.item(current_selection[0])
            try:
                selected_index = int(selected_item['values'][0]) - 1  # 转换为0基索引
            except Exception:
                selected_index = None

        # 重新构建可见项目，考虑过滤选项
        for child in self.tree.get_children():
            self.tree.delete(child)

        hide_goodcrc = bool(getattr(self, 'filter_goodcrc_var', tk.BooleanVar()).get())

        for item in self.data_list:
            if hide_goodcrc and isinstance(item.msg_type, str) and 'goodcrc' in item.msg_type.lower():
                continue
            self.tree.insert('', tk.END, values=(
                item.index,
                item.timestamp,
                item.sop,
                item.ppr,
                item.pdr,
                item.msg_type
            ))

        # 恢复选中状态（如果选中的项仍然可见）
        if selected_index is not None and 0 <= selected_index < len(self.data_list):
            for child in self.tree.get_children():
                item_values = self.tree.item(child)['values']
                if item_values and int(item_values[0]) == selected_index + 1:
                    self.tree.selection_set(child)
                    self.tree.focus(child)
                    self.tree.see(child)
                    break
    
    def on_item_select(self, event):
        """处理列表项选择事件"""
        selection = self.tree.selection()
        if selection:
            item = self.tree.item(selection[0])
            index = int(item['values'][0]) - 1  # 转换为0基索引
            
            if 0 <= index < len(self.data_list):
                self.current_selection = self.data_list[index]
                self.display_data(self.current_selection)
                self.copy_button.config(state=tk.NORMAL)
    
    def on_item_click(self, event):
        """处理鼠标点击事件，确保选中行高亮显示"""
        # 获取点击位置的项目
        item = self.tree.identify_row(event.y)
        if item:
            # 清除所有选择
            for selected in self.tree.selection():
                self.tree.selection_remove(selected)
            
            # 确保该项目被选中
            self.tree.selection_set(item)
            self.tree.focus(item)
            # 确保选中状态可见
            self.tree.see(item)
            
            # 强制刷新显示
            self.root.update_idletasks()
            
            # 触发选择事件
            self.on_item_select(None)
            
            # 调试信息
            print(f"选中项目: {item}, 值: {self.tree.item(item)['values']}")
            print(f"当前选择: {self.tree.selection()}")
            
            # 再次确保选中状态
            self.root.after(10, lambda: self.ensure_selection(item))
    
    def ensure_selection(self, item):
        """确保项目被选中并高亮显示"""
        if item and item in self.tree.get_children():
            self.tree.selection_set(item)
            self.tree.focus(item)
            self.tree.see(item)
            print(f"确保选中: {item}")
    
    def display_data(self, item: DataItem):
        """在右侧显示选中的数据"""
        # 临时启用以写入，然后恢复为只读
        try:
            self.data_text.config(state=tk.NORMAL)
            self.data_text.delete(1.0, tk.END)

            # 基本信息
            info = f"""基本信息:
序号: {item.index}
时间: {item.timestamp}
SOP: {item.sop}
PPR: {item.ppr}
PDR: {item.pdr}
消息类型: {item.msg_type}

详细数据:
"""

            data_str = str(item.data)

            self.data_text.insert(tk.END, info + data_str)
        finally:
            # 设回只读，防止用户编辑
            self.data_text.config(state=tk.DISABLED)
    
    def copy_data(self):
        """复制当前显示的数据到剪贴板"""
        if self.current_selection:
            data = self.data_text.get(1.0, tk.END)
            self.root.clipboard_clear()
            self.root.clipboard_append(data)
            self.status_var.set("数据已复制到剪贴板")
        else:
            messagebox.showwarning("警告", "请先选择要复制的数据项")
    
    def clear_list(self):
        """清空数据列表"""
        if messagebox.askyesno("确认", "确定要清空所有数据吗？"):
            self.data_list.clear()
            self.current_selection = None
            self.update_treeview()
            # 临时启用以清空显示区域，然后恢复为只读
            try:
                self.data_text.config(state=tk.NORMAL)
                self.data_text.delete(1.0, tk.END)
            finally:
                self.data_text.config(state=tk.DISABLED)
            self.copy_button.config(state=tk.DISABLED)
            self.status_var.set("列表已清空")
    
    def start_collection(self):
        """开始数据收集"""
        self.data_collection_active = True
        self.is_paused = False
        self.start_button.config(state=tk.DISABLED)
        self.pause_button.config(state=tk.NORMAL)
        self.status_var.set("数据收集中...")
        # 初始连接成功则禁用重连按钮
        try:
            self.reconnect_button.config(state=tk.DISABLED)
        except Exception:
            pass
    
    def pause_collection(self):
        """暂停/恢复数据收集"""
        if self.is_paused:
            # 恢复收集
            self.is_paused = False
            self.pause_button.config(text="暂停")
            self.status_var.set("数据收集中...")
        else:
            # 暂停收集
            self.is_paused = True
            self.pause_button.config(text="恢复")
            self.status_var.set("数据收集已暂停")
    
    def refresh_data_loop(self):
        """数据刷新循环（在后台线程中运行）"""
        last_data_count = 0
        while True:
            try:
                # 只有在数据发生变化时才更新显示
                current_data_count = len(self.data_list)
                if current_data_count != last_data_count:
                    self.root.after(0, self.update_treeview)
                    last_data_count = current_data_count
                time.sleep(0.1)  # 100ms检查一次
            except Exception as e:
                print(f"刷新数据时出错: {e}")
                time.sleep(1)
    
    def get_data_list(self) -> List[DataItem]:
        """获取当前数据列表（供外部程序调用）"""
        return self.data_list.copy()
    
    def set_data_list(self, data_list: List[DataItem]):
        """设置数据列表（供外部程序调用）"""
        self.data_list = data_list.copy()
        self.update_treeview()
    
    def run(self):
        """运行GUI应用程序"""
        self.root.mainloop()


if __name__ == "__main__":
    import witrnhid
    import threading
    
    # 创建并运行GUI
    app = WITRNGUI()
    
    # 尝试创建WITRN设备，如果无法打开则弹窗并退出
    try:
        k2 = WITRN_DEV()
        app.k2 = k2
        # 启用开始按钮（如果之前被禁用）
        try:
            app.start_button.config(state=tk.NORMAL)
            app.status_var.set("设备已连接")
            try:
                app.reconnect_button.config(state=tk.DISABLED)
            except Exception:
                pass
            try:
                app.pause_button.config(text="暂停", state=tk.DISABLED)
            except Exception:
                pass
        except Exception:
            pass
    except Exception as e:
        # 使用 messagebox 提示错误（app.root 已存在）但保留 GUI，允许手动重试
        messagebox.showerror("错误", f"无法连接到K2:\n{e}")
        # 更新状态并禁用开始按钮
        try:
            app.status_var.set("无法连接到K2")
            app.start_button.config(state=tk.DISABLED)
        except Exception:
            pass
        k2 = None
    
    def collect_data():
        """数据收集函数"""
        while True:
            try:
                k2.read_data()
                _, pkg = k2.auto_unpack()
                if pkg.field() == "pd":
                    sop = pkg["SOP*"].value()
                    ppr = pkg["Message Header"][3].value()
                    pdr = pkg["Message Header"][5].value()
                    msg_type = pkg["Message Header"]["Message Type"].value()
                    data = pkg
                    app.add_data_item(sop, ppr, pdr, msg_type, data)
            except Exception as e:
                pass
            
    
    # 启动数据收集线程（如果设备已连接，方法内部会避免重复启动）
    app.start_data_thread_if_needed()

    # 运行GUI（如果未连接设备，GUI 仍然可运行，用户可看到错误状态并可手动重连）
    app.run()

# python -m nuitka witrn_pd_gui_demo.py --standalone --onefile --windows-disable-console --enable-plugin=tk-inter