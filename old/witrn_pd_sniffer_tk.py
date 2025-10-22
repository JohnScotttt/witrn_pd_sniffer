#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WITRN HID PD查看器 GUI 应用程序
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import os
import base64
import tempfile
import csv
import threading
import time
from datetime import datetime
from typing import List, Dict, Any, Optional
import json
from witrnhid import WITRN_DEV, metadata, is_pdo, is_rdo, provide_ext
from icon import brain_ico


MT = {
    "GoodCRC": "#e4fdff",
    "GotoMin": "#eea8e8",
    "Accept": "#caffbf",
    "Reject": "#ec7777",
    "Ping": "#adc178",
    "PS_RDY": "#fff0d4",
    "Get_Source_Cap": "#949af1",
    "Get_Sink_Cap": "#a0b6ff",
    "DR_Swap": "#00bfff",
    "PR_Swap": "#4293e4",
    "VCONN_Swap": "#ffa7ff",
    "Wait": "#ff8fab",
    "Soft_Reset": "#da96ac",
    "Data_Reset": "#afeeee",
    "Data_Reset_Complete": "#dba279",
    "Not_Supported": "#a9a9a9",
    "Get_Source_Cap_Extended": "#bdb76b",
    "Get_Status": "#d884d8",
    "FR_Swap": "#556b2f",
    "Get_PPS_Status": "#ff8c00",
    "Get_Country_Codes": "#dbadf1",
    "Get_Sink_Cap_Extended": "#eb7676",
    "Get_Source_Info": "#e9967a",
    "Get_Revision": "#8fbc8f",
    "Source_Capabilities": "#abc4ff",
    "Request": "#ffc6ff",
    "BIST": "#BDDA56",
    "Sink_Capabilities": "#20b2aa",
    "Battery_Status": "#ffb6c0",
    "Alert": "#afeeee",
    "Get_Country_Info": "#ffffe0",
    "Enter_USB": "#92ba92",
    "EPR_Request": "#ffc6ff",
    "EPR_Mode": "#92ba92",
    "Source_Info": "#fff0d4",
    "Revision": "#f0ead2",
    "Vendor_Defined": "#bdb2ff",
    "Source_Capabilities_Extended": "#b8e0d4",
    "Status": "#d8bfd8",
    "Get_Battery_Cap": "#f3907f",
    "Get_Battery_Status": "#40e0d0",
    "Battery_Capabilities": "#dde5b4",
    "Get_Manufacturer_Info": "#f5deb3",
    "Manufacturer_Info": "#c0c0c0",
    "Security_Request": "#9acd32",
    "Security_Response": "#da70d6",
    "Firmware_Update_Request": "#d2b48c",
    "Firmware_Update_Response": "#00ff7f",
    "PPS_Status": "#7CE97C",
    "Country_Info": "#6a5acd",
    "Country_Codes": "#87ceeb",
    "Sink_Capabilities_Extended": "#ee82ee",
    "Extended_Control": "#e99771",
    "EPR_Source_Capabilities": "#8199d1",
    "EPR_Sink_Capabilities": "#f4a460",
    "Vendor_Defined_Extended": "#bdb2ff",
    "Reserved": "#fa8072",
}


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
        # 先隐藏主窗口，等布局和几何设置完成后再显示，避免启动时小窗闪烁
        try:
            self.root.withdraw()
        except Exception:
            pass
        self.root.title("WITRN PD解析v1.4 by JohnScotttt")
        # 使用内置的 base64 图标（brain_ico）设置窗口图标；失败则回退到本地 brain.ico
        try:
            ico_bytes = base64.b64decode(brain_ico)
            tmp_path = os.path.join(tempfile.gettempdir(), "witrn_pd_sniffer_brain.ico")
            with open(tmp_path, "wb") as f:
                f.write(ico_bytes)
            self.root.iconbitmap(tmp_path)
        except Exception:
            pass
        # 锁定窗口大小，禁止用户调整（固定宽高）
        self.root.resizable(False, False)
        try:
            w, h = 1600, 870
            self.root.minsize(w, h)
            self.root.maxsize(w, h)
        except Exception:
            pass
        # 尝试将窗口放在屏幕中央（在设置固定大小后计算）
        try:
            # 使用请求的大小或者当前窗口大小作为目标宽高
            # 注意：在某些环境下 winfo_width/winfo_height 可能在窗口尚未显示前返回 1，
            # 因此优先使用我们设置的固定尺寸 w,h（如果可用）
            target_w = locals().get('w', None) or self.root.winfo_width()
            target_h = locals().get('h', None) or self.root.winfo_height()

            # 如果尺寸不合理（如 1），使用默认值
            if not target_w or target_w <= 1:
                target_w = 1600
            if not target_h or target_h <= 1:
                target_h = 870

            screen_w = self.root.winfo_screenwidth()
            screen_h = self.root.winfo_screenheight()
            x = int((screen_w - target_w) / 2)
            y = int((screen_h - target_h) / 2) - 50
            # 设置几何位置为 WxH+X+Y
            self.root.geometry(f"{int(target_w)}x{int(target_h)}+{x}+{y}")
        except Exception:
            # 若任何一步失败，不阻塞主程序
            pass
        
        # 数据存储
        self.data_list: List[DataItem] = []
        self.current_selection: Optional[DataItem] = None
        
        # 控制状态
        self.is_paused = False
        self.data_collection_active = False
        # 导入模式：当通过CSV导入数据后置True，禁止开始采集；清空后复位
        self.import_mode = False
        # 设备连接状态（open 成功后为 True，断开或未打开为 False）
        self.device_open = False
        
        # 创建界面
        self.create_widgets()
        
        # 启动数据刷新线程
        self.refresh_thread = threading.Thread(target=self.refresh_data_loop, daemon=True)
        self.refresh_thread.start()

        # 设备句柄和数据采集线程控制
        self.k2 = None
        self.data_thread_started = False
        # 所有初始化完成后再显示窗口，减少启动闪烁
        try:
            # 先让 Tk 计算完布局和几何信息，再一次性显示
            self.root.update_idletasks()
            self.root.deiconify()
        except Exception:
            pass
    
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
        column_widths = {'序号': 30, '时间': 80, 'SOP': 90, 'PPR': 120, 'PDR': 80, 'Msg Type': 160}
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

        # 为不同的消息类型注册 tag（用于行背景色）
        # Treeview 支持按 tag 设置行的 background/foreground
        try:
            for msg_type, hex_color in MT.items():
                # 使用 msg_type 名称作为 tag 名
                # 注意：Treeview 的 tag_configure 接受颜色名称或十六进制
                try:
                    self.tree.tag_configure(msg_type, background=hex_color)
                except Exception:
                    # 如果 msg_type 包含特殊字符导致失败，则用安全的标签名
                    safe_tag = f"mt_{abs(hash(msg_type))}"
                    self.tree.tag_configure(safe_tag, background=hex_color)
        except Exception:
            # 忽略标签配置错误，程序仍能正常工作
            pass
        
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
        # 注册颜色标签（Text widget 使用 tag 来控制文本样式）
        try:
            self.data_text.config(selectbackground="#b4d9fb", selectforeground="black")
            # 这里用 tag_configure 注册需要的样式名
            self.data_text.tag_configure('red', foreground='red')
            self.data_text.tag_configure('blue', foreground='blue')
            self.data_text.tag_configure('green', foreground='green')
            self.data_text.tag_configure('bold', font=('Consolas', 10, 'bold'))
        except Exception:
            # 在极端环境下 tag_configure 可能失败，但不影响基本功能
            pass
        
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
        
        # 导出列表按钮（导出为 CSV）
        self.export_button = ttk.Button(
            data_frame,
            text="导出列表",
            command=self.export_list,
            state=tk.DISABLED
        )
        self.export_button.pack(side=tk.LEFT, padx=(0, 10))

        # 导入CSV按钮
        self.import_button = ttk.Button(
            data_frame,
            text="导入CSV",
            command=self.import_csv
        )
        self.import_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # 清空按钮
        self.clear_button = ttk.Button(
            data_frame, 
            text="清空列表", 
            command=self.clear_list
        )
        self.clear_button.pack(side=tk.LEFT)

        # 状态变量（显示在右侧数据显示区域底部的状态栏）
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        # 在 right_frame 底部放置横向状态栏，便于显示较长文本
        status_label = ttk.Label(right_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_label.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 0))

        # 在 header 中放置复选框以与标题并列
        self.filter_goodcrc_var = tk.BooleanVar(value=False)
        self.filter_goodcrc_cb = ttk.Checkbutton(
            header_frame,
            text="屏蔽GoodCRC",
            variable=self.filter_goodcrc_var,
            command=self.update_treeview
        )
        self.filter_goodcrc_cb.pack(side=tk.LEFT, padx=(10, 0))
    
    def add_data_item(self, sop: str, ppr: str, pdr: str, msg_type: str, data: Any = None, force: bool = False, timestamp: Optional[str] = None):
        """添加新的数据项到列表"""
        # 只有在数据收集激活且未暂停时才添加数据，除非强制添加
        if not force and (not self.data_collection_active or self.is_paused):
            return
            
        timestamp = timestamp or datetime.now().strftime("%H:%M:%S.%f")[:-3]
        index = len(self.data_list) + 1
        
        item = DataItem(index, timestamp, sop, ppr, pdr, msg_type, data)
        self.data_list.append(item)
        
        # 更新状态
        self.status_var.set(f"已读取 {len(self.data_list)} 条数据")

    def reconnect_device(self):
        """尝试重新连接设备（由 UI 按钮调用）。成功则启用开始按钮并启动数据线程。"""
        try:
            self.status_var.set("尝试连接设备...")
            # 重新创建实例并显式打开
            self.k2 = WITRN_DEV()
            try:
                self.k2.open()
            except Exception as e:
                raise RuntimeError(f"open() 失败: {e}")

            self.status_var.set("设备已连接")
            self.device_open = True
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
                if self.device_open is False:
                    time.sleep(0.1)
                    continue
                self.k2.read_data()
                timestamp, pkg = self.k2.auto_unpack()
                if pkg.field() == "pd":
                    sop = pkg["SOP*"].value()
                    try:
                        ppr = pkg["Message Header"][3].value()
                    except:
                        ppr = None
                    try:
                        pdr = pkg["Message Header"][5].value()
                    except:
                        pdr = None
                    try:
                        msg_type = pkg["Message Header"]["Message Type"].value()
                    except:
                        msg_type = None
                    data = pkg
                    self.add_data_item(sop, ppr, pdr, msg_type, data, timestamp=timestamp)
            except Exception as e:
                # 仅当是真正的 read error 时才视为设备断开；其他错误可能不严重
                err_text = str(e).lower()
                if 'read error' in err_text or 'readerror' in err_text:
                    # 发生 read error，认为设备已断开或不可用
                    print(f"数据采集异常（断开）: {e}")
                    # 将设备句柄置空，停止当前数据收集状态
                    try:
                        self.data_collection_active = False
                        self.device_open = False
                        self.k2 = WITRN_DEV()
                    except Exception:
                        pass

                    # 在线程安全地更新 UI：启用重连按钮，禁用开始/暂停按钮，更新状态
                    def _on_disconnect():
                        try:
                            self.status_var.set("设备断开")
                            self.device_open = False
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
                else:
                    # 非 read error：记录为警告并继续（不认为设备断开）
                    print(f"数据采集警告（非断开）: {e}")
                    try:
                        # 在状态栏短暂显示警告信息（不打断用户）
                        self.root.after(0, lambda: self.status_var.set(f"数据采集错误（可忽略）: {e}"))
                    except Exception:
                        pass
                    # 短暂停顿后继续循环，避免高速打印
                    time.sleep(0.1)
    
    def update_treeview(self):
        """更新Treeview显示"""
        # 保存当前选中的项目（按 Treeview item id），以及当前滚动位置和可见性/焦点状态。
        current_selection = self.tree.selection()
        selected_index = None
        selected_item_id = None
        selected_was_visible = False
        selected_was_focused = False
        try:
            if current_selection:
                selected_item_id = current_selection[0]
                selected_item_vals = self.tree.item(selected_item_id)
                try:
                    selected_index = int(selected_item_vals['values'][0]) - 1  # 转换为0基索引
                except Exception:
                    selected_index = None

                # 如果选中项在更新前是可见的（bbox 返回非 None），记录下来
                try:
                    bbox = self.tree.bbox(selected_item_id)
                    selected_was_visible = bbox is not None
                except Exception:
                    selected_was_visible = False

                # 记录选中项是否拥有焦点
                try:
                    selected_was_focused = (self.tree.focus() == selected_item_id)
                except Exception:
                    selected_was_focused = False
        except Exception:
            selected_index = None

        # 保存当前滚动位置（yview），以便在不希望强制滚动时恢复用户的视图。
        try:
            prev_yview = self.tree.yview()
        except Exception:
            prev_yview = None

        # 重新构建可见项目，考虑过滤选项
        for child in self.tree.get_children():
            self.tree.delete(child)

        hide_goodcrc = bool(getattr(self, 'filter_goodcrc_var', tk.BooleanVar()).get())

        for item in self.data_list:
            if hide_goodcrc and isinstance(item.msg_type, str) and 'goodcrc' in item.msg_type.lower():
                continue
            # 计算用于 tag 的名称，优先使用 msg_type 的原始名称（如果在 MT 中注册过）
            tag_name = None
            try:
                if isinstance(item.msg_type, str) and item.msg_type in MT:
                    tag_name = item.msg_type
                elif isinstance(item.msg_type, str):
                    # 有时 msg_type 可能带有空格或大小写不同，尝试按不区分大小写匹配
                    lowered = item.msg_type.lower()
                    for k in MT.keys():
                        if k.lower() == lowered:
                            tag_name = k
                            break
                # 如果仍然没有匹配，但 msg_type 是字符串且 MT 中存在相同颜色值的安全标签，我们 won't create new colors here
            except Exception:
                tag_name = None

            # 如果 tag_name 为 None，则不传递 tag（使用默认背景）
            if tag_name:
                try:
                    self.tree.insert('', tk.END, values=(
                        item.index,
                        item.timestamp,
                        item.sop,
                        item.ppr,
                        item.pdr,
                        item.msg_type
                    ), tags=(tag_name,))
                except Exception:
                    # 如果插入时 tag 出错，退回到无 tag 插入
                    self.tree.insert('', tk.END, values=(
                        item.index,
                        item.timestamp,
                        item.sop,
                        item.ppr,
                        item.pdr,
                        item.msg_type
                    ))
            else:
                self.tree.insert('', tk.END, values=(
                    item.index,
                    item.timestamp,
                    item.sop,
                    item.ppr,
                    item.pdr,
                    item.msg_type
                ))

        # 恢复选中状态（如果选中的项仍然可见）。
        # 重要：不要在每次刷新时强制滚动到选中项，以免打断用户的手动滚动。
        if selected_index is not None and 0 <= selected_index < len(self.data_list):
            target_child = None
            for child in self.tree.get_children():
                item_values = self.tree.item(child)['values']
                if item_values and int(item_values[0]) == selected_index + 1:
                    target_child = child
                    break

            if target_child is not None:
                try:
                    # 恢复 selection（高亮），但不要强制滚动到该项。
                    # 将 selection_set 放在前面以保持高亮，但随后我们将恢复 yview，
                    # 以确保用户的可视窗口不会被刷新打断。
                    self.tree.selection_set(target_child)
                    if selected_was_focused:
                        try:
                            self.tree.focus(target_child)
                        except Exception:
                            pass
                except Exception:
                    pass
            # 无论选中项是否在可见区域，都尝试恢复先前的滚动位置，优先保持用户视图不变。
            try:
                if prev_yview and len(prev_yview) == 2:
                    self.tree.yview_moveto(prev_yview[0])
            except Exception:
                pass
            else:
                # 如果找不到对应项，恢复原来的 yview（如果可用），以保持用户视图
                try:
                    if prev_yview and len(prev_yview) == 2:
                        self.tree.yview_moveto(prev_yview[0])
                except Exception:
                    pass
        # 根据是否有数据启用/禁用导出按钮
        try:
            if self.data_list:
                self.export_button.config(state=tk.NORMAL)
            else:
                self.export_button.config(state=tk.DISABLED)
        except Exception:
            pass
    
    def on_item_select(self, event):
        """处理列表项选择事件"""
        selection = self.tree.selection()
        if selection:
            item = self.tree.item(selection[0])
            index = int(item['values'][0]) - 1  # 转换为0基索引
            
            if 0 <= index < len(self.data_list):
                self.current_selection = self.data_list[index]
                self.display_data(self.current_selection)
    
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
            info = (f"基本信息:\n序号: {item.index}\n时间: {item.timestamp}\nSOP: {item.sop}\n"
                    f"PPR: {item.ppr}\nPDR: {item.pdr}\n消息类型: {item.msg_type}\n\n详细数据:\n")

            self.data_text.insert(tk.END, info)
            self.data_text.insert(tk.END, f"Raw: 0x{int(item.data.raw(), 2):0{int(len(item.data.raw())/4)+(1 if len(item.data.raw())%4!=0 else 0)}X}\n", 'green')
            for value1 in item.data.value():
                if not isinstance(value1.value(), list):
                    self.data_text.insert(tk.END, f"{value1.field()+':':<30}")
                    self.data_text.insert(tk.END, f"(Raw: 0x{int(value1.raw(), 2):0{int(len(value1.raw())/4)+(1 if len(value1.raw())%4!=0 else 0)}X})\n", 'green')
                    self.data_text.insert(tk.END, f"    {value1.value()}\n")
                else:
                    self.data_text.insert(tk.END, f"{value1.field()+':':<30}")
                    self.data_text.insert(tk.END, f"(Raw: 0x{int(value1.raw(), 2):0{int(len(value1.raw())/4)+(1 if len(value1.raw())%4!=0 else 0)}X})\n", 'green')
                    for value2 in value1.value():
                        if not isinstance(value2.value(), list):
                            self.data_text.insert(tk.END, f"    {value2.field()}: {value2.value()}\n")
                        else:
                            self.data_text.insert(tk.END, f"    {value2.field()+':':<30}")
                            self.data_text.insert(tk.END, f"(Raw: 0x{int(value2.raw(), 2):08X})\n", 'green')
                            for value3 in value2.value():
                                if not isinstance(value3.value(), list):
                                    self.data_text.insert(tk.END, f"        {value3.field()}: {value3.value()}\n")
                                else:
                                    self.data_text.insert(tk.END, f"        {value3.field()}:\n")
                                    for value4 in value3.value():
                                        self.data_text.insert(tk.END, f"            {value4.field()}: {value4.value()}\n")

        finally:
            # 设回只读，防止用户编辑
            self.data_text.config(state=tk.DISABLED)


    def format_data(self, data: Any) -> str:
        """将数据格式化为用于 CSV 的字符串——按需直接使用 repr。
        注意：用户自定义了 metadata.__repr__ 与 __str__，此处必须调用 repr。
        """
        try:
            return repr(data)
        except Exception as e:
            # 避免退回到 str，保持显式标注错误
            return f"<repr-error: {e}>"

    def export_list(self):
        """将当前数据列表导出为 CSV 文件（包含详细数据字段）"""
        if not self.data_list:
            messagebox.showwarning("警告", "没有数据可导出")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[('CSV 文件', '*.csv')],
            title='保存为 CSV'
        )
        if not file_path:
            return

        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['序号', '时间', 'SOP', 'PPR', 'PDR', 'Msg Type', '详细数据', 'Raw'])
                for item in self.data_list:
                    # 使用 format_data 输出的人类可读文本作为详细数据字段
                    data_text = self.format_data(item.data)
                    writer.writerow([item.index,
                                     item.timestamp,
                                     item.sop,
                                     item.ppr,
                                     item.pdr,
                                     item.msg_type,
                                     data_text,
                                     f"{int(item.data.raw(), 2):0{int(len(item.data.raw())/4)+(1 if len(item.data.raw())%4!=0 else 0)}X}"])

            self.status_var.set(f"已导出 {len(self.data_list)} 条数据 到 {file_path}")
        except Exception as e:
            messagebox.showerror("导出失败", f"导出 CSV 失败:\n{e}")

    def import_csv(self):
        """从CSV导入数据，仅解析两列：时间、Raw（全大写HEX）。
        Raw 将被转换为长度为64字节的uint8列表（不足末尾补0，超出则截断），
        然后使用 WITRN_DEV.auto_unpack(data) 解析并加入列表。
        """
        # 若正在进行数据采集（未暂停），阻止导入
        if self.data_collection_active and not self.is_paused:
            messagebox.showwarning("操作受限", "正在收集数据，无法导入CSV。请先暂停并清空列表后再试。")
            return
        file_path = filedialog.askopenfilename(
            filetypes=[('CSV 文件', '*.csv')],
            title='选择CSV文件'
        )
        if not file_path:
            return

        # 重新导入应覆盖当前列表：先静默清空
        try:
            self.clear_list(ask_user=False)
        except Exception:
            pass

        # 确保有一个设备实例用于解析（仅需构造实例用于 auto_unpack，无需调用 open()）
        try:
            if self.k2 is None:
                self.k2 = WITRN_DEV()
        except Exception as e:
            messagebox.showerror("导入失败", f"无法创建 WITRN_DEV 实例以解析CSV：\n{e}\n\n请连接设备或确保依赖可用后重试。")
            return

        success, failed = 0, 0
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                # 兼容可能的列名大小写或空格
                col_map = {k.strip(): k for k in reader.fieldnames or []}
                def get_col(name_alternatives):
                    for n in name_alternatives:
                        if n in col_map:
                            return col_map[n]
                    return None

                time_col = get_col(["时间", "time", "Time"])
                raw_col = get_col(["Raw", "RAW", "raw"])
                if raw_col is None:
                    raise ValueError("CSV中缺少列：Raw")
                
                last_pdo = None
                last_rdo = None
                last_ext = None

                for row in reader:
                    try:
                        t_str = row[time_col].strip() if time_col and row.get(time_col) is not None else None
                        raw_hex = (row.get(raw_col) or "").strip()
                        if not raw_hex:
                            failed += 1
                            continue

                        # 允许前缀0x，可选空格，统一为连续十六进制
                        raw_hex = raw_hex.replace(" ", "").replace("0X", "0x")
                        if raw_hex.startswith("0x"):
                            raw_hex = raw_hex[2:]
                        # 若为奇数长度，前置0补齐
                        if len(raw_hex) % 2 == 1:
                            raw_hex = '0' + raw_hex

                        # 将HEX字符串转为字节数组（大写/小写均可）
                        try:
                            data_bytes = bytearray.fromhex(raw_hex)
                        except Exception:
                            failed += 1
                            continue

                        # 规范到64字节：超出则截断，不足则末尾补0
                        if len(data_bytes) > 64:
                            data_bytes = data_bytes[:64]
                        elif len(data_bytes) < 64:
                            data_bytes.extend([0] * (64 - len(data_bytes)))

                        # 解析
                        try:
                            if self.k2 is None:
                                # 若仍无k2，跳过解析
                                failed += 1
                                continue
                            _, pkg = self.k2.auto_unpack(data_bytes, last_pdo, last_ext, last_rdo)
                            if is_pdo(pkg):
                                last_pdo = pkg
                            if is_rdo(pkg):
                                last_rdo = pkg
                            if provide_ext(pkg):
                                last_ext = pkg
                            
                        except Exception:
                            failed += 1
                            continue

                        try:
                            if pkg.field() == "pd":
                                sop = pkg["SOP*"].value()
                                try:
                                    ppr = pkg["Message Header"][3].value()
                                except Exception:
                                    ppr = None
                                try:
                                    pdr = pkg["Message Header"][5].value()
                                except Exception:
                                    pdr = None
                                try:
                                    msg_type = pkg["Message Header"]["Message Type"].value()
                                except Exception:
                                    msg_type = None
                                self.add_data_item(sop, ppr, pdr, msg_type, pkg, force=True, timestamp=t_str)
                                success += 1
                            else:
                                failed += 1
                        except Exception:
                            failed += 1
                            continue
                    except Exception:
                        failed += 1
                        continue

            # 导入完成，刷新视图
            self.update_treeview()
            # 启用导出按钮（若有数据）
            try:
                if self.data_list:
                    self.export_button.config(state=tk.NORMAL)
            except Exception:
                pass
            # 置为导入模式：设备已连接则开启“开始收集”，否则保持禁用；“暂停/恢复”重置为“暂停”并禁用
            self.import_mode = True
            try:
                if getattr(self, 'device_open', False):
                    self.start_button.config(state=tk.NORMAL)
                else:
                    self.start_button.config(state=tk.DISABLED)
                self.pause_button.config(text="暂停", state=tk.DISABLED)
            except Exception:
                pass
            if getattr(self, 'device_open', False):
                self.status_var.set(f"导入完成：成功 {success} 条，失败 {failed} 条。可开始收集（会先清空）。")
            else:
                self.status_var.set(f"导入完成：成功 {success} 条，失败 {failed} 条。设备未连接，无法开始收集；请先连接设备。")
        except Exception as e:
            messagebox.showerror("导入失败", f"无法导入CSV:\n{e}")
    
    def clear_list(self, ask_user: bool = True):
        """清空数据列表，并重置导入模式。"""
        if (not ask_user) or messagebox.askyesno("确认", "确定要清空所有数据吗？"):
            self.data_list.clear()
            self.current_selection = None
            self.update_treeview()
            # 临时启用以清空显示区域，然后恢复为只读
            try:
                self.data_text.config(state=tk.NORMAL)
                self.data_text.delete(1.0, tk.END)
            finally:
                self.data_text.config(state=tk.DISABLED)
            try:
                self.export_button.config(state=tk.DISABLED)
            except Exception:
                pass
            # 退出导入模式
            self.import_mode = False
            # 如果当前处于正常采集状态（且未暂停），不更改开始/暂停按钮状态
            if not (self.data_collection_active and not self.is_paused):
                # 根据设备连接状态控制“开始收集”可用性，并重置暂停按钮
                try:
                    if getattr(self, 'device_open', False):
                        self.start_button.config(state=tk.NORMAL)
                    else:
                        self.start_button.config(state=tk.DISABLED)
                    # 清空后恢复暂停按钮默认状态（未收集时禁用）
                    self.pause_button.config(state=tk.DISABLED, text="暂停")
                except Exception:
                    pass
            self.status_var.set("列表已清空")
    
    def start_collection(self):
        """开始数据收集"""
        # 未连接设备（未 open）时不允许开始
        if not getattr(self, 'device_open', False):
            messagebox.showwarning("设备未连接", "未检测到已连接的设备。请先连接设备后再开始收集。")
            return
        # 导入模式时需清空后才能开始
        if self.import_mode:
            if not self.data_list:
                # 仅有标志位，防御性复位
                self.import_mode = False
            else:
                if messagebox.askyesno("清空后开始", "当前列表来自CSV导入，须清空后才能开始收集。是否清空并开始？"):
                    self.clear_list(ask_user=False)
                else:
                    return

        # 如果列表非空且不是导入模式，也询问是否清空以保持干净数据
        if self.data_list:
            if not messagebox.askyesno("清空后开始", "开始收集前建议清空现有列表。是否清空并开始？"):
                return
            self.clear_list(ask_user=False)

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
    import threading
    
    # 创建并运行GUI
    app = WITRNGUI()
    
    # 创建实例并尝试打开设备；失败时保留GUI并允许手动重连
    k2 = None
    try:
        k2 = WITRN_DEV()
        try:
            k2.open()
            app.k2 = k2
            app.device_open = True
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
            # open 失败
            messagebox.showerror("错误", f"无法连接到K2 (open失败):\n{e}")
            try:
                app.status_var.set("无法连接到K2")
                app.start_button.config(state=tk.DISABLED)
                app.device_open = False
            except Exception:
                pass
            # 保留未打开的实例用于离线解析（如导入CSV）
            app.k2 = k2
    except Exception as e:
        # 构造失败
        messagebox.showerror("错误", f"无法创建K2实例:\n{e}")
        try:
            app.status_var.set("无法连接到K2")
            app.start_button.config(state=tk.DISABLED)
            app.device_open = False
        except Exception:
            pass
        k2 = None
    
    # 启动数据收集线程（如果设备已连接，方法内部会避免重复启动）
    app.start_data_thread_if_needed()

    # 运行GUI（如果未连接设备，GUI 仍然可运行，用户可看到错误状态并可手动重连）
    app.run()

# Nuitka 打包示例（确保 brain.ico 被打包并供 Tk 在运行时读取）：
# python -m nuitka witrn_pd_sniffer_tk.py --standalone --onefile --windows-console-mode=disable --enable-plugin=tk-inter --windows-icon-from-ico=brain.ico