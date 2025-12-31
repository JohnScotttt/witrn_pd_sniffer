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
import ctypes
import csv
import threading
import time
from datetime import datetime
from typing import List, Dict, Any, Optional
from witrnhid import WITRN_DEV, metadata, is_pdo, is_rdo, provide_ext
from b64 import brain_ico, jb_r_tff, jb_b_tff
from vendor_ids_dict import VENDOR_IDS
from collections import deque
import multiprocessing
from multiprocessing import Process, Queue, Event, Value
import queue

# 强制导入 matplotlib（环境已保证存在）
import matplotlib
matplotlib.use('TkAgg')  # 必须在导入 pyplot 之前设置
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1) # Windows 8.1 及以上
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware() # 旧版 Windows 兼容
    except Exception:
        pass

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


# ==================== 数据采集进程函数 ====================
def data_collection_worker(data_queue, iv_queue, stop_event, pause_flag):
    """
    独立的数据采集进程
    
    Args:
        data_queue: PD数据队列（发送给主进程）
        iv_queue: 电参数据队列（发送给主进程）- 包含plot和iv_info数据
        stop_event: 停止信号
        pause_flag: 暂停标志（0=收集中, 1=暂停）
    """
    k2 = WITRN_DEV()
    last_pdo = None
    last_rdo = None
    last_general_timestamp = 0.0
    
    try:
        k2.open()
        
        while not stop_event.is_set():
            try:
                # 读取数据（不管是否暂停都读取，避免缓冲区堆积）
                k2.read_data()
                timestamp_str, pkg = k2.auto_unpack()
                current_time = time.time()
                
                if pkg.field() == "general":
                    # general包始终处理（用于iv_info和plot），不受暂停影响
                    # 提取电参数据
                    try:
                        current = float(pkg["Current"].value()[:-1])
                    except:
                        current = 0.0
                    try:
                        voltage = float(pkg["VBus"].value()[:-1])
                    except:
                        voltage = 0.0
                    try:
                        power = abs(current * voltage)
                    except:
                        power = 0.0
                    try:
                        cc1 = float(pkg["CC1"].value()[:-1])
                    except:
                        cc1 = 0.0
                    try:
                        cc2 = float(pkg["CC2"].value()[:-1])
                    except:
                        cc2 = 0.0
                    try:
                        dp = float(pkg["D+"].value()[:-1])
                    except:
                        dp = 0.0
                    try:
                        dn = float(pkg["D-"].value()[:-1])
                    except:
                        dn = 0.0
                    
                    now = time.time()
                    
                    # 发送电参数据（非阻塞）
                    # 包含两个标志：update_plot（总是True）和 update_iv_info（根据频率限制）
                    update_iv_info = (now - last_general_timestamp >= 0.1)  # 10Hz限制
                    
                    try:
                        iv_queue.put_nowait({
                            'timestamp': current_time,
                            'voltage': voltage,
                            'current': current,
                            'power': power,
                            'cc1': cc1,
                            'cc2': cc2,
                            'dp': dp,
                            'dn': dn,
                            'update_plot': True,          # plot数据点无限制
                            'update_iv_info': update_iv_info  # iv_info限制在10Hz
                        })
                    except queue.Full:
                        pass  # 队列满时丢弃旧数据
                    
                    if update_iv_info:
                        last_general_timestamp = now
                    
                elif pkg.field() == "pd":
                    # PD包只在未暂停时处理
                    if pause_flag.value == 1:
                        continue  # 暂停时跳过PD数据
                    
                    # 提取PD数据
                    sop = pkg["SOP*"].value()
                    try:
                        rev = pkg["Message Header"][4].value()[4:]
                        if rev == 'rved':
                            rev = None
                    except:
                        rev = None
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
                    
                    is_pdo_flag = is_pdo(pkg)
                    is_rdo_flag = is_rdo(pkg)
                    
                    if is_pdo_flag:
                        last_pdo = pkg
                    if is_rdo_flag:
                        last_rdo = pkg
                    
                    # 发送PD数据（非阻塞）
                    try:
                        data_queue.put_nowait({
                            'timestamp': timestamp_str,
                            'time_sec': current_time,
                            'sop': sop,
                            'rev': rev,
                            'ppr': ppr,
                            'pdr': pdr,
                            'msg_type': msg_type,
                            'data': pkg,
                            'is_pdo': is_pdo_flag,
                            'is_rdo': is_rdo_flag,
                            'last_pdo': last_pdo,
                            'last_rdo': last_rdo
                        })
                    except queue.Full:
                        pass  # 队列满时丢弃旧数据
                        
            except Exception as e:
                err_text = str(e).lower()
                if 'read error' in err_text:
                    # 设备断开，发送错误信号
                    try:
                        data_queue.put_nowait({'error': 'device_disconnected'})
                    except:
                        pass
                    break
                else:
                    # 其他错误继续
                    time.sleep(0.01)
                    
    except Exception as e:
        # 连接失败
        try:
            data_queue.put_nowait({'error': f'connection_failed: {e}'})
        except:
            pass
    finally:
        try:
            k2.close()
        except:
            pass


def renderer(msg: metadata, level: int, lst: list):
    indent = '    ' * level
    if not isinstance(msg.value(), list):
        if msg.bit_loc()[0] == msg.bit_loc()[1]:
            lst.append((f"{indent}{'[b'+str(msg.bit_loc()[0])+'] ':<12}", 'red'))
        else:
            lst.append((f"{indent}{'[b'+str(msg.bit_loc()[0])+'-b'+str(msg.bit_loc()[1])+'] ':<12}", 'red'))
        lst.append((f"{msg.field()+': '}", ('black', 'bold')))
        lst.append((f"{str(msg.value())} ", 'blue'))
        if msg.field() in ["USB Vendor ID", "VID"]:
            lst.append((f"[{VENDOR_IDS.get(str(msg.value()), 'Unknown Vendor')}] ", 'blue'))
        if level < 1:
            lst.append((f"(0x{int(msg.raw(), 2):0{int(len(msg.raw())/4)+(1 if len(msg.raw())%4 else 0)}X})\n", 'green'))
        else:
            lst.append((f"({msg.raw()}b)\n", 'green'))
    else:
        if msg.bit_loc()[0] == msg.bit_loc()[1]:
            lst.append((f"{indent}{'[b'+str(msg.bit_loc()[0])+'] ':<12}", 'red'))
        else:
            lst.append((f"{indent}{'[b'+str(msg.bit_loc()[0])+'-b'+str(msg.bit_loc()[1])+'] ':<12}", 'red'))
        lst.append((f"{msg.field()+': '}", ('black', 'bold')))
        if msg.quick_pdo() != "Not a PDO":
            lst.append((f"{msg.quick_pdo()} ", 'purple'))
        if msg.quick_rdo() != "Not a RDO":
            lst.append((f"{msg.quick_rdo()} ", 'purple'))
        lst.append((f"(0x{int(msg.raw(), 2):0{int(len(msg.raw())/4)+(1 if len(msg.raw())%4 else 0)}X})\n", 'green'))
        for submsg in msg.value():
            renderer(submsg, level + 1, lst)
    return lst


class DataItem:
    """数据项类，表示列表中的一行数据"""
    def __init__(self, index: int, timestamp: str, sop: str, rev: str, ppr: str, pdr: str, msg_type: str, data: Any = None, time_sec: Optional[float] = None):
        self.index = index
        self.timestamp = timestamp
        # 采集时的绝对时间戳（time.time()，单位：秒），用于与曲线相对时间对齐
        self.time_sec: Optional[float] = time_sec
        self.sop = sop
        self.rev = rev
        self.ppr = ppr
        self.pdr = pdr
        self.msg_type = msg_type
        self.data = data or metadata()


class WITRNGUI:
    """WITRN HID 数据查看器主类"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.ui_scale = self._detect_ui_scale()
        # 先隐藏主窗口，等布局和几何设置完成后再显示，避免启动时小窗闪烁
        try:
            self.root.withdraw()
        except Exception:
            pass
        self.root.title("WITRN PD Sniffer v3.7.3 by JohnScotttt")
        # 使用内置的 base64 图标（brain_ico）设置窗口图标；失败则回退到本地 brain.ico
        try:
            ico_bytes = base64.b64decode(brain_ico)
            icon_tmp_path = os.path.join(tempfile.gettempdir(), "witrn_pd_sniffer_brain.ico")
            with open(icon_tmp_path, "wb") as f:
                f.write(ico_bytes)
            self.root.iconbitmap(icon_tmp_path)
        except Exception:
            pass
        # 导入字体
        try:
            jb_r_bytes = base64.b64decode(jb_r_tff)
            jb_b_bytes = base64.b64decode(jb_b_tff)
            font_tmp_path_r = os.path.join(tempfile.gettempdir(), "fonts/JetBrainsMono-Regular.ttf")
            font_tmp_path_b = os.path.join(tempfile.gettempdir(), "fonts/JetBrainsMono-Bold.ttf")
            os.makedirs(os.path.dirname(font_tmp_path_r), exist_ok=True)
            with open(font_tmp_path_r, "wb") as f:
                f.write(jb_r_bytes)
            os.makedirs(os.path.dirname(font_tmp_path_b), exist_ok=True)
            with open(font_tmp_path_b, "wb") as f:
                f.write(jb_b_bytes)
            FR_PRIVATE  = 0x10
            FR_NOT_ENUM = 0x20
            path_r = os.path.abspath(font_tmp_path_r)
            path_b = os.path.abspath(font_tmp_path_b)

            for font_path in [path_r, path_b]:
                # 注册字体
                res = ctypes.windll.gdi32.AddFontResourceExW(font_path, FR_PRIVATE, 0)
                if res == 0:
                    raise RuntimeError(f"无法加载字体文件: {font_path}")
        except Exception:
            pass

        # self.root.resizable(False, False)
        try:
            base_w, base_h = 1000, 500
            w = self._scale_size(base_w)
            h = self._scale_size(base_h)
            self.root.geometry(f"{w}x{h}")
            self.root.minsize(self._scale_size(780), self._scale_size(450))
            # self.root.maxsize(w, h)
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

        # 数据文本显示区域
        self.font_en = 'JetBrains Mono'
        self.font_cn = 'Microsoft YaHei'
        self.font_status = 'Microsoft YaHei'
        self.size_en = 9
        self.size_cn = 9
        self.size_status = 9
        try:
            with open("fonts.toml", "r", encoding="utf-8") as f:
                lines = f.readlines()
                for line in lines:
                    if line.startswith("font_en"):
                        self.font_en = line.split('=')[1].strip().strip('"').strip("'")
                    elif line.startswith("font_cn"):
                        self.font_cn = line.split('=')[1].strip().strip('"').strip("'")
                    elif line.startswith("size_en"):
                        self.size_en = int(line.split('=')[1].strip())
                    elif line.startswith("size_cn"):
                        self.size_cn = int(line.split('=')[1].strip())
                    elif line.startswith("font_status"):
                        self.font_status = line.split('=')[1].strip().strip('"').strip("'")
                    elif line.startswith("size_status"):
                        self.size_status = int(line.split('=')[1].strip())
        except Exception:
            pass
        
        # 数据存储
        self.data_list: List[DataItem] = []
        self.current_selection: Optional[DataItem] = None
        
        # 控制状态
        self.is_paused = True
        self.import_mode = False
        self.device_open = False
        
        # Treeview刷新优化
        self.last_treeview_update_time = 0.0  # 上次刷新时间
        self.pending_treeview_update = False   # 是否有待处理的刷新
        self.last_rendered_count = 0           # 上次渲染的数据条数
        self.treeview_update_job = None        # 定时刷新任务ID
        self.last_relative_time_mode = False   # 跟踪上次的相对时间模式
        self.last_filter_goodcrc_mode = False  # 跟踪上次的过滤模式
        
        # 启动数据刷新线程
        self.refresh_thread = threading.Thread(target=self.refresh_data_loop, daemon=True)
        self.refresh_thread.start()

        # 解析器实例（仅用于离线/CSV 数据解析，不负责设备连接）
        self.parser = WITRN_DEV()
        # 连接确认等待标志：启动子进程后，直到收到第一条数据才算真正“已连接”
        self.awaiting_connection_ack = False
        # 若用户在等待连接过程中按下“开始”（或 F5），连接成功后自动开始收集
        self.autostart_after_connect = False
        self.data_thread_started = False
        self.last_pdo = None
        self.last_rdo = None
        
        # ===== 多进程相关 =====
        self.collection_process = None
        self.data_queue = None  # PD数据队列
        self.iv_queue = None    # 电参数据队列
        self.stop_event = None
        self.pause_flag = None  # 共享值：0=收集中, 1=暂停
        self.queue_consumer_thread = None  # 消费队列数据的线程
        self.queue_consumer_running = False  # 消费线程运行标志

        # 创建界面
        self.create_widgets()

        # 彩蛋：全局键入“brain”触发
        self._egg_secret = "brain"
        self._egg_buffer = ""
        self._egg_activated = False
        # 彩蛋-顶部水平容器与右侧电参信息面板
        self.egg_top_row = None
        self.iv_info_frame = None
        self.iv_info_label = None
        self.iv_labels = {}
        # 每个电参标签的固定字符宽度（可按需单独调整）
        self.iv_label_char_widths = {
            'current': 14,
            'voltage': 14,
            'power': 14,
            'cc1': 14,
            'cc2': 14,
            'dp': 14,
            'dn': 14,
        }
        # 电参信息缓存（在未激活彩蛋前可先写入）
        self._iv_info_cached = {
            'current': '-',  # 电流
            'voltage': '-',  # 电压
            'power': '-',    # 功率
            'cc1': '-',      # CC1
            'cc2': '-',      # CC2
            'dp': '-',       # D+
            'dn': '-',       # D-
        }
        self.last_general_timestamp = 0.0
        try:
            # 监听全局按键（窗口任何位置）
            self.root.bind_all('<Key>', self._on_global_keypress, add='+')
            # 绑定 F5：未连接则自动连接并开始，已连接则开始收集
            self.root.bind_all('<F5>', self._on_f5_press, add='+')
            # 绑定 Shift+F5：断开连接
            self.root.bind_all('<Shift-F5>', self._on_shift_f5_press, add='+')
        except Exception:
            pass

        # 在 Windows 上尝试禁用 IME，使窗口内控件仅产生直接按键（等效锁定英文输入）
        try:
            self._install_disable_ime_hooks()
        except Exception:
            # 失败不影响主流程
            pass

        # 所有初始化完成后再显示窗口，减少启动闪烁
        try:
            # 先让 Tk 计算完布局和几何信息，再一次性显示
            self.root.update_idletasks()
            self.root.deiconify()
        except Exception:
            pass

    def create_widgets(self):
        """创建界面组件"""
        style = ttk.Style()
        style.configure("Treeview", font=(self.font_en, self.size_en), rowheight=self._scale_size(25))  # 表格文字字体
        style.configure("Treeview.Heading", font=(self.font_en, self.size_en))  # 表头文字字体
        style.map("Treeview", 
                  background=[('selected', '#0078d4')],  # 选中行背景色
                  foreground=[('selected', 'white')])    # 选中行文字颜色
        style.configure("TButton", font=(self.font_cn, self.size_cn))  # 按钮字体
        style.configure("TCheckbutton", font=(self.font_cn, self.size_cn))  # 复选框字体
        style.configure("TLabelframe.Label", font=(self.font_cn, self.size_cn))  # LabelFrame标题字体
        style.configure("TLabel", font=(self.font_status, self.size_status))  # 标签字体

        # 主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=self._scale_size(10), pady=self._scale_size(10))
        
        # 左侧列表框架
        left_frame = ttk.Frame(main_frame)
        # 保存引用，供彩蛋在最上方插入控件
        self.left_frame = left_frame
        # 固定左侧宽度为750，并禁止根据子控件自动调整大小
        try:
            left_frame.configure(width=self._scale_size(750))
        except Exception:
            pass
        try:
            left_frame.pack_propagate(False)
        except Exception:
            pass
        # 固定宽度750：仅纵向扩展，不在水平方向拉伸
        left_frame.pack(side=tk.LEFT, fill=tk.Y, expand=True, padx=(0, self._scale_size(5)))

        # 按钮与数据操作区（移动到左侧）
        button_frame = ttk.Frame(left_frame)
        # 保存引用，供彩蛋插入控件时控制相对位置
        self.button_frame = button_frame
        button_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, self._scale_size(10)))

        # 屏蔽GoodCRC 复选框（与按钮同一层级，靠右）
        self.filter_goodcrc_var = tk.BooleanVar(value=False)
        self.filter_goodcrc_cb = ttk.Checkbutton(
            button_frame,
            text="屏蔽GoodCRC",
            variable=self.filter_goodcrc_var,
            command=self.update_treeview
        )
        # 先放置右侧控件，再放置左侧按钮，有利于布局
        self.filter_goodcrc_cb.pack(side=tk.RIGHT, padx=(0, self._scale_size(5)))

        # 相对时间 复选框（放在“屏蔽GoodCRC”的左边）
        self.relative_time_var = tk.BooleanVar(value=False)
        self.relative_time_cb = ttk.Checkbutton(
            button_frame,
            text="相对时间",
            variable=self.relative_time_var,
            command=self.update_treeview
        )
        # 也使用靠右布局，后放置因此位于“屏蔽GoodCRC”的左侧
        self.relative_time_cb.pack(side=tk.RIGHT, padx=(0, self._scale_size(10)))

        # 状态栏将放到按钮区下方，见后文

        # 控制按钮框架
        control_frame = ttk.Frame(button_frame)
        control_frame.pack(side=tk.LEFT, padx=(0, self._scale_size(20)))

        # 连接按钮
        self.connect_button = ttk.Button(
            control_frame, 
            text="连接设备", 
            command=self.connect_device
        )
        self.connect_button.pack(side=tk.LEFT, padx=(0, self._scale_size(5)))
        
        # 暂停按钮
        self.pause_button = ttk.Button(
            control_frame, 
            text="开始", 
            command=self.pause_collection,
            state=tk.DISABLED
        )
        self.pause_button.pack(side=tk.LEFT, padx=(0, self._scale_size(5)))

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
        self.export_button.pack(side=tk.LEFT, padx=(0, self._scale_size(10)))

        # 导入CSV按钮
        self.import_button = ttk.Button(
            data_frame,
            text="导入CSV",
            command=self.import_csv
        )
        self.import_button.pack(side=tk.LEFT, padx=(0, self._scale_size(10)))
        
        # 清空按钮
        self.clear_button = ttk.Button(
            data_frame, 
            text="清空列表", 
            command=self.clear_list
        )
        self.clear_button.pack(side=tk.LEFT)

        # 数据列表容器（用 LabelFrame 框住 Treeview），统一与右侧 padding 保持一致
        list_group = ttk.LabelFrame(left_frame, text="数据列表", padding=self._scale_size(10))
        list_group.pack(side=tk.TOP, fill=tk.BOTH, expand=True)


        # 创建Treeview（表格）
        columns = ('Index', 'Time', 'SOP', 'Rev', 'PPR', 'PDR', 'Msg Type')
        self.tree = ttk.Treeview(list_group, columns=columns, show='headings', height=20)
        
        # 设置列标题和宽度
        column_widths = {'Index': 50, 'Time': 110, 'SOP': 90, 'Rev': 50, 'PPR': 140, 'PDR': 60, 'Msg Type': 210}
        for col in columns:
            self.tree.heading(col, text=col)
            scaled_width = self._scale_size(column_widths[col])
            # 禁止随容器自动伸缩，固定列宽
            self.tree.column(col, width=scaled_width, anchor=tk.CENTER, stretch=False)
        
        # 添加滚动条
        tree_scrollbar = ttk.Scrollbar(list_group, orient=tk.VERTICAL, command=self.tree.yview)
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
        
        # 定义本地事件处理，阻止列宽被拖拽调整（不污染类命名空间）
        def on_tree_click(event):
            try:
                region = self.tree.identify_region(event.x, event.y)
                if region == 'separator':
                    return 'break'
            except Exception:
                pass
            return None

        def on_tree_drag(event):
            try:
                region = self.tree.identify_region(event.x, event.y)
                if region == 'separator':
                    return 'break'
            except Exception:
                pass
            return None

        # 使用 grid 在列表容器内更稳定地布局，保证滚动条始终可见
        list_group.columnconfigure(0, weight=1)
        list_group.rowconfigure(0, weight=1)
        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_scrollbar.grid(row=0, column=1, sticky="ns")

        # 禁止拖动列分隔线调整列宽（绑定本地处理函数）
        self.tree.bind('<Button-1>', on_tree_click)
        self.tree.bind('<B1-Motion>', on_tree_drag)
        self.tree.bind('<Double-1>', on_tree_click)
        
        # 绑定选择事件
        self.tree.bind('<<TreeviewSelect>>', self.on_item_select)
        self.tree.bind('<Button-1>', self.on_item_click, add='+')
        
        # 右侧数据显示框架
        right_frame = ttk.LabelFrame(main_frame, text="数据显示", padding=self._scale_size(10))
        # 固定右侧宽度以配合左侧750和内边距（main_frame左右各10、左右分隔各5），此处取 820
        try:
            right_frame.configure(width=10000)
            right_frame.pack_propagate(False)
        except Exception:
            pass
        # 仅纵向扩展，宽度固定
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(self._scale_size(5), 0))
        # 保存引用，供彩蛋绘图在其下方插入曲线区
        self.right_frame = right_frame

        self.data_text = scrolledtext.ScrolledText(
            right_frame,
            wrap=tk.WORD,
            width=50,
            height=25,
            font=(self.font_en, self.size_en),
            state=tk.DISABLED,  # 初始为只读
            bd=0,
            relief='flat',
            highlightthickness=1,
            highlightbackground="#D9D9D9",  # 未聚焦时浅色线框
            highlightcolor="#BFBFBF"        # 聚焦时更明显一点
        )
        # 让文本区域位于上方，留出底部空间用于后续插入曲线
        self.data_text.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        # 注册颜色标签（Text widget 使用 tag 来控制文本样式）
        try:
            self.data_text.config(selectbackground="#b4d9fb", selectforeground="black")
            # 这里用 tag_configure 注册需要的样式名（动态字体名称）
            self.data_text.tag_configure('red', foreground='red', font=(self.font_en, self.size_en))
            self.data_text.tag_configure('blue', foreground='#476fD5', font=(self.font_en, self.size_en))
            self.data_text.tag_configure('green', foreground='green', font=(self.font_en, self.size_en))
            self.data_text.tag_configure('black', foreground='black', font=(self.font_en, self.size_en))
            self.data_text.tag_configure('gray', foreground="#474747", font=(self.font_en, self.size_en))
            self.data_text.tag_configure('purple', foreground='purple', font=(self.font_en, self.size_en))
            self.data_text.tag_configure('cn', font=(self.font_cn, self.size_cn))
            self.data_text.tag_configure('bold', font=(self.font_en, self.size_en, 'bold'))
        except Exception:
            # 在极端环境下 tag_configure 可能失败，但不影响基本功能
            pass
        
        # 底部全局状态栏容器（跨全宽）- 方便左右放置多个标签
        self.status_bar = tk.Frame(self.root, bd=0, relief='flat', highlightthickness=0, )
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # 左侧：常规状态文本
        self.status_var = tk.StringVar(value="就绪")
        self.status_label = tk.Label(
            self.status_bar,
            textvariable=self.status_var,
            anchor='w',
            bd=0,
            relief='flat',
            highlightthickness=0,
            padx=self._scale_size(8),
            pady=self._scale_size(4),
            font=(self.font_status, self.size_status)
        )
        # 左侧标签占据剩余空间
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 右侧：快速 PDO/RDO 文本（右对齐）
        self.quick_pd_var = tk.StringVar(value="")
        self.quick_pd_label = tk.Label(
            self.status_bar,
            textvariable=self.quick_pd_var,
            anchor='e',
            bd=0,
            relief='flat',
            highlightthickness=0,
            padx=self._scale_size(8),
            pady=self._scale_size(4),
            font=(self.font_status, self.size_status)
        )
        self.quick_pd_label.pack(side=tk.RIGHT)

        # 初始化状态样式
        try:
            self.set_status("就绪", level='info')
        except Exception:
            pass

        # ===== 彩蛋曲线相关（延迟初始化） =====
        self.plot_group = None            # LabelFrame 容器
        self.plot_container = None        # 固定高度的承载容器
        self.plot_canvas = None           # FigureCanvasTkAgg
        self.plot_fig = None              # matplotlib Figure
        self.plot_ax_v = None             # 电压轴（左）
        self.plot_ax_i = None             # 电流轴（右）
        self._plot_lock = threading.Lock()
        self.plot_times = deque(maxlen=60000)     # 时间戳（秒）
        self.plot_voltage = deque(maxlen=60000)   # 电压 V
        self.plot_current = deque(maxlen=60000)   # 电流 A
        self.plot_update_job = None
        self.plot_update_interval_ms = 500
        self.plot_window_seconds = 60.0          # 默认显示最近60秒
        self.plot_latest_time = 0.0              # 最新数据的时间戳
        self.plot_start_time = None              # 数据采集开始时间
        # 事件标记（PDO/RDO）
        self.marker_events = deque(maxlen=3600)  # (timestamp_sec, kind: 'pdo'|'rdo')
        self._marker_artists = []  # 兼容旧逻辑：绘制时生成的 vline 句柄
        # 新增：为事件建立 artist 映射，避免重复创建，并支持保留历史
        self._marker_artists_map = {}  # key: (x, kind) -> artist
        self.keep_marker_history = True  # 为 True 时,PDO/RDO 标线在历史中保留
        # 选中项的固定竖线（列表选中或点击锁定时使用）
        self._selected_vline_artist = None
        self._selected_text_artist = None     # 固定线旁边的文本标签
        # hover预览竖线（独立的一条，仅悬停时显示）
        self._hover_vline_artist = None
        self._hover_text_artist = None        # hover预览线旁边的文本标签
        # 鼠标悬停相关状态
        self._hover_update_job = None         # 悬停刷新任务ID
        self._hover_last_x = None             # 上次悬停的x坐标
        self._hover_last_item_index = None    # 上次悬停对应的item索引
        self._is_mouse_in_plot = False        # 鼠标是否在plot内
        self._last_click_time = 0.0           # 上次点击时间（用于区分点击与拖拽）
        self._click_start_x = None            # 点击开始时的x坐标
        self._is_vline_locked = False         # 固定竖线是否已锁定（点击后固定）

    def _on_global_keypress(self, event: tk.Event) -> None:
        """捕获全局键盘输入，用于检测彩蛋口令。"""
        if self._egg_activated:
            return
        try:
            ch = event.char or ''
        except Exception:
            ch = ''

        # 仅处理可见 ASCII 字符；退格清空缓冲以避免误触
        if event.keysym in ('BackSpace', 'Escape'):
            self._egg_buffer = ''
            return
        if len(ch) != 1 or not (32 <= ord(ch) <= 126):
            return

        # 追加并截断到口令长度
        self._egg_buffer = (self._egg_buffer + ch).lower()[-len(self._egg_secret):]
        if self._egg_buffer == self._egg_secret:
            self._egg_buffer = ''
            self._activate_easter_egg()

    def _on_f5_press(self, event: Optional[tk.Event] = None):
        """F5 快捷：
        - 未连接设备 -> 尝试连接设备，连接成功后开始收集。
        - 已连接设备 -> 若处于暂停，则开始收集；若已在收集则忽略。
        """
        try:
            # 若按下了 Shift，则交由 Shift+F5 处理
            if event is not None and (getattr(event, 'state', 0) & 0x0001):
                return None
            # 未连接：尝试连接
            if not self.device_open:
                self.connect_device()
                # 若连接成功，开始收集（避免阻塞提示后的状态不一致）
                if self.device_open and self.is_paused:
                    self.pause_collection()
                return 'break'

            # 已连接：如处于暂停则开始
            if self.is_paused:
                self.pause_collection()
            return 'break'
        except Exception:
            # 无论如何拦截默认行为，避免触发其他绑定
            return 'break'

    def _on_shift_f5_press(self, event: Optional[tk.Event] = None):
        """Shift+F5：断开连接（若当前已连接）。"""
        try:
            if self.device_open:
                # 复用切换逻辑（已连接时 connect_device 会执行断开路径）
                self.connect_device()
            else:
                self.set_status("未连接设备", level='info')
            return 'break'
        except Exception:
            return 'break'

    def _activate_easter_egg(self) -> None:
        """激活开发者界面：在左边框架最上方添加设备菜单（下拉，内容暂为空）。"""
        if self._egg_activated:
            return
        self._egg_activated = True
        try:
            # 1) 顶部水平容器：放置“设备列表”与“基本信息”并排
            if self.egg_top_row is None or not str(self.egg_top_row):
                self.egg_top_row = ttk.Frame(self.left_frame)
                try:
                    self.egg_top_row.pack(side=tk.TOP, fill=tk.X, pady=(0, 8), before=self.button_frame)
                except Exception:
                    self.egg_top_row.pack(side=tk.TOP, fill=tk.X, pady=(0, 8))
                    try:
                        self.button_frame.pack_forget()
                        self.button_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 10))
                    except Exception:
                        pass

            # 3) 在设备列表右侧创建“基本信息”面板
            if self.iv_info_frame is None or getattr(self.iv_info_frame, 'master', None) is not self.egg_top_row:
                try:
                    if self.iv_info_frame is not None:
                        self.iv_info_frame.destroy()
                except Exception:
                    pass
                self.iv_info_frame = ttk.LabelFrame(self.egg_top_row, text="基本信息", padding=self._scale_size(8))
                # 右侧面板占据剩余空间，并按Y方向填充与左侧保持同高
                self.iv_info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(self._scale_size(8), 0))
                # 放置7个横向标签（电流、电压、功率、CC1、CC2、D+、D-）
                try:
                    for child in getattr(self.iv_info_frame, 'winfo_children', lambda: [])():
                        try:
                            child.destroy()
                        except Exception:
                            pass
                except Exception:
                    pass
                self.iv_labels = {}
                display_names = {
                    'current': '电流',
                    'voltage': '电压',
                    'power': '功率',
                    'cc1': 'CC1',
                    'cc2': 'CC2',
                    'dp': 'D+',
                    'dn': 'D-',
                }
                order = ['current', 'voltage', 'power', 'cc1', 'cc2', 'dp', 'dn']
                for key in order:
                    text = f"{display_names[key]}: {self._iv_info_cached.get(key, '-')}"
                    width = int(self.iv_label_char_widths.get(key, 16))
                    lbl = ttk.Label(self.iv_info_frame, text=text, anchor='w', justify='left', width=width)
                    lbl.pack(side=tk.LEFT, padx=(0, self._scale_size(2)))
                    self.iv_labels[key] = lbl

            # 4) 首次激活后，用缓存的电参信息刷新显示
            try:
                self._refresh_iv_label_async()
            except Exception:
                pass
            # 5) 初始化并显示电压/电流曲线区（在右侧“数据显示”下方）
            try:
                self._ensure_plot_initialized()
            except Exception:
                pass
            # 反馈状态
            try:
                self.set_status("开发者模式已开启", level='egg')
            except Exception:
                pass
        except Exception:
            # 激活失败不影响主流程
            try:
                self.set_status("彩蛋激活失败", level='warn')
            except Exception:
                pass

    # ===== Windows IME 禁用（方式二）=====
    def _disable_ime_for_hwnd(self, hwnd: int) -> None:
        """对指定 HWND 禁用 IME（ImmAssociateContext(hwnd, NULL)）。仅 Windows 生效。"""
        try:
            if os.name != 'nt' or not hwnd:
                return
            imm32 = ctypes.windll.imm32  # type: ignore[attr-defined]
            # 直接传入空指针，解除该窗口的 IME 关联
            imm32.ImmAssociateContext(ctypes.c_void_p(hwnd), ctypes.c_void_p(0))
        except Exception:
            pass

    def _install_disable_ime_hooks(self) -> None:
        """为整个应用安装 IME 禁用钩子：
        - 当前窗口与其获得焦点的子控件会被禁用 IME，达到“锁定英文输入”的效果。
        - 仅在 Windows 上启用；其他平台忽略。
        """
        if os.name != 'nt':
            return

        # 先对顶层窗口禁用一次 IME
        try:
            self._disable_ime_for_hwnd(int(self.root.winfo_id()))
        except Exception:
            pass

        # 焦点切换时，对获得焦点的控件禁用 IME（覆盖后续创建的控件）
        def _on_focus_in(e: tk.Event):
            try:
                w = getattr(e, 'widget', None)
                if w is None:
                    return
                hwnd = int(w.winfo_id())
                self._disable_ime_for_hwnd(hwnd)
            except Exception:
                pass

        try:
            self.root.bind_all('<FocusIn>', _on_focus_in, add='+')
        except Exception:
            pass

    def _detect_ui_scale(self) -> float:
        """Return the OS-level Tk scaling factor (>=1.0 when unavailable)."""
        try:
            scale = float(self.root.tk.call('tk', 'scaling'))
            if scale > 0:
                return scale
        except Exception:
            pass
        return 1.0

    def _scale_size(self, value: float) -> int:
        """Scale pixel-based dimensions by the current DPI factor."""
        try:
            scale = getattr(self, 'ui_scale', 1.0)
            return max(1, int(round(float(value) * scale)))
        except Exception:
            return int(value)

    def _refresh_iv_label_async(self):
        """异步刷新右侧电参标签的内容，保证在主线程更新。"""
        try:
            self.root.after(0, self._refresh_iv_label_now)
        except Exception:
            # 兜底：直接尝试更新
            self._refresh_iv_label_now()

    def _refresh_iv_label_now(self):
        try:
            if isinstance(self.iv_labels, dict) and self.iv_labels:
                display_names = {
                    'current': '电流',
                    'voltage': '电压',
                    'power': '功率',
                    'cc1': 'CC1',
                    'cc2': 'CC2',
                    'dp': 'D+',
                    'dn': 'D-',
                }
                for key, lbl in list(self.iv_labels.items()):
                    try:
                        if lbl is not None:
                            lbl.config(text=f"{display_names.get(key, key)}: {self._iv_info_cached.get(key, '-')} ")
                    except Exception:
                        pass
        except Exception:
            pass

    def reset_iv_info(self) -> None:
        """将“基本信息”面板恢复为初始状态（全部为 '-'）。"""
        try:
            self._iv_info_cached.update({
                'current': '-',
                'voltage': '-',
                'power': '-',
                'cc1': '-',
                'cc2': '-',
                'dp': '-',
                'dn': '-',
            })
        except Exception:
            pass
        try:
            if self._egg_activated:
                self._refresh_iv_label_async()
        except Exception:
            pass

    def set_iv_info(self, current: str, voltage: str, power: str, cc1: str, cc2: str, dp: str, dn: str) -> None:
        """更新“基本信息”面板内容（全部为 str）。
        即使彩蛋未激活也会先缓存，待激活后自动渲染。
        """
        # 写入缓存
        try:
            self._iv_info_cached.update({
                'current': str(current),
                'voltage': str(voltage),
                'power': str(power),
                'cc1': str(cc1),
                'cc2': str(cc2),
                'dp': str(dp),
                'dn': str(dn),
            })
        except Exception:
            # 即使 update 失败，也不要抛出到调用方
            pass
        # 如果彩蛋已激活，刷新 UI；否则保持缓存
        if self._egg_activated:
            self._refresh_iv_label_async()

    # ===== 曲线绘图支持 =====
    def _ensure_plot_initialized(self):
        """在右侧数据显示区域底部初始化曲线LabelFrame与matplotlib画布。"""
        # 若缺少 matplotlib，放置提示标签
        if self.plot_group is not None and str(self.plot_group):
            return
        
        # 设置matplotlib的后端参数，优化性能
        plt.rcParams['path.simplify'] = True
        plt.rcParams['path.simplify_threshold'] = 1.0
        plt.rcParams['agg.path.chunksize'] = 10000
        if getattr(self, 'right_frame', None) is None:
            return
        self.plot_group = ttk.LabelFrame(self.right_frame, text="电流/电压 曲线", padding=self._scale_size(6))
        # 放在底部，不参与 expand（保持固定高度）
        try:
            self.plot_group.pack(side=tk.BOTTOM, fill=tk.X, expand=False, pady=(self._scale_size(8), 0))
        except Exception:
            self.plot_group.pack(side=tk.BOTTOM, fill=tk.X, pady=(self._scale_size(8), 0))

        # 承载画布的容器，固定高度，防止与上方文本竞争空间
        self.plot_container = tk.Frame(self.plot_group, height=self._scale_size(260))
        try:
            self.plot_container.pack_propagate(False)
        except Exception:
            pass
        self.plot_container.pack(side=tk.TOP, fill=tk.X)

        # 初始化 Figure/Axes（双Y轴：左电压V，右电流A）
        self.plot_fig = plt.Figure(figsize=(7.8, 2.2), dpi=100)
        self.plot_ax_v = self.plot_fig.add_subplot(111)
        try:
            self.plot_ax_i = self.plot_ax_v.twinx()
        except Exception:
            self.plot_ax_i = None

        self.plot_ax_v.set_ylabel("VBus (V)", color="#1f77b4")
        if self.plot_ax_i is not None:
            self.plot_ax_i.set_ylabel("Current (A)", color="#d62728")

        # 预置两条线对象（设置 zorder 确保在竖线上方）
        self._line_v, = self.plot_ax_v.plot([], [], color="#1f77b4", linewidth=1.5, label="VBus", zorder=10)
        if self.plot_ax_i is not None:
            self._line_i, = self.plot_ax_i.plot([], [], color="#d62728", linewidth=1.3, label="Current", zorder=11)
        else:
            self._line_i, = self.plot_ax_v.plot([], [], color="#d62728", linewidth=1.3, label="Current", zorder=11)

        try:
            self.plot_ax_v.grid(True, linestyle='--', alpha=0.3)
        except Exception:
            pass
        try:
            handles = [self._line_v]
            if self._line_i is not None:
                handles.append(self._line_i)
            labels = [h.get_label() for h in handles]
            ax_for_legend = self.plot_ax_i if self.plot_ax_i is not None else self.plot_ax_v
            legend = ax_for_legend.legend(handles, labels, loc='upper left')
            try:
                legend.set_zorder(10)
            except Exception:
                pass
        except Exception:
            pass
        # 优化布局，避免右侧 Y 轴标签被裁剪
        try:
            self.plot_fig.tight_layout()
        except Exception:
            pass

        # 嵌入 Tk 画布
        try:
            self.plot_canvas = FigureCanvasTkAgg(self.plot_fig, master=self.plot_container)
            self.plot_canvas_widget = self.plot_canvas.get_tk_widget()
            self.plot_canvas_widget.pack(fill=tk.BOTH, expand=True)

            # 创建自定义 matplotlib 工具栏（不显示 Configure subplots）
            from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
            class CustomToolbar(NavigationToolbar2Tk):
                # 只保留需要的按钮
                toolitems = [t for t in NavigationToolbar2Tk.toolitems if t[0] != 'Subplots']
            
            self.plot_toolbar = CustomToolbar(self.plot_canvas, self.plot_container)
            self.plot_toolbar.update()  # 更新工具栏
            # 初始状态根据设备连接状态决定是否显示工具栏
            if self.device_open:
                try:
                    self._deactivate_plot_interactions()
                finally:
                    self.plot_toolbar.pack_forget()
            else:
                self.plot_toolbar.pack(side=tk.TOP, fill=tk.X)
            self.plot_canvas._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
            
            # 绑定鼠标事件（用于hover预览与点击锁定）
            self._bind_plot_mouse_events()
        except Exception:
            # 若画布创建失败，给出提示
            try:
                for w in self.plot_container.winfo_children():
                    w.destroy()
            except Exception:
                pass
            tk.Label(self.plot_container, text="创建绘图画布失败", fg="#a4262c").pack(fill=tk.BOTH, expand=True)
            return

        # 启动定时刷新
        if self.plot_update_job is None:
            self._schedule_plot_update()

    def _schedule_plot_update(self):
        try:
            self.plot_update_job = self.root.after(self.plot_update_interval_ms, self._update_plot)
        except Exception:
            self.plot_update_job = None

    def _deactivate_plot_interactions(self) -> None:
        """取消 Matplotlib 工具栏的抓手（Pan/Zoom）模式，恢复默认光标。
        在隐藏工具栏或切换设备状态时调用，避免画布仍处于拖拽/缩放模式。
        """
        try:
            tb = getattr(self, 'plot_toolbar', None)
            if not tb:
                return
            # 依据不同版本实现，尽量安全地关闭活动模式
            try:
                active = getattr(tb, '_active', None)
            except Exception:
                active = None
            # 主动切换一次以关闭对应模式（pan/zoom 都是切换式）
            try:
                if active == 'PAN' and hasattr(tb, 'pan'):
                    tb.pan()
                elif active == 'ZOOM' and hasattr(tb, 'zoom'):
                    tb.zoom()
            except Exception:
                pass
            # 双保险：直接清空内部标志与模式文本
            try:
                if hasattr(tb, '_active'):
                    tb._active = None
            except Exception:
                pass
            try:
                if getattr(tb, 'mode', None):
                    tb.mode = ''
            except Exception:
                pass
            # 恢复画布默认光标
            try:
                canvas = getattr(self, 'plot_canvas', None)
                if canvas is not None:
                    w = canvas.get_tk_widget()
                    try:
                        w.configure(cursor='')
                    except Exception:
                        pass
            except Exception:
                pass
        except Exception:
            # 任意异常不影响主流程
            pass
    
    def _stop_plot_updates(self):
        """停止曲线的定时刷新。"""
        try:
            if self.plot_update_job is not None:
                try:
                    self.root.after_cancel(self.plot_update_job)
                except Exception:
                    pass
                self.plot_update_job = None
        except Exception:
            pass

    def _start_plot_updates(self):
        """启动曲线的定时刷新（若未启动）。"""
        try:
            if self.plot_update_job is None:
                self._schedule_plot_update()
        except Exception:
            pass

    def _reset_plot(self):
        """清空缓冲并重置坐标轴与线条，作为连接设备时的初始化。"""
        try:
            with self._plot_lock:
                try:
                    self.plot_times.clear()
                    self.plot_voltage.clear()
                    self.plot_current.clear()
                    self.plot_start_time = None  # 重置起始时间
                except Exception:
                    # 若 deque 不存在或已被销毁，忽略
                    pass
            # 线条清空
            try:
                if hasattr(self, '_line_v') and self._line_v is not None:
                    self._line_v.set_data([], [])
            except Exception:
                pass
            try:
                if hasattr(self, '_line_i') and self._line_i is not None:
                    self._line_i.set_data([], [])
            except Exception:
                pass
            # 清空事件标记与已绘制的标记线
            try:
                with self._plot_lock:
                    self.marker_events.clear()
            except Exception:
                pass
            # 彻底移除已创建的 artist（兼容旧列表与新映射）
            try:
                for a in (getattr(self, '_marker_artists', []) or []):
                    try:
                        a.remove()
                    except Exception:
                        pass
            except Exception:
                pass
            finally:
                self._marker_artists = []
            try:
                for k, a in list((getattr(self, '_marker_artists_map', {}) or {}).items()):
                    try:
                        a.remove()
                    except Exception:
                        pass
            except Exception:
                pass
            finally:
                self._marker_artists_map = {}
            # 移除当前选中竖线和文本标签并解锁
            try:
                if getattr(self, '_selected_vline_artist', None) is not None:
                    try:
                        self._selected_vline_artist.remove()
                    except Exception:
                        pass
            except Exception:
                pass
            finally:
                self._selected_vline_artist = None
            try:
                if getattr(self, '_selected_text_artist', None) is not None:
                    try:
                        self._selected_text_artist.remove()
                    except Exception:
                        pass
            except Exception:
                pass
            finally:
                self._selected_text_artist = None
                self._is_vline_locked = False
                self._hover_last_item_index = None
            # 移除hover预览竖线和文本标签
            try:
                if getattr(self, '_hover_vline_artist', None) is not None:
                    try:
                        self._hover_vline_artist.remove()
                    except Exception:
                        pass
            except Exception:
                pass
            finally:
                self._hover_vline_artist = None
            try:
                if getattr(self, '_hover_text_artist', None) is not None:
                    try:
                        self._hover_text_artist.remove()
                    except Exception:
                        pass
            except Exception:
                pass
            finally:
                self._hover_text_artist = None
            # 坐标轴复位
            try:
                if self.plot_ax_v is not None:
                    self.plot_ax_v.set_xlim(0, self.plot_window_seconds)
                    # 设定一个合理的初始电压范围（0~25V）
                    self.plot_ax_v.set_ylim(0, 25)
            except Exception:
                pass
            try:
                if self.plot_ax_i is not None:
                    # 初始电流范围（0~5A）
                    self.plot_ax_i.set_ylim(0, 5)
            except Exception:
                pass
            # 立即刷新一次
            try:
                if self.plot_canvas is not None:
                    self.plot_canvas.draw_idle()
            except Exception:
                pass
        except Exception:
            pass

    def _append_plot_point(self, t_sec: float, v: float, i: float):
        try:
            with self._plot_lock:
                # 设置起始时间（如果还未设置）
                if self.plot_start_time is None:
                    self.plot_start_time = t_sec
                # 计算相对时间（从数据开始采集时算起）
                rel_time = t_sec - self.plot_start_time
                self.plot_times.append(rel_time)
                self.plot_voltage.append(float(v))
                self.plot_current.append(float(i))
        except Exception:
            pass

    def _append_marker_event(self, t_sec: float, kind: str):
        """记录一个事件标记（PDO/RDO）。kind 取 'pdo' 或 'rdo'。"""
        try:
            if kind not in ('pdo', 'rdo'):
                return
            with self._plot_lock:
                # 使用相对时间记录事件
                if self.plot_start_time is not None:
                    rel_time = float(t_sec) - self.plot_start_time
                    self.marker_events.append((rel_time, kind))
        except Exception:
            pass

    def _get_item_plot_x(self, item: 'DataItem') -> Optional[float]:
        """根据选中项计算其在曲线中的X坐标（相对时间）。
        优先使用采集时的绝对时间戳time_sec与plot_start_time对齐；
        如无time_sec，则尝试解析文本时间并与plot_start_time的当天秒数对齐，
        处理跨天小概率场景（±12小时阈值修正）。
        """
        try:
            if item is None or self.plot_start_time is None:
                return None
            # 首选：绝对时间
            if getattr(item, 'time_sec', None) is not None:
                return float(item.time_sec) - float(self.plot_start_time)

            # 兜底：解析文本时间，按“当天秒”对齐
            ts_sec = self._parse_timestamp_to_seconds(getattr(item, 'timestamp', None))
            if ts_sec is None:
                return None
            lt = time.localtime(float(self.plot_start_time))
            plot_day_sec = lt.tm_hour * 3600 + lt.tm_min * 60 + lt.tm_sec + (float(self.plot_start_time) % 1.0)
            rel = ts_sec - plot_day_sec
            # 简单跨天修正：若差异超过±12小时，则加/减24小时
            if rel < -43200:
                rel += 86400
            elif rel > 43200:
                rel -= 86400
            return rel
        except Exception:
            return None

    def _update_selection_vline(self, item: Optional['DataItem']) -> None:
        """在曲线里绘制/更新当前选中的竖线：
        - 颜色与左侧消息类型背景色一致
        - 线宽为3
        - 每次仅保留一条，切换选择时删除旧线
        """
        try:
            # 若未初始化曲线或未启用彩蛋，或处于导入模式，则忽略
            if not self._egg_activated or self.plot_ax_v is None or self.plot_canvas is None or self.import_mode:
                return
            # 先移除旧的选中竖线
            try:
                if self._selected_vline_artist is not None:
                    try:
                        self._selected_vline_artist.remove()
                    except Exception:
                        pass
            finally:
                self._selected_vline_artist = None

            # 移除旧的文本标签
            try:
                if self._selected_text_artist is not None:
                    try:
                        self._selected_text_artist.remove()
                    except Exception:
                        pass
            finally:
                self._selected_text_artist = None

            if item is None:
                # 仅移除
                try:
                    self.plot_canvas.draw_idle()
                except Exception:
                    pass
                return

            # 计算X坐标（相对时间）
            x = self._get_item_plot_x(item)
            if x is None:
                return

            # 计算颜色：与左侧列表相同的颜色（基于消息类型）
            color_hex = None
            try:
                if isinstance(item.msg_type, str) and item.msg_type in MT:
                    color_hex = MT[item.msg_type]
                elif isinstance(item.msg_type, str):
                    lowered = item.msg_type.lower()
                    for k, v in MT.items():
                        if k.lower() == lowered:
                            color_hex = v
                            break
            except Exception:
                color_hex = None
            if not color_hex:
                color_hex = '#000000'

            # 绘制新的竖线（黑色细线，放在最底层）
            try:
                self._selected_vline_artist = self.plot_ax_v.axvline(
                    x=x, color='#000000', linewidth=1, alpha=0.9, linestyle='-', zorder=0)
            except Exception:
                self._selected_vline_artist = None
            
            # 在竖线旁边添加消息类型文本标签
            try:
                # 获取消息类型文本
                msg_text = str(item.msg_type) if item.msg_type else "Unknown"
                
                # 获取当前Y轴范围，将文本放在顶部
                try:
                    ylim = self.plot_ax_v.get_ylim()
                    y_pos = ylim[1] * 0.95  # 放在Y轴95%的位置
                except Exception:
                    y_pos = 20  # 默认位置
                
                # 绘制文本（带半透明白色背景框，显示在竖线右边）
                self._selected_text_artist = self.plot_ax_v.text(
                    x, y_pos, msg_text,
                    rotation=90,  # 竖直显示
                    verticalalignment='top',
                    horizontalalignment='left',  # 文本左边对齐到x（竖直时，文本在竖线右边）
                    fontsize=9,  # 固定线的文本稍大一点
                    color="black",
                    bbox=dict(
                        boxstyle='round,pad=0.4',
                        facecolor=color_hex,
                        edgecolor=color_hex,
                        alpha=0.7,  # 固定线的背景更不透明
                        linewidth=1.5
                    ),
                    zorder=2  # 比hover预览线的文本更高一层
                )
            except Exception as e:
                print(f"[DEBUG] 添加固定线文本标签失败: {e}")
                self._selected_text_artist = None

            try:
                self.plot_canvas.draw_idle()
            except Exception:
                pass
        except Exception:
            pass

    def _update_hover_preview_vline(self, item: Optional['DataItem']) -> None:
        """更新hover预览竖线（独立的一条，与固定竖线不同）。
        - 颜色与报文类型一致
        - 线宽为2（比固定竖线细一点）
        - alpha=0.6（半透明，区别于固定竖线）
        - 虚线样式（进一步区分）
        """
        try:
            if not self._egg_activated or self.plot_ax_v is None or self.plot_canvas is None or self.import_mode:
                return
            
            # 先移除旧的hover预览线
            try:
                if self._hover_vline_artist is not None:
                    try:
                        self._hover_vline_artist.remove()
                    except Exception:
                        pass
            finally:
                self._hover_vline_artist = None
            
            # 移除旧的文本标签
            try:
                if self._hover_text_artist is not None:
                    try:
                        self._hover_text_artist.remove()
                    except Exception:
                        pass
            except Exception:
                pass
            finally:
                self._hover_text_artist = None
            
            if item is None:
                try:
                    self.plot_canvas.draw_idle()
                except Exception:
                    pass
                return
            
            # 计算X坐标
            x = self._get_item_plot_x(item)
            if x is None:
                return
            
            # 计算颜色
            color_hex = None
            try:
                if isinstance(item.msg_type, str) and item.msg_type in MT:
                    color_hex = MT[item.msg_type]
                elif isinstance(item.msg_type, str):
                    lowered = item.msg_type.lower()
                    for k, v in MT.items():
                        if k.lower() == lowered:
                            color_hex = v
                            break
            except Exception:
                color_hex = None
            if not color_hex:
                color_hex = '#000000'
            
            # 绘制hover预览竖线（黑色细线、半透明、虚线，放在最底层）
            try:
                self._hover_vline_artist = self.plot_ax_v.axvline(
                    x=x, color='#000000', linewidth=1, alpha=0.6, linestyle='--', zorder=0)
            except Exception:
                self._hover_vline_artist = None
            
            # 在竖线旁边添加消息类型文本标签
            try:
                # 获取消息类型文本
                msg_text = str(item.msg_type) if item.msg_type else "Unknown"
                
                # 获取当前Y轴范围，将文本放在顶部
                try:
                    ylim = self.plot_ax_v.get_ylim()
                    y_pos = ylim[1] * 0.95  # 放在Y轴95%的位置
                except Exception:
                    y_pos = 20  # 默认位置
                
                # 绘制文本（带半透明白色背景框，显示在竖线右边）
                self._hover_text_artist = self.plot_ax_v.text(
                    x, y_pos, msg_text,
                    rotation=90,  # 竖直显示
                    verticalalignment='top',
                    horizontalalignment='left',  # 文本左边对齐到x（竖直时，文本在竖线右边）
                    fontsize=8,
                    color="black",
                    bbox=dict(
                        boxstyle='round,pad=0.3',
                        facecolor=color_hex,
                        edgecolor=color_hex,
                        alpha=0.7,
                        linewidth=1
                    ),
                    zorder=1  # 确保在最上层
                )
            except Exception as e:
                print(f"[DEBUG] 添加文本标签失败: {e}")
                self._hover_text_artist = None
            
            try:
                self.plot_canvas.draw_idle()
            except Exception:
                pass
        except Exception:
            pass

    def _bind_plot_mouse_events(self) -> None:
        """绑定plot的鼠标事件，用于hover预览和点击锁定。"""
        try:
            if self.plot_canvas is None:
                print("[DEBUG] plot_canvas为None，无法绑定事件")
                return
            print("[DEBUG] 开始绑定plot鼠标事件")
            # motion_notify_event: 鼠标移动
            self.plot_canvas.mpl_connect('motion_notify_event', self._on_plot_mouse_move)
            # button_press_event: 鼠标按下
            self.plot_canvas.mpl_connect('button_press_event', self._on_plot_mouse_press)
            # button_release_event: 鼠标释放
            self.plot_canvas.mpl_connect('button_release_event', self._on_plot_mouse_release)
            # axes_enter_event/axes_leave_event: 进入/离开坐标轴区域
            self.plot_canvas.mpl_connect('axes_enter_event', self._on_plot_mouse_enter)
            self.plot_canvas.mpl_connect('axes_leave_event', self._on_plot_mouse_leave)
            print("[DEBUG] plot鼠标事件绑定完成")
        except Exception as e:
            print(f"[DEBUG] 绑定事件异常: {e}")
            import traceback
            traceback.print_exc()
            pass

    def _on_plot_mouse_enter(self, event) -> None:
        """鼠标进入plot区域，启动hover刷新。"""
        try:
            print(f"[DEBUG] 鼠标进入plot区域")
            self._is_mouse_in_plot = True
            # 启动5Hz刷新（200ms间隔）
            if self._hover_update_job is None:
                self._schedule_hover_update()
                print(f"[DEBUG] 启动hover刷新")
        except Exception as e:
            print(f"[DEBUG] 进入事件异常: {e}")
            pass

    def _on_plot_mouse_leave(self, event) -> None:
        """鼠标离开plot区域，停止hover刷新并清除预览线。"""
        try:
            print(f"[DEBUG] 鼠标离开plot区域，开始清理")
            self._is_mouse_in_plot = False
            # 停止hover刷新
            if self._hover_update_job is not None:
                try:
                    self.root.after_cancel(self._hover_update_job)
                except Exception:
                    pass
                self._hover_update_job = None
            # 清除hover预览竖线和文本标签
            self._update_hover_preview_vline(None)
            # 清除hover状态
            self._hover_last_x = None
            self._hover_last_item_index = None
            print(f"[DEBUG] hover预览清理完成")
        except Exception as e:
            print(f"[DEBUG] 离开事件异常: {e}")
            pass

    def _on_plot_mouse_move(self, event) -> None:
        """鼠标移动事件，记录当前x坐标供hover更新使用。"""
        try:
            # 检查是否在任一坐标轴内（电压轴或电流轴都可以）
            if event.inaxes in (self.plot_ax_v, self.plot_ax_i) and event.xdata is not None:
                self._hover_last_x = float(event.xdata)
                # print(f"[DEBUG] 鼠标x={self._hover_last_x:.2f}")
        except Exception as e:
            print(f"[DEBUG] 移动事件异常: {e}")
            import traceback
            traceback.print_exc()
            pass

    def _on_plot_mouse_press(self, event) -> None:
        """鼠标按下，记录点击位置和时间。"""
        try:
            # 检查是否在任一坐标轴内
            if event.inaxes in (self.plot_ax_v, self.plot_ax_i) and event.xdata is not None:
                self._click_start_x = float(event.xdata)
                self._last_click_time = time.time()
                print(f"[DEBUG] 鼠标按下 x={self._click_start_x:.2f}")
        except Exception:
            pass

    def _on_plot_mouse_release(self, event) -> None:
        """鼠标释放，判断是点击还是拖拽：
        - 点击：释放位置与按下位置接近（<0.05s窗口内移动<视窗宽度5%），则锁定当前hover的竖线并同步选中列表
        - 拖拽：视为工具栏操作（pan/zoom），不触发锁定
        """
        try:
            # 导入模式下不处理点击
            if self.import_mode:
                return
            # 检查是否在任一坐标轴内
            if event.inaxes not in (self.plot_ax_v, self.plot_ax_i) or event.xdata is None:
                return
            if self._click_start_x is None:
                return
            
            # 检查工具栏是否处于激活状态（Pan或Zoom模式）
            toolbar_mode = None
            try:
                if hasattr(self, 'plot_toolbar') and self.plot_toolbar is not None:
                    toolbar_mode = getattr(self.plot_toolbar, 'mode', None)
            except Exception:
                pass
            
            # 如果工具栏处于Pan或Zoom模式，不触发点击锁定
            if toolbar_mode in ('pan/zoom', 'zoom rect'):
                print(f"[DEBUG] 工具栏模式激活: {toolbar_mode}，跳过点击锁定")
                return
            
            # 计算点击持续时间和移动距离
            click_duration = time.time() - self._last_click_time
            x_release = float(event.xdata)
            x_diff = abs(x_release - self._click_start_x)
            
            # 获取当前视窗宽度
            try:
                xlim = self.plot_ax_v.get_xlim()
                x_range = abs(xlim[1] - xlim[0])
            except Exception:
                x_range = self.plot_window_seconds
            
            # 判断是否为点击（而非拖拽）：时间<0.5s且移动距离<视窗宽度的5%
            is_click = (click_duration < 0.5) and (x_diff < x_range * 0.05)
            
            print(f"[DEBUG] 鼠标释放 x={x_release:.2f}, diff={x_diff:.3f}, duration={click_duration:.3f}s, is_click={is_click}, toolbar_mode={toolbar_mode}")
            
            if is_click and self._hover_last_item_index is not None:
                # 锁定并绘制固定竖线
                self._is_vline_locked = True
                print(f"[DEBUG] 锁定竖线，选中列表项 {self._hover_last_item_index}")
                
                # 清除hover预览线
                self._update_hover_preview_vline(None)
                
                # 获取对应的item并绘制固定竖线
                if self._hover_last_item_index < len(self.data_list):
                    locked_item = self.data_list[self._hover_last_item_index]
                    self._update_selection_vline(locked_item)
                
                # 同步选中左侧列表中的对应项
                self._select_tree_item_by_index(self._hover_last_item_index)
                # 自动聚焦（当设备断开、非导入模式且开发者模式开启时）
                try:
                    if (not self.device_open) and (not self.import_mode) and getattr(self, '_egg_activated', False):
                        if self._hover_last_item_index < len(self.data_list):
                            item = self.data_list[self._hover_last_item_index]
                            x = self._get_item_plot_x(item)
                            if x is not None:
                                self._focus_on_time_x(x, window_seconds=2.0)
                except Exception:
                    pass
        except Exception:
            pass
        finally:
            self._click_start_x = None

    def _schedule_hover_update(self) -> None:
        """启动hover刷新定时任务（5Hz = 200ms）。"""
        try:
            if self._is_mouse_in_plot:
                self._update_hover_vline()
                self._hover_update_job = self.root.after(200, self._schedule_hover_update)
        except Exception:
            self._hover_update_job = None

    def _update_hover_vline(self) -> None:
        """根据当前鼠标x坐标，查找最接近的报文并更新hover预览竖线。
        注意：即使固定竖线已锁定，hover预览线仍然会显示（两条独立的线）。
        """
        try:
            # 导入模式下不更新hover
            if self.import_mode:
                return
            # 检查工具栏是否处于激活状态，如果是则不更新hover
            try:
                if hasattr(self, 'plot_toolbar') and self.plot_toolbar is not None:
                    toolbar_mode = getattr(self.plot_toolbar, 'mode', None)
                    if toolbar_mode in ('pan/zoom', 'zoom rect'):
                        # print(f"[DEBUG] 工具栏激活，跳过hover")
                        return
            except Exception:
                pass
            if self._hover_last_x is None:
                # print(f"[DEBUG] 没有鼠标x坐标")
                return
            if not self.data_list:
                # print(f"[DEBUG] 数据列表为空")
                return
            if self.plot_start_time is None:
                # print(f"[DEBUG] plot_start_time未设置")
                return
            
            # 查找时间最接近的报文
            closest_item = None
            closest_index = None
            min_diff = float('inf')
            
            for idx, item in enumerate(self.data_list):
                item_x = self._get_item_plot_x(item)
                if item_x is not None:
                    diff = abs(item_x - self._hover_last_x)
                    if diff < min_diff:
                        min_diff = diff
                        closest_item = item
                        closest_index = idx
            
            # 如果找到且与上次不同，更新hover预览竖线
            if closest_item is not None and closest_index != self._hover_last_item_index:
                print(f"[DEBUG] 找到最近报文 idx={closest_index}, msg={closest_item.msg_type}, diff={min_diff:.3f}s")
                self._hover_last_item_index = closest_index
                # 绘制hover预览竖线（独立的一条）
                self._update_hover_preview_vline(closest_item)
        except Exception as e:
            print(f"[DEBUG] hover更新异常: {e}")
            import traceback
            traceback.print_exc()
            pass

    def _select_tree_item_by_index(self, index: int) -> None:
        """根据data_list索引，在左侧treeview中选中对应项。"""
        try:
            if not (0 <= index < len(self.data_list)):
                return
            item = self.data_list[index]
            # 在treeview中查找index+1匹配的行（第0列为序号）
            for child in self.tree.get_children():
                values = self.tree.item(child)['values']
                if values and int(values[0]) == item.index:
                    # 选中该行
                    self.tree.selection_set(child)
                    self.tree.focus(child)
                    self.tree.see(child)
                    # 更新当前选择和右侧显示
                    self.current_selection = item
                    self.display_data(item)
                    break
        except Exception:
            pass

    def _focus_on_time_x(self, x: float, window_seconds: float = 2.0) -> None:
        """将曲线视窗聚焦到给定相对时间x附近，窗口宽度为window_seconds（默认2s）。
        仅调整X轴范围，保持Y轴范围不变（避免标签跑到画面外）。
        仅做一次性调整，不会持续强制（不修改任何自动刷新策略）。
        """
        try:
            if self.plot_ax_v is None or self.plot_canvas is None:
                return
            half = max(0.1, float(window_seconds) / 2.0)
            xmin = x - half
            xmax = x + half
            # 设置X轴范围
            try:
                self.plot_ax_v.set_xlim(xmin, xmax)
            except Exception:
                pass

            # 刷新一次画布
            try:
                self.plot_canvas.draw_idle()
            except Exception:
                pass
        except Exception:
            pass

    def _update_plot(self):
        try:
            if self.plot_canvas is None or self.plot_ax_v is None:
                return
            
            with self._plot_lock:
                if not self.plot_times:
                    return
                    
                # 获取所有数据
                t_max = max(self.plot_times)
                # 始终显示最新的60秒数据
                t_min = t_max - float(self.plot_window_seconds)
                
                # 直接使用实际时间值，不再转换为相对值
                xs = list(self.plot_times)
                vs = list(self.plot_voltage)
                is_ = list(self.plot_current)

            # 更新数据
            self._line_v.set_data(xs, vs)
            if self.plot_ax_i is not None:
                self._line_i.set_data(xs, is_)
            else:
                self._line_i.set_data(xs, is_)

            # 更新时间轴范围
            try:
                # 显示实际时间范围，保持60秒的窗口宽度
                self.plot_ax_v.set_xlim(t_min, t_max)
            except Exception:
                pass
            try:
                if vs:
                    vmin = min(vs)
                    vmax = max(vs)
                    if vmin == vmax:
                        vmin -= 0.5
                        vmax += 0.5
                    self.plot_ax_v.set_ylim(vmin - 0.2, vmax + 0.2)
            except Exception:
                pass
            try:
                if is_:
                    imin = min(is_)
                    imax = max(is_)
                    if imin == imax:
                        # 若为常数线，先给出基本范围
                        imin -= 0.1
                        imax += 0.1
                    rng = max(imax - imin, 0.2)
                    # 增加底部留白，让曲线整体更靠下显示
                    pad_bottom = max(0.1, 0.4 * rng)
                    pad_top = max(0.05, 0.12 * rng)
                    if self.plot_ax_i is not None:
                        self.plot_ax_i.set_ylim(imin - pad_bottom, imax + pad_top)
            except Exception:
                pass

            try:
                self.plot_canvas.draw_idle()
            except Exception:
                pass
            # 更新 PDO/RDO 纵向标记线
            try:
                # 基于开关决定是仅绘制窗口内事件，还是持久化历史事件
                if not self.plot_times:
                    return
                t_max = max(self.plot_times)
                t_min = t_max - float(self.plot_window_seconds)
                with self._plot_lock:
                    if self.keep_marker_history:
                        evts = list(self.marker_events)
                    else:
                        evts = [ev for ev in self.marker_events if ev[0] >= t_min]

                # 懒创建：仅为尚未创建 artist 的事件创建一次，避免重复（放在最底层）
                for x, kind in evts:
                    key = (float(x), str(kind))
                    if key in self._marker_artists_map:
                        continue
                    try:
                        if kind == 'pdo':
                            art = self.plot_ax_v.axvline(x=x, color="#093E72", linewidth=1, alpha=0.5, linestyle='-', zorder=0)
                        else:
                            art = self.plot_ax_v.axvline(x=x, color="#ff69b4", linewidth=1, alpha=0.5, linestyle='-', zorder=0)
                        self._marker_artists_map[key] = art
                    except Exception:
                        pass
                try:
                    self.plot_canvas.draw_idle()
                except Exception:
                    pass
            except Exception:
                pass
        finally:
            # 继续定时刷新
            try:
                self._schedule_plot_update()
            except Exception:
                pass

    def set_status(self, text: str, level: str = 'info') -> None:
        """设置状态文本并根据级别调整状态栏颜色。
        level: info | ok | busy | warn | error
        """
        styles = {
            'info':  {'bg': "#d4d4d4", 'fg': "#353535"},
            'ok':    {'bg': "#bbe7c2", 'fg': "#0a5c0a"},
            'busy':  {'bg': "#b1cee9", 'fg': "#06315c"},
            'warn':  {'bg': "#ece1b8", 'fg': "#745203"},
            'error': {'bg': "#f0a4aa", 'fg': "#79161b"},
            'egg':   {'bg': "#c698ec", 'fg': "#490696"},
        }
        style = styles.get(level, styles['info'])
        try:
            self.status_var.set(text)
            if hasattr(self, 'status_label') and self.status_label:
                # 仅在颜色变化时更新以避免频繁重绘
                current_bg = self.status_label.cget('bg')
                current_fg = self.status_label.cget('fg')
                if current_bg != style['bg'] or current_fg != style['fg']:
                    self.status_label.configure(bg=style['bg'], fg=style['fg'])
            # 同步右侧快速信息标签的配色
            if hasattr(self, 'quick_pd_label') and self.quick_pd_label:
                q_bg = self.quick_pd_label.cget('bg')
                q_fg = self.quick_pd_label.cget('fg')
                if q_bg != style['bg'] or q_fg != style['fg']:
                    self.quick_pd_label.configure(bg=style['bg'], fg=style['fg'])
            # 同步状态栏容器背景色
            if hasattr(self, 'status_bar') and self.status_bar:
                if self.status_bar.cget('bg') != style['bg']:
                    self.status_bar.configure(bg=style['bg'])
        except Exception:
            pass

    def set_quick_pdo_rdo(self, pdo: metadata, rdo: metadata, force: bool = False) -> None:
        """更新底部状态栏右侧的快速 PDO/RDO 文本。
        """
        if self.is_paused and not force:
            return
        text = ""
        if pdo is not None:
            if not pdo["Message Header"]["Extended"].value():
                DO = pdo[3].value()
                for i, obj in enumerate(DO):
                    if obj.quick_pdo() != "Not a PDO":
                        text += f" [{i+1}] {obj.quick_pdo()} |"
            else:
                DO = pdo[4].value()
                if DO != None and DO != "Incomplete Data":
                    for i, obj in enumerate(DO):
                        if obj.quick_pdo() != "Not a PDO":
                            text += f" [{i+1}] {obj.quick_pdo()} |"
        if rdo is not None:
            DO = rdo[3].value()
            if DO == "Invalid Request Message":
                text += "| Invalid RDO"
            else:
                text += f"| {DO[0].quick_rdo()}"

        try:
            self.quick_pd_var.set(text)
        except Exception:
            pass

    def add_data_item(self,
                      sop: str,
                      rev: str,
                      ppr: str,
                      pdr: str,
                      msg_type: str,
                      data: Any = None,
                      timestamp: Optional[str] = None,
                      time_sec: Optional[float] = None,
                      force: bool = False):
        """添加新的数据项到列表"""
        # 只有在未暂停时才添加数据，除非强制添加
        if not force and self.is_paused:
            return
            
        timestamp = timestamp or datetime.now().strftime("%H:%M:%S.%f")[:-3]
        index = len(self.data_list) + 1

        item = DataItem(index, timestamp, sop, rev, ppr, pdr, msg_type, data, time_sec=time_sec)
        self.data_list.append(item)
        
        # 更新状态
        self.set_status(f"已读取 {len(self.data_list)} 条数据", level='busy')

    def start_data_thread_if_needed(self):
        """在设备可用时安全地启动数据采集进程一次。"""
        if not self.device_open or not self.parser:
            return
        if self.data_thread_started:
            return

        # 启动多进程数据采集
        try:
            # 创建进程间通信组件
            self.data_queue = Queue(maxsize=1000)  # PD数据队列
            self.iv_queue = Queue(maxsize=500)     # 电参数据队列
            self.stop_event = Event()
            self.pause_flag = Value('i', 1)  # 初始为暂停状态
            
            # 启动数据采集进程
            self.collection_process = Process(
                target=data_collection_worker,
                args=(self.data_queue, self.iv_queue, 
                      self.stop_event, self.pause_flag),
                daemon=True
            )
            self.collection_process.start()
            
            # 启动消费队列的线程（在主进程中）- 只启动一次
            if self.queue_consumer_thread is None or not self.queue_consumer_thread.is_alive():
                self.queue_consumer_running = True
                self.queue_consumer_thread = threading.Thread(
                    target=self._consume_queue_data,
                    daemon=True
                )
                self.queue_consumer_thread.start()
            
            self.data_thread_started = True
            
        except Exception as e:
            messagebox.showerror("启动失败", f"无法启动数据采集进程：{e}")
            self.set_status("启动数据采集失败", level='error')
    
    def _consume_queue_data(self):
        """消费队列中的数据（在主进程的后台线程中运行）"""
        while self.queue_consumer_running:
            try:
                # 检查队列是否已被清理
                if self.data_queue is None and self.iv_queue is None:
                    time.sleep(0.1)
                    continue
                
                # 处理PD数据队列
                if self.data_queue is not None:
                    try:
                        while True:
                            pd_data = self.data_queue.get_nowait()
                            # 先检查错误信号，避免误报“设备已连接”
                            if 'error' in pd_data:
                                err = pd_data.get('error')
                                if err == 'device_disconnected':
                                    self.root.after(0, self._handle_device_disconnect)
                                elif isinstance(err, str) and err.startswith('connection_failed'):
                                    # 连接失败：回退UI并提示
                                    self.root.after(0, lambda e=err: self._handle_connection_failed(e))
                                # 无论何种错误，跳出本轮读取，避免死循环
                                break
                            # 非错误的第一条数据才确认连接成功
                            if self.device_open and self.awaiting_connection_ack:
                                try:
                                    self.set_status("设备已连接", level='ok')
                                except Exception:
                                    pass
                                self.awaiting_connection_ack = False
                                # 若用户请求了“连接后自动开始”，此处自动开始收集
                                try:
                                    if self.autostart_after_connect:
                                        self.is_paused = False
                                        if self.pause_flag is not None:
                                            self.pause_flag.value = 0
                                        try:
                                            self.pause_button.config(text="暂停")
                                        except Exception:
                                            pass
                                        self.set_status("数据收集中...", level='busy')
                                        self.autostart_after_connect = False
                                except Exception:
                                    pass
                            
                            # 添加数据项
                            self.add_data_item(
                                pd_data['sop'],
                                pd_data['rev'],
                                pd_data['ppr'],
                                pd_data['pdr'],
                                pd_data['msg_type'],
                                pd_data['data'],
                                pd_data['timestamp'],
                                time_sec=pd_data.get('time_sec')
                            )
                            
                            # 更新PDO/RDO
                            if pd_data['is_pdo']:
                                self.last_pdo = pd_data['last_pdo']
                                try:
                                    self._append_marker_event(pd_data['time_sec'], 'pdo')
                                except:
                                    pass
                            
                            if pd_data['is_rdo']:
                                self.last_rdo = pd_data['last_rdo']
                                try:
                                    self._append_marker_event(pd_data['time_sec'], 'rdo')
                                except:
                                    pass
                            
                            self.set_quick_pdo_rdo(self.last_pdo, self.last_rdo)
                            
                    except queue.Empty:
                        pass
                    except Exception as e:
                        if self.queue_consumer_running:  # 只在运行时报错
                            print(f"处理PD数据时出错: {e}")
                
                # 处理电参数据队列
                if self.iv_queue is not None:
                    try:
                        while True:
                            iv_data = self.iv_queue.get_nowait()
                            # 首次收到数据，确认连接成功
                            if self.device_open and self.awaiting_connection_ack:
                                try:
                                    self.set_status("设备已连接", level='ok')
                                except Exception:
                                    pass
                                self.awaiting_connection_ack = False
                                # 若用户请求了“连接后自动开始”，此处自动开始收集
                                try:
                                    if self.autostart_after_connect:
                                        self.is_paused = False
                                        if self.pause_flag is not None:
                                            self.pause_flag.value = 0
                                        try:
                                            self.pause_button.config(text="暂停")
                                        except Exception:
                                            pass
                                        self.set_status("数据收集中...", level='busy')
                                        self.autostart_after_connect = False
                                except Exception:
                                    pass
                            
                            # plot数据点：无限制更新
                            if iv_data.get('update_plot', False):
                                try:
                                    self._append_plot_point(
                                        iv_data['timestamp'],
                                        iv_data['voltage'],
                                        iv_data['current']
                                    )
                                except:
                                    pass
                            
                            # iv_info显示：限制在10Hz
                            if iv_data.get('update_iv_info', False):
                                if self.device_open:
                                    self.set_iv_info(
                                        f"{iv_data['current']:.3f}A",
                                        f"{iv_data['voltage']:.3f}V",
                                        f"{iv_data['power']:.3f}W",
                                        f"{iv_data['cc1']:.1f}V",
                                        f"{iv_data['cc2']:.1f}V",
                                        f"{iv_data['dp']:.2f}V",
                                        f"{iv_data['dn']:.2f}V"
                                    )
                            
                    except queue.Empty:
                        pass
                    except Exception as e:
                        if self.queue_consumer_running:  # 只在运行时报错
                            print(f"处理电参数据时出错: {e}")
                
                # 短暂休眠避免空转
                time.sleep(0.01)
                
            except Exception as e:
                if self.queue_consumer_running:  # 只在运行时报错
                    print(f"消费队列数据时出错: {e}")
                time.sleep(0.1)
    
    def _handle_device_disconnect(self):
        """处理设备断开（在主线程中调用）"""
        try:
            self.is_paused = True
            self.device_open = False
            self.awaiting_connection_ack = False
            self.autostart_after_connect = False
            self.last_pdo = None
            self.last_rdo = None
            self.set_quick_pdo_rdo(None, None, force=True)
            self.set_status("设备断开", level='error')
            self.pause_button.config(text="开始", state=tk.DISABLED)
            self.connect_button.config(text="连接设备", state=tk.NORMAL)
            
            # 重置基本信息面板
            try:
                self.reset_iv_info()
            except:
                pass
            
            # 停止曲线刷新
            try:
                self._stop_plot_updates()
            except:
                pass
            
            # 停止数据采集进程
            self._stop_collection_process()
            
            # 弹窗提示
            try:
                messagebox.showwarning("设备断开", "检测到设备已断开，请重连或检查连接。")
            except:
                pass
                
        except Exception as e:
            print(f"处理设备断开时出错: {e}")

    def _handle_connection_failed(self, err_msg: str):
        """处理连接失败：回退UI状态并清理资源。"""
        try:
            # 清理采集进程（若已启动会很快退出）
            self._stop_collection_process()
        except Exception:
            pass
        try:
            self.device_open = False
            self.is_paused = True
            self.awaiting_connection_ack = False
            self.autostart_after_connect = False
            self.pause_button.config(text="开始", state=tk.DISABLED)
            self.connect_button.config(text="连接设备", state=tk.NORMAL)
            if hasattr(self, 'plot_toolbar') and self.plot_toolbar:
                self.plot_toolbar.pack(side=tk.TOP, fill=tk.X)
            # 停止曲线刷新，重置iv显示
            try:
                self._stop_plot_updates()
            except Exception:
                pass
            try:
                self.reset_iv_info()
            except Exception:
                pass
            # 状态与提示
            self.set_status("连接设备失败", level='error')
            try:
                messagebox.showerror("连接失败", f"无法连接到设备：{err_msg}")
            except Exception:
                pass
        except Exception as e:
            print(f"处理连接失败时出错: {e}")

    def _stop_collection_process(self):
        """停止数据采集进程"""
        try:
            # 先停止采集进程
            if self.stop_event is not None:
                self.stop_event.set()
            
            if self.collection_process is not None and self.collection_process.is_alive():
                self.collection_process.join(timeout=2.0)
                if self.collection_process.is_alive():
                    self.collection_process.terminate()
                self.collection_process = None
            
            # 标记数据采集线程已停止（但保持消费线程运行）
            self.data_thread_started = False
            
            # 清理队列（在清理前等待一小段时间让消费线程处理完剩余数据）
            time.sleep(0.05)
            
            if self.data_queue is not None:
                try:
                    while not self.data_queue.empty():
                        self.data_queue.get_nowait()
                except:
                    pass
                # 不要立即设为None，让消费线程能检测到空队列
                # self.data_queue = None  # 移除这行
            
            if self.iv_queue is not None:
                try:
                    while not self.iv_queue.empty():
                        self.iv_queue.get_nowait()
                except:
                    pass
                # 不要立即设为None，让消费线程能检测到空队列
                # self.iv_queue = None  # 移除这行
            
            self.stop_event = None
            self.pause_flag = None
            
        except Exception as e:
            print(f"停止采集进程时出错: {e}")
    
    def update_treeview(self):
        """更新Treeview显示（优化版：支持增量更新）"""
        current_data_count = len(self.data_list)
        
        # 获取当前的过滤条件和相对时间模式
        hide_goodcrc = bool(getattr(self, 'filter_goodcrc_var', tk.BooleanVar()).get())
        relative_mode = bool(getattr(self, 'relative_time_var', tk.BooleanVar()).get())
        
        # 判断是否需要完全重建（过滤条件变化、数据减少等情况）
        need_full_rebuild = False
        
        # 检查是否是清空操作
        if current_data_count == 0:
            for child in self.tree.get_children():
                self.tree.delete(child)
            self.last_rendered_count = 0
            self.last_relative_time_mode = relative_mode
            self.last_filter_goodcrc_mode = hide_goodcrc
            return
        
        # 检查模式是否发生变化（相对时间或过滤模式）
        mode_changed = (relative_mode != self.last_relative_time_mode or 
                       hide_goodcrc != self.last_filter_goodcrc_mode)
        
        # 检查是否需要完全重建（数据减少、过滤条件可能变化）
        existing_count = len(self.tree.get_children())
        if current_data_count < self.last_rendered_count or existing_count == 0 or mode_changed:
            need_full_rebuild = True
        
        # 预计算相对时间的基准（第一条数据的时间）
        base_seconds = None
        if relative_mode and self.data_list:
            base_seconds = self._parse_timestamp_to_seconds(self.data_list[0].timestamp)
        
        if need_full_rebuild:
            # 完全重建模式（导入、清空、过滤等情况）
            self._full_rebuild_treeview(hide_goodcrc, relative_mode, base_seconds)
        else:
            # 增量更新模式（正常采集数据）
            self._incremental_update_treeview(hide_goodcrc, relative_mode, base_seconds)
        
        self.last_rendered_count = current_data_count
        self.last_relative_time_mode = relative_mode
        self.last_filter_goodcrc_mode = hide_goodcrc
        
        # 根据是否有数据启用/禁用导出按钮
        try:
            if self.data_list:
                self.export_button.config(state=tk.NORMAL)
            else:
                self.export_button.config(state=tk.DISABLED)
        except Exception:
            pass
    
    def _full_rebuild_treeview(self, hide_goodcrc, relative_mode, base_seconds):
        """完全重建Treeview（用于导入、清空等操作）"""
        # 保存当前选中和滚动位置
        current_selection = self.tree.selection()
        selected_index = None
        selected_item_id = None
        
        try:
            if current_selection:
                selected_item_id = current_selection[0]
                selected_item_vals = self.tree.item(selected_item_id)
                try:
                    selected_index = int(selected_item_vals['values'][0]) - 1
                except Exception:
                    selected_index = None
        except Exception:
            selected_index = None
        
        try:
            prev_yview = self.tree.yview()
        except Exception:
            prev_yview = None
        
        # 清空现有项
        for child in self.tree.get_children():
            self.tree.delete(child)
        
        # 重新插入所有数据
        for item in self.data_list:
            if hide_goodcrc and isinstance(item.msg_type, str) and 'goodcrc' in item.msg_type.lower():
                continue
            
            self._insert_tree_item(item, relative_mode, base_seconds)
        
        # 恢复选中状态和滚动位置
        if selected_index is not None and 0 <= selected_index < len(self.data_list):
            target_child = None
            for child in self.tree.get_children():
                item_values = self.tree.item(child)['values']
                if item_values and int(item_values[0]) == selected_index + 1:
                    target_child = child
                    break
            
            if target_child is not None:
                try:
                    self.tree.selection_set(target_child)
                    self.tree.focus(target_child)
                except Exception:
                    pass
        
        # 恢复滚动位置
        try:
            if prev_yview and len(prev_yview) == 2:
                self.tree.yview_moveto(prev_yview[0])
        except Exception:
            pass
    
    def _incremental_update_treeview(self, hide_goodcrc, relative_mode, base_seconds):
        """增量更新Treeview（只添加新数据）"""
        # 找出需要添加的新数据
        existing_count = len(self.tree.get_children())
        
        # 计算应该有多少条数据（考虑过滤）
        if hide_goodcrc:
            # 需要计算已渲染的实际数据索引
            # 这里简化处理：如果有过滤，使用完全重建
            self._full_rebuild_treeview(hide_goodcrc, relative_mode, base_seconds)
            return
        
        # 只添加新增的数据项
        for i in range(existing_count, len(self.data_list)):
            item = self.data_list[i]
            self._insert_tree_item(item, relative_mode, base_seconds)
        
        # 自动滚动到底部（如果用户没有手动滚动）
        try:
            yview = self.tree.yview()
            # 如果滚动条接近底部（>0.9），自动滚动到最新数据
            if yview[1] > 0.9:
                children = self.tree.get_children()
                if children:
                    self.tree.see(children[-1])
        except Exception:
            pass
    
    def _insert_tree_item(self, item, relative_mode, base_seconds):
        """插入单个数据项到Treeview"""
        # 计算tag
        tag_name = None
        try:
            if isinstance(item.msg_type, str) and item.msg_type in MT:
                tag_name = item.msg_type
            elif isinstance(item.msg_type, str):
                lowered = item.msg_type.lower()
                for k in MT.keys():
                    if k.lower() == lowered:
                        tag_name = k
                        break
        except Exception:
            tag_name = None
        
        # 计算显示时间
        display_time = item.timestamp
        if relative_mode and base_seconds is not None:
            cur_sec = self._parse_timestamp_to_seconds(item.timestamp)
            if cur_sec is not None:
                dt = cur_sec - base_seconds
                if dt < 0:
                    dt = 0.0
                display_time = self._format_relative_time(dt)
        
        # 插入数据
        if tag_name:
            try:
                self.tree.insert('', tk.END, values=(
                    item.index,
                    display_time,
                    item.sop,
                    item.rev,
                    item.ppr,
                    item.pdr,
                    item.msg_type
                ), tags=(tag_name,))
            except Exception:
                self.tree.insert('', tk.END, values=(
                    item.index,
                    display_time,
                    item.sop,
                    item.rev,
                    item.ppr,
                    item.pdr,
                    item.msg_type
                ))
        else:
            self.tree.insert('', tk.END, values=(
                item.index,
                display_time,
                item.sop,
                item.rev,
                item.ppr,
                item.pdr,
                item.msg_type
            ))

    def _parse_timestamp_to_seconds(self, ts: Optional[str]) -> Optional[float]:
        """将 'HH:MM:SS.mmm' 或 'H:M:S.mmm' 样式的字符串解析为当天秒数（浮点）。
        如果解析失败，返回 None。
        """
        if not ts:
            return None
        try:
            # 尝试严格格式
            if len(ts) >= 12 and ts[2] == ':' and ts[5] == ':':
                dt = datetime.strptime(ts, "%H:%M:%S.%f")
            else:
                # 宽松一点：没有毫秒或位数不同
                try:
                    dt = datetime.strptime(ts, "%H:%M:%S")
                except Exception:
                    # 其他可能格式：MM:SS.mmm（无小时）或 SS.mmm（不太可能），尽量兜底
                    parts = ts.split(':')
                    if len(parts) == 2:
                        m = int(parts[0])
                        s = float(parts[1])
                        return m * 60 + s
                    elif len(parts) == 1:
                        return float(parts[0])
                    else:
                        return None
            return dt.hour * 3600 + dt.minute * 60 + dt.second + dt.microsecond / 1e6
        except Exception:
            return None

    def _format_relative_time(self, seconds: float) -> str:
        """将秒数格式化为 'MM:SS.mmm'（分钟:秒.毫秒），分钟可超过两位。"""
        try:
            if seconds < 0:
                seconds = 0.0
            total_ms = int(round(seconds * 1000))
            mins, ms_rem = divmod(total_ms, 60_000)
            secs, ms = divmod(ms_rem, 1000)
            return f"{mins:02d}:{secs:02d}.{ms:03d}"
        except Exception:
            # 兜底：直接返回秒数字符串
            return f"{seconds:.3f}"
    
    def on_item_select(self, event):
        """处理列表项选择事件"""
        selection = self.tree.selection()
        if selection:
            item = self.tree.item(selection[0])
            index = int(item['values'][0]) - 1  # 转换为0基索引
            
            if 0 <= index < len(self.data_list):
                self.current_selection = self.data_list[index]
                self.display_data(self.current_selection)
                # 锁定竖线（列表选中视为明确操作）
                self._is_vline_locked = True
                self._hover_last_item_index = index
                # 更新曲线中的选中竖线
                self._update_selection_vline(self.current_selection)
    
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

            # 当设备断开、非导入模式且开发者模式开启时：选中报文后自动将竖线置于焦点
            try:
                if (not self.device_open) and (not self.import_mode) and getattr(self, '_egg_activated', False):
                    # 使用当前选择项计算x并聚焦到2秒窗口（仅点击时触发一次）
                    if self.current_selection is not None:
                        x = self._get_item_plot_x(self.current_selection)
                        if x is not None:
                            self._focus_on_time_x(x, window_seconds=4)
            except Exception:
                pass

    def display_data(self, item: DataItem):
        """在右侧显示选中的数据"""
        # 临时启用以写入，然后恢复为只读
        try:
            self.data_text.config(state=tk.NORMAL)
            self.data_text.delete(1.0, tk.END)

            # 基本信息
            self.data_text.insert(tk.END, "==== 基本信息 ====\n", 'cn')
            self.data_text.insert(tk.END, "Index: ", ('black', 'bold'))
            self.data_text.insert(tk.END, f"{item.index}\n", 'gray')
            self.data_text.insert(tk.END, "Time: ", ('black', 'bold'))
            self.data_text.insert(tk.END, f"{item.timestamp}\n", 'gray')
            self.data_text.insert(tk.END, "SOP: ", ('black', 'bold'))
            self.data_text.insert(tk.END, f"{item.sop}\n", 'gray')
            self.data_text.insert(tk.END, "Rev: ", ('black', 'bold'))
            self.data_text.insert(tk.END, f"{item.rev}\n", 'gray')
            self.data_text.insert(tk.END, "PPR: ", ('black', 'bold'))
            self.data_text.insert(tk.END, f"{item.ppr}\n", 'gray')
            self.data_text.insert(tk.END, "PDR: ", ('black', 'bold'))
            self.data_text.insert(tk.END, f"{item.pdr}\n", 'gray')
            self.data_text.insert(tk.END, "Message Type: ", ('black', 'bold'))
            self.data_text.insert(tk.END, f"{item.msg_type}\n", 'gray')
            self.data_text.insert(tk.END, "\n==== 详细数据 ====\n", 'cn')
            self.data_text.insert(tk.END, f"Raw: 0x{int(item.data.raw(), 2):0{int(len(item.data.raw())/4)+(1 if len(item.data.raw())%4!=0 else 0)}X}\n", 'green')
            lst = []
            for i in item.data.value():
                renderer(i, 0, lst)
            for line in lst:
                self.data_text.insert(tk.END, line[0], line[1])
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
            self.set_status("正在导出数据...", level='busy')
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['序号', '时间', 'SOP', 'Rev', 'PPR', 'PDR', 'Msg Type', '详细数据', 'Raw'])
                for item in self.data_list:
                    # 使用 format_data 输出的人类可读文本作为详细数据字段
                    data_text = self.format_data(item.data)
                    writer.writerow([item.index,
                                     item.timestamp,
                                     item.sop,
                                     item.rev,
                                     item.ppr,
                                     item.pdr,
                                     item.msg_type,
                                     data_text,
                                     f"{int(item.data.raw(), 2):0{int(len(item.data.raw())/4)+(1 if len(item.data.raw())%4!=0 else 0)}X}"])

            self.set_status(f"已导出 {len(self.data_list)} 条数据 到 {file_path}", level='ok')
        except Exception as e:
            messagebox.showerror("导出失败", f"导出 CSV 失败:\n{e}")

    def import_csv(self):
        """从CSV导入数据，仅解析两列：时间、Raw（全大写HEX）。
        Raw 将被转换为长度为64字节的uint8列表（不足末尾补0，超出则截断），
        然后使用 WITRN_DEV.auto_unpack(data) 解析并加入列表。
        """
        # 若正在进行数据采集（未暂停），阻止导入
        if not self.is_paused:
            messagebox.showwarning("操作受限", "正在收集数据，无法导入CSV。请先暂停并清空列表后再试。")
            return
        file_path = filedialog.askopenfilename(
            filetypes=[('CSV 文件', '*.csv')],
            title='选择CSV文件'
        )
        if not file_path:
            return

        # 导入应覆盖当前列表
        try:
            if len(self.data_list) > 0:
                if not messagebox.askyesno("确认", "导入将清空当前数据列表，是否继续？"):
                    return
                self.clear_list(ask_user=False)
        except Exception:
            pass

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
                            if self.parser is None:
                                # 若仍无k2，跳过解析
                                failed += 1
                                continue
                            _, pkg = self.parser.auto_unpack(data_bytes, last_pdo, last_ext, last_rdo)
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
                                    rev = pkg["Message Header"][4].value()[4:]
                                    if rev == 'rved':
                                        rev = None
                                except Exception:
                                    rev = None
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
                                self.add_data_item(sop, rev, ppr, pdr, msg_type, pkg, force=True, timestamp=t_str)
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
            # 进入导入模式
            self.import_mode = True

            if self.device_open:
                self.set_status(f"导入完成：成功 {success} 条，失败 {failed} 条。可开始收集（会先清空）。", level='ok')
            else:
                self.set_status(f"导入完成：成功 {success} 条，失败 {failed} 条。设备未连接，无法开始收集；请先连接设备。", level='warn')
        except Exception as e:
            messagebox.showerror("导入失败", f"无法导入CSV:\n{e}")
        
    def connect_device(self):
        """连接/断开设备（连接/断开仅由采集进程处理，主进程不直接 open/close 设备）。"""
        try:
            if not self.device_open:
                # 准备连接
                if len(self.data_list) > 0:
                    if not messagebox.askyesno("确认", "连接设备将清空当前数据列表，是否继续？"):
                        return
                    self.clear_list(ask_user=False)
                # 标记为“已连接（待开始）”，实际打开由采集进程完成
                self.device_open = True
                self.is_paused = True  # 初始暂停，等待用户点击“开始”
                self.set_status("正在连接设备...", level='warn')
                self.awaiting_connection_ack = True
                self.autostart_after_connect = False
                # 启动采集进程（由其负责 open(device_path)）
                self.start_data_thread_if_needed()
                # 更新按钮与菜单
                self.pause_button.config(state=tk.NORMAL, text="开始")
                self.connect_button.config(text="断开连接")
                # 连接后初始化并启动曲线刷新（仅当彩蛋激活后才有曲线区）
                try:
                    if self._egg_activated:
                        self._ensure_plot_initialized()
                        self._reset_plot()
                        self._start_plot_updates()
                except Exception:
                    pass
            else:
                # 断开连接：仅停止采集进程与本地状态，设备由子进程 close()
                self._stop_collection_process()
                self.last_pdo = None
                self.last_rdo = None
                self.set_quick_pdo_rdo(None, None, True)
                self.device_open = False
                self.is_paused = True
                self.awaiting_connection_ack = False
                self.autostart_after_connect = False
                self.set_status("设备已断开", level='warn')
                self.pause_button.config(state=tk.DISABLED, text="开始")
                self.connect_button.config(text="连接设备")
                # 重置基本信息面板（同步调用，保证立即刷新）
                try:
                    self.reset_iv_info()
                except Exception:
                    pass
                # 断开后停止曲线刷新
                try:
                    self._stop_plot_updates()
                except Exception:
                    pass
                time.sleep(0.1)
                # 断开后无需重置解析器实例；设备连接实例在子进程中，会随子进程结束而释放。
        except Exception as e:
            try:
                messagebox.showerror("连接失败", f"无法连接到设备：{e}")
            except Exception:
                pass
            self.set_status("连接设备失败", level='error')
    
    def pause_collection(self):
        """暂停/恢复数据收集"""
        if self.is_paused:
            # 若仍在等待连接确认，则只标记连接后自动开始，避免错误显示“数据收集中...”
            if self.awaiting_connection_ack:
                self.autostart_after_connect = True
                # 保持状态为“正在连接设备...”，按钮文案先行更新为“暂停”以告诉用户将自动开始
                try:
                    self.pause_button.config(text="暂停")
                except Exception:
                    pass
                return
            # 恢复收集
            if self.import_mode:
                if len(self.data_list) > 0:
                    if not messagebox.askyesno("确认", "开始收集将清空当前数据列表，是否继续？"):
                        return
                    self.clear_list(ask_user=False)
                # 退出导入模式
                self.import_mode = False
            self.is_paused = False
            # 通知采集进程恢复
            if self.pause_flag is not None:
                self.pause_flag.value = 0
            self.pause_button.config(text="暂停")
            self.set_status("数据收集中...", level='busy')
        else:
            # 暂停收集
            self.is_paused = True
            # 通知采集进程暂停
            if self.pause_flag is not None:
                self.pause_flag.value = 1
            self.pause_button.config(text="开始")
            self.set_status("数据收集已暂停", level='warn')

    def clear_list(self, ask_user: bool = True):
        """清空数据列表，并重置导入模式。"""
        if (not ask_user) or messagebox.askyesno("确认", "确定要清空所有数据吗？"):
            self.data_list.clear()
            self.current_selection = None
            # 解锁竖线并清除（固定线和预览线都清除）
            try:
                self._is_vline_locked = False
                self._hover_last_item_index = None
                self._update_selection_vline(None)
                self._update_hover_preview_vline(None)
            except Exception:
                pass
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
            self.last_pdo = None
            self.last_rdo = None
            self.set_quick_pdo_rdo(None, None, True)
            # 退出导入模式
            self.import_mode = False
            if self.is_paused and not self.device_open:
                self.set_status("列表已清空", level='info')
            elif self.device_open and self.is_paused:
                self.set_status("列表已清空，设备已连接", level='ok')
            elif self.device_open and not self.is_paused:
                self.set_status("列表已清空，数据收集中...", level='busy')

    def refresh_data_loop(self):
        """数据刷新循环（在后台线程中运行）
        智能刷新机制：根据数据量动态调整刷新频率
        """
        last_data_count = 0
        while True:
            try:
                current_data_count = len(self.data_list)
                
                # 只有在数据发生变化时才考虑更新
                if current_data_count != last_data_count:
                    # 根据数据量计算刷新间隔（动态调整）
                    if current_data_count < 1000:
                        min_interval = 0.1    # <1k: 100ms (10Hz)
                    elif current_data_count < 5000:
                        min_interval = 0.3    # 1k-5k: 300ms (~3Hz)
                    elif current_data_count < 10000:
                        min_interval = 0.5    # 5k-10k: 500ms (2Hz)
                    elif current_data_count < 20000:
                        min_interval = 1.0    # 10k-20k: 1000ms (1Hz)
                    else:
                        min_interval = 2.0    # >20k: 2000ms (0.5Hz)
                    
                    now = time.time()
                    time_since_last_update = now - self.last_treeview_update_time
                    
                    # 如果距离上次刷新已经超过最小间隔，立即刷新
                    if time_since_last_update >= min_interval:
                        self.root.after(0, self._safe_update_treeview)
                        self.last_treeview_update_time = now
                        last_data_count = current_data_count
                    else:
                        # 否则标记为待刷新，等待下次检查
                        if not self.pending_treeview_update:
                            self.pending_treeview_update = True
                            # 计算还需要等待多久
                            delay_ms = int((min_interval - time_since_last_update) * 1000)
                            self.root.after(delay_ms, self._delayed_update_treeview)
                
                # 检查频率也根据数据量调整
                if current_data_count < 5000:
                    time.sleep(0.1)   # <5k: 快速检查
                elif current_data_count < 20000:
                    time.sleep(0.2)   # 5k-20k: 中速检查
                else:
                    time.sleep(0.5)   # >20k: 慢速检查
                    
            except Exception as e:
                print(f"刷新数据时出错: {e}")
                time.sleep(1)
    
    def _safe_update_treeview(self):
        """安全地更新Treeview（在主线程中调用）"""
        try:
            self.update_treeview()
        except Exception as e:
            print(f"更新Treeview时出错: {e}")
    
    def _delayed_update_treeview(self):
        """延迟更新Treeview（用于批量更新）"""
        try:
            if self.pending_treeview_update:
                self.pending_treeview_update = False
                self.last_treeview_update_time = time.time()
                self.update_treeview()
        except Exception as e:
            print(f"延迟更新Treeview时出错: {e}")
            

    def run(self):
        """运行GUI应用程序"""
        try:
            self.root.mainloop()
        finally:
            # 程序退出时先停止消费线程
            self.queue_consumer_running = False
            
            # 等待消费线程结束
            if self.queue_consumer_thread is not None and self.queue_consumer_thread.is_alive():
                self.queue_consumer_thread.join(timeout=1.0)
            
            # 再停止采集进程
            self._stop_collection_process()
            
            # 最后清理队列
            self.data_queue = None
            self.iv_queue = None
    
if __name__ == "__main__":
    # Windows上必须设置为spawn模式以避免freeze_support问题
    try:
        multiprocessing.set_start_method('spawn')
    except RuntimeError:
        # 已经设置过了
        pass
    
    app = WITRNGUI()
    app.run()
"""
python -m nuitka witrn_pd_sniffer.py ^
--standalone ^
--onefile ^
--windows-console-mode=disable ^
--enable-plugin=tk-inter ^
--windows-icon-from-ico=brain.ico ^
--product-name="WITRN PD Sniffer" ^
--product-version=3.7.3.0 ^
--copyright="JohnScotttt" ^
--output-dir=output ^
--output-filename=witrn_pd_sniffer_v3.7.3.exe
"""