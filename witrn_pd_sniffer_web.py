#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WITRN HID PD查看器 Web 应用程序 (基于pywebview)
"""

import webview
import threading
import time
import csv
import json
import os
from datetime import datetime
from typing import List, Dict, Any, Optional
from witrnhid import WITRN_DEV, metadata


class WITRNWebBackend:
    """WITRN HID 数据查看器后端类"""
    
    def __init__(self):
        self.data_list: List[Dict[str, Any]] = []
        self.current_selection: Optional[Dict[str, Any]] = None
        
        # 控制状态
        self.is_paused = False
        self.data_collection_active = False
        
        # 设备句柄和数据采集线程控制
        self.k2 = None
        self.data_thread_started = False
        self.data_thread = None
        
        # 初始化设备连接
        self.init_device()
    
    def init_device(self):
        """初始化设备连接"""
        try:
            self.k2 = WITRN_DEV()
            print("设备已连接")
            self.start_data_thread_if_needed()
        except Exception as e:
            print(f"无法连接到K2: {e}")
            self.k2 = None
    
    def start_data_thread_if_needed(self):
        """在设备可用时安全地启动数据采集线程一次"""
        if self.data_thread_started or self.k2 is None:
            return
        
        # 启动收集线程
        self.data_thread = threading.Thread(target=self._collect_data_loop, daemon=True)
        self.data_thread.start()
        self.data_thread_started = True
    
    def _collect_data_loop(self):
        """内部数据收集循环"""
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
                    data = self.format_data_for_web(pkg)
                    
                    # 添加数据项
                    self.add_data_item(sop, ppr, pdr, msg_type, data)
                    
            except Exception as e:
                # 仅当是真正的 read error 时才视为设备断开
                err_text = str(e).lower()
                if 'read error' in err_text or 'readerror' in err_text:
                    print(f"数据采集异常（断开）: {e}")
                    try:
                        self.k2 = None
                        self.data_collection_active = False
                    except Exception:
                        pass
                    
                    # 通知前端设备断开，更新UI状态
                    try:
                        disconnect_data = {
                            'type': 'device_disconnected',
                            'message': '检测到设备已断开，请重连或检查连接。'
                        }
                        webview.windows[0].evaluate_js(f'window.onDeviceDisconnected({json.dumps(disconnect_data, ensure_ascii=False)});')
                    except Exception as notify_error:
                        print(f"通知前端设备断开时出错: {notify_error}")
                    
                    time.sleep(0.5)
                else:
                    # 非 read error：记录为警告并继续
                    print(f"数据采集警告（非断开）: {e}")
                    time.sleep(0.1)
    
    def format_data_for_web(self, data: metadata) -> Dict[str, Any]:
        """将metadata对象格式化为Web可用的字典格式"""
        result = {}
        try:
            for value1 in data.value():
                if not isinstance(value1.value(), list):
                    result[value1.field()] = value1.value()
                else:
                    result[value1.field()] = {}
                    for value2 in value1.value():
                        if not isinstance(value2.value(), list):
                            result[value1.field()][value2.field()] = value2.value()
                        else:
                            result[value1.field()][value2.field()] = {}
                            for value3 in value2.value():
                                if not isinstance(value3.value(), list):
                                    result[value1.field()][value2.field()][value3.field()] = value3.value()
        except Exception as e:
            print(f"格式化数据时出错: {e}")
            result = {"error": str(e)}
        
        return result
    
    def add_data_item(self, sop: str, ppr: str, pdr: str, msg_type: str, data: Dict[str, Any] = None, force: bool = False):
        """添加新的数据项到列表"""
        # 只有在数据收集激活且未暂停时才添加数据，除非强制添加
        if not force and (not self.data_collection_active or self.is_paused):
            return
            
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        index = len(self.data_list) + 1
        
        item = {
            'index': index,
            'timestamp': timestamp,
            'sop': sop,
            'ppr': ppr,
            'pdr': pdr,
            'msg_type': msg_type,
            'data': data or {}
        }
        
        self.data_list.append(item)
        
        # 通知前端新数据到达
        try:
            webview.windows[0].evaluate_js(f'window.onDataReceived({json.dumps(item, ensure_ascii=False)});')
        except Exception as e:
            print(f"通知前端数据时出错: {e}")
    
    # Web API 方法
    def get_status(self):
        """获取当前状态"""
        return {
            'device_connected': self.k2 is not None,
            'data_count': len(self.data_list),
            'is_paused': self.is_paused,
            'data_collection_active': self.data_collection_active
        }
    
    def start_collection(self):
        """开始数据收集"""
        self.data_collection_active = True
        self.is_paused = False
        return {'success': True, 'message': '数据收集已开始'}
    
    def pause_collection(self, paused: bool):
        """暂停/恢复数据收集"""
        self.is_paused = paused
        return {'success': True, 'message': '暂停状态已更新'}
    
    def reconnect_device(self):
        """重新连接设备"""
        try:
            self.k2 = WITRN_DEV()
            self.start_data_thread_if_needed()
            return {'success': True, 'message': '设备已连接'}
        except Exception as e:
            return {'success': False, 'error': str(e)}
    
    def export_list(self):
        """导出数据列表为CSV文件"""
        if not self.data_list:
            return {'success': False, 'error': '没有数据可导出'}
        
        try:
            # 使用文件对话框选择保存位置
            file_path = webview.windows[0].create_file_dialog(
                webview.SAVE_DIALOG,
                directory='',
                save_filename='witrn_pd_data.csv',
                file_types=['CSV Files (*.csv)']
            )
            
            if not file_path:
                return {'success': False, 'error': '用户取消了保存'}
            
            # 处理可能的元组返回值
            if isinstance(file_path, tuple) and len(file_path) > 0:
                file_path = file_path[0]
            elif isinstance(file_path, tuple):
                return {'success': False, 'error': '文件路径无效'}
            
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['序号', '时间', 'SOP', 'PPR', 'PDR', 'Msg Type', '详细数据'])
                for item in self.data_list:
                    data_text = json.dumps(item['data'], ensure_ascii=False, indent=2)
                    writer.writerow([
                        item['index'], 
                        item['timestamp'], 
                        item['sop'], 
                        item['ppr'], 
                        item['pdr'], 
                        item['msg_type'], 
                        data_text
                    ])
            
            return {'success': True, 'message': f'已导出 {len(self.data_list)} 条数据'}
            
        except Exception as e:
            return {'success': False, 'error': str(e)}
    
    def clear_list(self):
        """清空数据列表"""
        self.data_list.clear()
        self.current_selection = None
        return {'success': True, 'message': '列表已清空'}
    
    def get_data_list(self):
        """获取当前数据列表"""
        return self.data_list.copy()
    
    def set_data_list(self, data_list: List[Dict[str, Any]]):
        """设置数据列表"""
        self.data_list = data_list.copy()
        return {'success': True, 'message': '数据列表已更新'}


def create_webview_window():
    """创建webview窗口"""
    # 获取HTML文件路径
    html_path = os.path.join(os.path.dirname(__file__), 'index.html')
    
    # 创建后端实例
    backend = WITRNWebBackend()
    
    # 创建webview窗口
    webview.create_window(
        title='WITRN PD解析v2.0',
        url=html_path,
        width=1300,
        height=800,
        resizable=True,
        js_api=backend
    )
    
    return backend


if __name__ == "__main__":
    # 创建webview应用
    backend = create_webview_window()
    
    # 启动webview
    webview.start(debug=False)
