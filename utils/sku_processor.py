#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
SKU处理模块
提供SKU处理相关的功能函数
"""

import time
import pyperclip
from pywinauto import Desktop


def split_sku(sku, parts=3):
    """
    将单个SKU拆分为指定份数
    
    Args:
        sku (str): 要拆分的SKU字符串
        parts (int): 拆分的份数，默认为3
        
    Returns:
        list: 拆分后的SKU片段列表
    """
    length = len(sku)
    return [sku[i*length // parts: (i+1)*length // parts]
            for i in range(parts)]


def navigate_to_start():
    """
    重新导航到起始页面
    
    Returns:
        bool: 是否成功导航
    """
    try:
        # 这里可以实现重新导航到起始页面的逻辑
        # 例如点击某个导航按钮或者刷新页面等
        return True
    except Exception as e:
        print(f"重新导航时出错: {e}")
        return False


def process_single_sku(sku, window, input_box, max_retries=3):
    """
    处理单个SKU，包括重试和拆分逻辑
    
    Args:
        sku (str): 要处理的SKU字符串
        window: 窗口对象
        input_box: 输入框控件对象
        max_retries (int): 最大重试次数，默认为3
        
    Returns:
        bool: 是否成功处理
    """
    retry_count = 0
    while retry_count < max_retries:
        try:
            # 粘贴SKU到输入框
            pyperclip.copy(sku)
            if input_box.exists():
                input_box.set_edit_text("")
                input_box.click_input()
                input_box.type_keys('^v')
                time.sleep(1)
                print(f"成功粘贴SKU: {sku}")
            else:
                print("未找到输入框控件")
                return False

            window.wait('ready', timeout=3)

            # 点击查询按钮
            all_buttons = window.descendants(control_type="Button")
            all_buttons[5].click_input()
            window.wait('ready', timeout=3000)

            # 点击下载链接
            download_link = window.child_window(
                title="只下载自己站点数据", control_type="Hyperlink")
            if download_link.exists():
                download_link.click_input()
                print(f"成功点击'只下载自己站点数据'超链接，SKU: {sku}")
            else:
                print(f"未找到'只下载自己站点数据'超链接，SKU: {sku}")
                return False

            # 处理保存文件对话框
            save_dialog = Desktop(backend="win32").window(
                title_re="保存文件|Save As|Save File", top_level_only=False, enabled_only=True)
            save_dialog.wait('visible', timeout=3000)
            save_dialog.type_keys("{ENTER}")
            time.sleep(5)

            # 处理消息提示对话框
            message_dialog = Desktop(backend="win32").window(
                title="消息提示", top_level_only=False, enabled_only=True)
            message_dialog.wait('visible', timeout=300000)
            message_dialog.set_focus()
            message_dialog.type_keys("{ENTER}")
            message_dialog.type_keys("{ENTER}")

            return True  # 成功完成

        except Exception as e:
            print(f"处理SKU: {sku}时出现错误: {e}")
            retry_count += 1
            if retry_count < max_retries:
                print(f"正在进行第{retry_count}次重试...")
                if navigate_to_start():
                    print("成功重新导航到起始页面")
                else:
                    print("重新导航失败，尝试下一次重试")
            else:
                print(f"SKU: {sku}处理失败，已达到最大重试次数")
                return False

        time.sleep(5)  # 等待一段时间，以便下一次查询
