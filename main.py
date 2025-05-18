#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
电脑自动化控制工具 - 主程序
用于自动化操作Windows应用程序，特别是针对"数据管理部工具台"的自动化操作。
"""

import os
import sys
import time
import csv
import re
import pyautogui
import pyperclip
import pygetwindow as gw
from pywinauto import Application, Desktop

# 导入自定义模块
from utils.sku_processor import process_single_sku, split_sku


def main():
    """主程序入口"""
    print("开始执行自动化控制程序...")
    
    # 连接到目标应用
    window_title = "数据管理部工具台"
    try:
        app = Application(backend="uia").connect(best_match=window_title)
        windows = gw.getWindowsWithTitle(window_title)
        if not windows:
            print(f"错误：找不到标题为 '{window_title}' 的窗口")
            return
            
        window = windows[0]
        window.activate()
        time.sleep(1)  # 等待窗口激活
        window = app.window(best_match=window_title)
        print(f"已成功连接到 '{window_title}' 窗口")
    except Exception as e:
        print(f"连接应用程序时出错: {e}")
        return
    
    # 导航到广告后台
    try:
        hyperlink = window.child_window(title="广告后台", control_type="Hyperlink")
        hyperlink.click_input()
        print("广告后台菜单项已点击")
        window.wait('ready', timeout=3)
    except Exception as e:
        print(f"点击广告后台时出错: {e}")
        return
    
    # 点击销售人员登录通道
    try:
        sales_login_hyperlink = window.child_window(
            title="销售人员登录通道", control_type="Hyperlink")
        sales_login_hyperlink.click_input()
        print("销售人员登录通道已点击")
        window.wait('ready', timeout=3)
    except Exception as e:
        print(f"点击销售人员登录通道时出错: {e}")
        return
    
    # 登录操作
    try:
        flower_name_edit = window.child_window(title="花名:", control_type="Edit")
        if flower_name_edit.exists(timeout=2):
            flower_name_edit.click_input()
            flower_name_edit.type_keys("Cloris", with_spaces=True)

            password_edit = window.child_window(title="密码:", control_type="Edit")
            password_edit.click_input()
            password_edit.type_keys("ChrdwHdhxt6688", with_spaces=True)

            submit_button = window.child_window(title="提交", control_type="Button")
            submit_button.click_input()
            print("已输入花名、密码并提交")
        else:
            print("没有找到花名输入框，跳过此步骤")
    except Exception as e:
        print(f"登录操作时出错: {e}")
    
    window.wait('ready', timeout=3)
    
    # 导航到工具栏
    all_controls = window.descendants()
    toolbar_hyperlink_controls = [ctrl for ctrl in all_controls if re.search(
        r"工具栏", ctrl.window_text()) and ctrl.friendly_class_name() == "Hyperlink"]
    
    if not toolbar_hyperlink_controls:
        print("未找到工具栏控件")
        return
    
    # 读取CSV文件
    csv_file_path = "需下载牛牛数据4.csv"
    if not os.path.exists(csv_file_path):
        print(f"错误：找不到文件 {csv_file_path}")
        return
    
    with open(csv_file_path, 'r', encoding='utf-8') as file:
        csv_reader = csv.DictReader(file)
        sku_list = [row['系统SKU'] for row in csv_reader if '系统SKU' in row]
    
    if not sku_list:
        print("错误：CSV文件中没有找到系统SKU数据")
        return
    
    # 初始化计数器
    total_skus = len(sku_list)
    successful_queries = 0
    failed_queries = 0
    successful_downloads = 0
    failed_downloads = 0
    
    # 点击工具栏并导航到共享关键词
    toolbar_hyperlink_controls[0].click_input()
    window.wait('visible', timeout=5)
    popup_controls = window.descendants()
    keyword_controls = [
        ctrl for ctrl in popup_controls
        if re.search(r"共享关键词", ctrl.window_text())
    ]
    
    if not keyword_controls:
        print("未找到'共享关键词'控件")
        return
    
    # 点击共享关键词
    keyword_controls[0].click_input()
    print("成功点击'共享关键词'")
    window.wait('ready', timeout=3)
    
    # 定位输入框
    input_box = window.child_window(
        title="请输入erpsku、sellersku、asin、关键词或站点名进行搜索", control_type="Edit")
    
    if not input_box.exists():
        print("未找到输入框控件")
        return
    
    # 处理每个SKU
    for sku in sku_list:
        print(f"\n开始处理SKU: {sku}")
        
        # 将当前SKU复制到剪贴板
        pyperclip.copy(sku)
        
        # 清空输入框并粘贴SKU
        input_box.set_edit_text("")
        input_box.click_input()
        input_box.type_keys('^v')
        time.sleep(1)
        print(f"成功粘贴SKU: {sku}")
        
        window.wait('ready', timeout=3)
        
        # 点击查询按钮
        all_buttons = window.descendants(control_type="Button")
        all_buttons[5].click_input()
        window.wait('ready', timeout=3000)
        
        # 点击下载链接
        download_link = window.child_window(
            title="只下载自己站点数据", control_type="Hyperlink")
        
        if not download_link.exists():
            print(f"未找到'只下载自己站点数据'超链接，SKU: {sku}")
            failed_downloads += 1
            continue
        
        download_link.click_input()
        print(f"成功点击'只下载自己站点数据'超链接，SKU: {sku}")
        
        try:
            # 处理保存文件对话框
            save_dialog = Desktop(backend="win32").window(
                title_re="保存文件|Save As|Save File", top_level_only=False, enabled_only=True)
            save_dialog.wait('visible', timeout=30000)
            print(f"保存文件对话框已出现，SKU: {sku}")
            save_dialog.type_keys("{ENTER}")
            print(f"已按下回车键保存文件，SKU: {sku}")
            time.sleep(5)
            
            # 处理消息提示对话框
            try:
                message_dialog = Desktop(backend="win32").window(
                    title="消息提示", top_level_only=False, enabled_only=True)
                message_dialog.wait('visible', timeout=300000)
                print("找到'消息提示'对话框")
                message_dialog.set_focus()
                message_dialog.type_keys("{ENTER}")
                message_dialog.type_keys("{ENTER}")
                print("已在'消息提示'对话框上按下回车键")
            except Exception as e:
                print(f"处理'消息提示'对话框时出错: {e}")
            
            successful_downloads += 1
            successful_queries += 1
        except Exception as e:
            print(f"处理保存文件对话框时出错: {e}")
            failed_downloads += 1
            failed_queries += 1
        
        # 等待一段时间，以便下一次查询
        time.sleep(5)
    
    # 打印统计结果
    print("\n统计结果:")
    print(f"总SKU数: {total_skus}")
    print(f"成功查询次数: {successful_queries}")
    print(f"失败查询次数: {failed_queries}")
    print(f"成功下载次数: {successful_downloads}")
    print(f"失败下载次数: {failed_downloads}")


if __name__ == "__main__":
    main()