# pylint: disable=all
import pyautogui
import time
import csv
import re
from pywinauto import Application
import subprocess
import pyperclip
from PIL import Image, ImageEnhance, ImageFilter
import pytesseract
import pygetwindow as gw
import easyocr
import numpy as np
import ocrmypdf
import tempfile
import os
import sys
import io
from pywinauto import Application
import math


# print('打开牛牛')
# subprocess.Popen('C:\\Program Files (x86)\\yibai_saler_tools\\yibai_saler_tools.exe')
# # app = Application(backend='uia').start(r'C:\Program Files (x86)\yibai_saler_tools\yibai_saler_tools.exe')
# time.sleep(10)

# # 激活牛牛窗口张梦MYXJ001
# print('激活窗口', pyautogui.size())
# time.sleep(5)
# # 自动输入用户名和密码
# # 尝试定位用户名输入框
# print('输入用户名')

# # 使用固定坐标代替图像识别
# # 计算相对位置
# screen_width, screen_height = pyautogui.size()
# # relative_x = int(screen_width * 0.293)  # 220 / 752 ≈ 0.286
# # relative_y = int(screen_height * 0.440)  # 373 / 847 ≈ 0.423
# relative_x = int(screen_width * 0.475)  # 685 / 1440 ≈ 0.475
# relative_y = int(screen_height * 0.425)  # 383 / 900 ≈ 0.425
# print(relative_x, relative_y)
# pyautogui.click(relative_x, relative_y)
# time.sleep(1)
# username = '张梦MYXJ001'
# pyperclip.copy(username)  # 复制用户名到剪贴板
# pyautogui.hotkey('ctrl', 'v')  # 使用热键Ctrl+V粘贴
# time.sleep(2)
# pyautogui.click(relative_x, relative_y+130)
# password = 'ChrdwHdhxt6688'
# pyperclip.copy(password)  # 复制用户名到剪贴板
# pyautogui.hotkey('ctrl', 'v')  # 使用热键Ctrl+V粘贴
# pyautogui.press('enter')  # 登录


from pywinauto import Desktop


def split_sku(sku, parts=3):
    """将单个 SKU 拆分为指定份数"""
    length = len(sku)
    return [sku[i*length // parts: (i+1)*length // parts]
            for i in range(parts)]


def process_single_sku(sku, max_retries=3):
    """处理单个 SKU，包括重试和拆分逻辑"""
    retry_count = 0
    while retry_count < max_retries:
        try:
            # 粘贴 SKU 到输入框
            pyperclip.copy(sku)
            if input_box.exists():
                input_box.set_edit_text("")
                input_box.click_input()
                input_box.type_keys('^v')
                time.sleep(1)
                print(f"成功粘贴 SKU: {sku}")
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
                print(f"成功点击 '只下载自己站点数据' 超链接，SKU: {sku}")
            else:
                print(f"未找到 '只下载自己站点数据' 超链接，SKU: {sku}")
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
            print(f"处理 SKU: {sku} 时出现错误: {e}")
            retry_count += 1
            if retry_count < max_retries:
                print(f"正在进行第 {retry_count} 次重试...")
                if navigate_to_start():
                    print("成功重新导航到起始页面")
                else:
                    print("重新导航失败，尝试下一次重试")
            else:
                print(f"SKU: {sku} 处理失败，已达到最大重试次数")
                return False

        time.sleep(5)  # 等待一段时间，以便下一次查询

# # 主程序
# total_skus = len(sku_list)
# successful_queries = 0
# failed_queries = 0
# successful_downloads = 0
# failed_downloads = 0

# for sku in sku_list:
#     if process_single_sku(sku):
#         successful_downloads += 1
#         successful_queries += 1
#     else:
#         print(f"SKU: {sku} 处理失败，尝试拆分处理")
#         split_skus = split_sku(sku)
#         split_success = True
#         for sub_sku in split_skus:
#             if process_single_sku(sub_sku):
#                 successful_downloads += 1
#                 successful_queries += 1
#             else:
#                 split_success = False
#                 failed_downloads += 1
#                 failed_queries += 1
#                 print(f"拆分后的 SKU: {sub_sku} 处理失败")

#         if split_success:
#             print(f"SKU: {sku} 拆分处理成功")
#         else:
#             print(f"SKU: {sku} 拆分处理部分失败")

# print("\n统计结果:")
# print(f"总SKU数: {total_skus}")
# print(f"成功查询次数: {successful_queries}")
# print(f"失败查询次数: {failed_queries}")
# print(f"成功下载次数: {successful_downloads}")
# print(f"失败下载次数: {failed_downloads}")

# 列出所有顶层窗口
windows = Desktop(backend="uia").windows()

# 打印所有顶层窗口的标题
for w in windows:
    print(w.window_text())

# 使用部分标题匹配窗口
window_title = "数据管理部工具台"
# window_title="yibai_saler_tools"

app = Application(backend="uia").connect(best_match=window_title)
window = gw.getWindowsWithTitle(window_title)
print(window)
window = window[0]
window.activate()
time.sleep(1)  # 等待窗口激活
# 获取窗口
window = app.window(best_match=window_title)

# 打印控件标识符
# window.print_control_identifiers()


# 获取广告后台的 Hyperlink 控件
hyperlink = window.child_window(title="广告后台", control_type="Hyperlink")

# 点击广告后台的 Hyperlink
hyperlink.click_input()
print('hyperlink', hyperlink)
# print(methods)
print("广告后台菜单项已点击")
window.wait('ready', timeout=3)
# 定位到 "销售人员登录通道" 的 Hyperlink 控件
sales_login_hyperlink = window.child_window(
    title="销售人员登录通道", control_type="Hyperlink")

# 点击该控件
sales_login_hyperlink.click_input()

print("销售人员登录通道已点击")
window.wait('ready', timeout=3)
try:
    # 定位 "花名" 的输入框并输入内容
    flower_name_edit = window.child_window(title="花名:", control_type="Edit")
    if flower_name_edit.exists(timeout=2):
        # 定位 "花名" 的输入框并输入内容
        flower_name_edit = window.child_window(
            title="花名:", control_type="Edit")
        flower_name_edit.click_input()
        flower_name_edit.type_keys("Cloris", with_spaces=True)

        # 定位 "密码" 的输入框并输入内容
        password_edit = window.child_window(title="密码:", control_type="Edit")
        password_edit.click_input()
        password_edit.type_keys("ChrdwHdhxt6688", with_spaces=True)

        # 定位 "提交" 按钮并点击
        submit_button = window.child_window(title="提交", control_type="Button")
        submit_button.click_input()

        print("已输入花名、密码并提交")
    else:
        print("没有找到花名输入框，跳过此步骤")
except Exception as e:
    print(f"花名输入框处理时出错: {e}")


window.wait('ready', timeout=3)
# 获取所有控件
all_controls = window.descendants()

# 使用正则表达式筛选标题中包含“工具栏”的控件，并且控件类型为Hyperlink
toolbar_hyperlink_controls = [ctrl for ctrl in all_controls if re.search(
    r"工具栏", ctrl.window_text()) and ctrl.friendly_class_name() == "Hyperlink"]

# 打印匹配到的控件
for control in toolbar_hyperlink_controls:
    print(control)


# 读取CSV文件
csv_file_path = "需下载牛牛数据4.csv"
if not os.path.exists(csv_file_path):
    print(f"错误：找不到文件 {csv_file_path}")
    sys.exit(1)

with open(csv_file_path, 'r', encoding='utf-8') as file:
    csv_reader = csv.DictReader(file)
    sku_list = [row['系统SKU'] for row in csv_reader if '系统SKU' in row]

if not sku_list:
    print("错误：CSV文件中没有找到系统SKU数据")
    sys.exit(1)

# 初始化计数器
total_skus = 0
successful_queries = 0
failed_queries = 0
successful_downloads = 0
failed_downloads = 0


# 点击第一个匹配的 Hyperlink 控件
if toolbar_hyperlink_controls:
    toolbar_hyperlink_controls[0].click_input()
    window.wait('visible', timeout=5)
    # 查找弹出菜单中的所有控件
    popup_controls = window.descendants()

    # 筛选标题中包含“共享关键词”的控件
    keyword_controls = [
        ctrl for ctrl in popup_controls
        if re.search(r"共享关键词", ctrl.window_text())
    ]

    if keyword_controls:
        # 点击第一个匹配的 "共享关键词" 控件
        keyword_controls[0].click_input()
        print("成功点击 '共享关键词'")
        window.wait('ready', timeout=3)
        # 定位到输入框
        input_box = window.child_window(
            title="请输入erpsku、sellersku、asin、关键词或站点名进行搜索", control_type="Edit")

        # 将 SKU 列表复制到剪贴板
        # sku_list = "DS04604,10026467,1010200028013"
        # pyperclip.copy(sku_list)

        # if input_box.exists():
        #     # 粘贴剪贴板内容到输入框中
        #     input_box.click_input()  # 先点击输入框
        #     input_box.type_keys('^v')  # 模拟 Ctrl + V 粘贴操作
        #     time.sleep(1)  # 等待粘贴完成
        #     print("成功粘贴 SKU")
        # else:
        #     print("未找到输入框控件")
        # window.wait('ready', timeout=3)
        # # 通过 child_window 来定位并点击按钮
        # # 定位超链接控件
        # # 使用坐标信息进行点击 (基于控件的位置进行点击)
        # # 获取所有的 Button 控件并打印其信息
        # all_buttons = window.descendants(control_type="Button")
        # for idx, button in enumerate(all_buttons):
        #     print(f"Button {idx}: {button.window_text()}")

        # # 点击特定索引的按钮
        # all_buttons[5].click_input()  # 如果第一个按钮是查询按钮
        # # 定位到 "只下载自己站点数据" 超链接
        # window.wait('ready', timeout=2)
        # download_link = window.child_window(
        #     title="只下载自己站点数据", control_type="Hyperlink")

        # # 确认超链接是否存在
        # if download_link.exists():
        #     # 点击 "只下载自己站点数据" 超链接
        #     download_link.click_input()
        #     print("成功点击 '只下载自己站点数据' 超链接")
        # else:
        #     print("未找到 '只下载自己站点数据' 超链接")
        # try:
        #     save_dialog = Desktop(backend="win32").window(
        #         title_re="保存文件|Save As|Save File", top_level_only=False, enabled_only=True)

        #     # 动态等待“保存文件”弹窗出现
        #     save_dialog.wait('visible', timeout=30)
        #     print("保存文件弹窗已出现")
        #     # 模拟按回车键进行保存
        #     save_dialog.type_keys("{ENTER}")
        #     print("成功按下回车键进行保存")
        #     save_dialog.type_keys("{ENTER}")
        #     print("成功按下回车键进行保存")

        # except TimeoutError:
        #     print("等待保存文件弹窗超时")

        for sku in sku_list:
            # 将当前SKU复制到剪贴板
            pyperclip.copy(sku)

            if input_box.exists():
                # 清空输入框
                input_box.set_edit_text("")
                # 粘贴剪贴板内容到输入框中
                input_box.click_input()  # 先点击输入框
                input_box.type_keys('^v')  # 模拟 Ctrl + V 粘贴操作
                time.sleep(1)  # 等待粘贴完成
                print(f"成功粘贴 SKU: {sku}")
            else:
                print("未找到输入框控件")
                failed_queries += 1
                continue

            window.wait('ready', timeout=3)

            # 获取所有的 Button 控件
            all_buttons = window.descendants(control_type="Button")

            # 点击查询按钮
            all_buttons[5].click_input()
            # 定位到 "只下载自己站点数据" 超链接
            window.wait('ready', timeout=3000)
            download_link = window.child_window(
                title="只下载自己站点数据", control_type="Hyperlink")

            # 确认超链接是否存在
            if download_link.exists():
                # 点击 "只下载自己站点数据" 超链接
                download_link.click_input()
                print(f"成功点击 '只下载自己站点数据' 超链接，SKU: {sku}")
            else:
                print(f"未找到 '只下载自己站点数据' 超链接，SKU: {sku}")
                failed_downloads += 1
                continue

            try:
                save_dialog = Desktop(backend="win32").window(
                    title_re="保存文件|Save As|Save File", top_level_only=False, enabled_only=True)

                # 动态等待"保存文件"弹窗出现
                save_dialog.wait('visible', timeout=30000)
                print(f"保存文件弹窗已出现，SKU: {sku}")
                # 模拟按回车键进行保存
                save_dialog.type_keys("{ENTER}")
                print(f"成功按下回车键进行保存，SKU: {sku}")
                time.sleep(5)  # 等待文件保存完成
                # 等待“消息提示”弹窗出现

                # 动态等待"消息提示"弹窗出现
                try:
                    message_dialog = Desktop(backend="win32").window(
                        title="消息提示", top_level_only=False, enabled_only=True)
                    message_dialog.wait('visible', timeout=300000)
                    print("找到 '消息提示' 弹窗")

                    # 聚焦弹窗
                    message_dialog.set_focus()

                    # 模拟按下回车键
                    message_dialog.type_keys("{ENTER}")
                    message_dialog.type_keys("{ENTER}")
                    print("成功在 '消息提示' 弹窗上按下回车键")
                except TimeoutError:
                    print("等待 '消息提示' 弹窗超时")
                except Exception as e:
                    print(f"处理 '消息提示' 弹窗时出现错误: {e}")

                successful_downloads += 1
            except TimeoutError:
                print(f"等待保存文件弹窗超时，SKU: {sku}")
                failed_downloads += 1

            # 等待一段时间，以便下一次查询
            time.sleep(5)
        print("\n统计结果:")
        print(f"总SKU数: {total_skus}")
        print(f"成功查询次数: {successful_queries}")
        print(f"失败查询次数: {failed_queries}")
        print(f"成功下载次数: {successful_downloads}")
        print(f"失败下载次数: {failed_downloads}")

    else:
        print("未找到 '共享关键词' 控件")
else:
    print("未找到匹配的 Hyperlink 控件")


# 打印页面跳转后的控件树结构
# # 递归函数打印层级
# def print_layer_identifiers(element, depth=0, max_depth=10):
#     if depth > max_depth:
#         return  # 如果超过指定层级，停止递归

#     # 获取控件的基本信息
#     control_type = element.element_info.control_type
#     title = element.window_text()
#     auto_id = element.element_info.automation_id

#     # 打印缩进，根据层级深度进行缩进
#     indent = "    " * depth
#     print(f"{indent}Control Type: {control_type}, Title: {title}, Auto ID: {auto_id}")

#     # 获取子控件
#     children = element.children()

#     # 递归处理每个子控件
#     for child in children:
#         print_layer_identifiers(child, depth + 1, max_depth)

# # 打印前三层控件
# print_layer_identifiers(window)

    # 捕获print_control_identifiers的输出
    # output = io.StringIO()  # 创建一个内存缓冲区
    # sys.stdout = output     # 将标准输出重定向到内存缓冲区

    # window.print_control_identifiers()  # 打印控件标识

    # sys.stdout = sys.__stdout__  # 重置标准输出

    # # 将捕获的输出写入文件
    # file_path = "control_identifiers.txt"
    # with open(file_path, "w", encoding="utf-8") as file:
    #     file.write(output.getvalue())  # 将输出写入文件

    # print(f"控件结构已保存到 {file_path} 文件中")
