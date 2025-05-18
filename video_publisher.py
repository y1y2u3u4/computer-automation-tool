#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
蚁小二视频发布自动化工具
用于自动读取Excel表格数据并操作蚁小二应用发布视频
"""

import os
import time
import pyautogui
import pandas as pd
import pyperclip
from datetime import datetime
import pygetwindow as gw
from pathlib import Path
import logging
import sys
import platform
import traceback
import uuid
from PIL import Image

# 配置日志
def setup_logger():
    """设置日志配置"""
    log_format = '%(asctime)s - %(levelname)s - %(message)s'
    logging.basicConfig(level=logging.INFO,
                      format=log_format,
                      handlers=[
                          logging.FileHandler("video_publisher.log", encoding='utf-8'),
                          logging.StreamHandler()
                      ])
    return logging.getLogger()

logger = setup_logger()

# 检测操作系统
SYSTEM = platform.system()
if SYSTEM == 'Windows':
    MODIFIER_KEY = 'ctrl'
    logger.info("检测到Windows系统，使用ctrl键")
else:
    MODIFIER_KEY = 'command'
    logger.info(f"检测到{SYSTEM}系统，使用command键")

# 配置参数
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "蚁小二-视频工作流模板-for宫卿.xlsx")
VIDEO_FOLDER = os.path.join(BASE_DIR, "视频文件")
APP_NAME = "蚁小二"  # 应用窗口名称

# 延迟时间配置（单位：秒）
CLICK_DELAY = 0.5  # 点击操作后的延迟
TYPE_DELAY = 0.2   # 输入操作后的延迟
LOAD_DELAY = 2.0   # 页面加载延迟
TIMEOUT = 30.0     # 操作超时时间

class VideoPublisher:
    """视频发布自动化工具类"""
    
    def __init__(self):
        """初始化视频发布工具"""
        self.excel_data = None
        self.app_window = None
        self.screenshot_dir = os.path.join(BASE_DIR, "screenshots")
        
        # 创建截图目录
        if not os.path.exists(self.screenshot_dir):
            os.makedirs(self.screenshot_dir)
            
    def take_screenshot(self, name=None):
        """截取当前屏幕并保存"""
        try:
            if name is None:
                name = f"screen_{uuid.uuid4().hex[:8]}_{int(time.time())}"
            
            filename = os.path.join(self.screenshot_dir, f"{name}.png")
            screenshot = pyautogui.screenshot()
            screenshot.save(filename)
            logger.info(f"截图已保存到: {filename}")
            return filename
        except Exception as e:
            logger.error(f"截图失败: {e}")
            return None
            
    def click_with_screenshot(self, x, y, name=None, delay=CLICK_DELAY):
        """点击指定坐标并截图"""
        try:
            # 点击前截图
            before_name = f"before_click_{name if name else uuid.uuid4().hex[:8]}"
            self.take_screenshot(before_name)
            
            # 记录点击位置
            logger.info(f"点击坐标: ({x}, {y})")
            
            # 执行点击
            pyautogui.click(x, y)
            time.sleep(delay)
            
            # 点击后截图
            after_name = f"after_click_{name if name else uuid.uuid4().hex[:8]}"
            self.take_screenshot(after_name)
            
            return True
        except Exception as e:
            logger.error(f"点击失败: {e}")
            return False
    
    def load_excel_data(self):
        """加载Excel表格数据"""
        try:
            logger.info(f"尝试加载Excel文件: {EXCEL_PATH}")
            if not os.path.exists(EXCEL_PATH):
                logger.error(f"Excel文件不存在: {EXCEL_PATH}")
                return False
            
            # 读取Excel文件，不将第一行作为列名
            raw_data = pd.read_excel(EXCEL_PATH, header=None)
            logger.info(f"成功读取Excel原始数据，共{len(raw_data)}行")
            
            # 检查数据是否至少有2行（标题行和字段名行）
            if len(raw_data) < 2:
                logger.error("Excel文件数据不足，至少需要2行")
                return False
            
            # 获取第二行作为真正的列名
            column_names = raw_data.iloc[1].tolist()
            logger.info(f"实际列名: {column_names}")
            
            # 使用第二行作为列名，并只保留第三行及之后的数据
            data = raw_data.iloc[2:].reset_index(drop=True)
            data.columns = column_names
            
            # 删除空行
            data = data.dropna(how='all').reset_index(drop=True)
            
            self.excel_data = data
            logger.info(f"处理后的数据条数: {len(self.excel_data)}")
            
            # 显示实际的列名
            logger.info(f"实际列名: {list(self.excel_data.columns)}")
            
            # 检查必要的列是否存在
            # 根据您的Excel文件实际字段名调整
            required_columns = ['序号', '账号', '标题', '描述']
            available_columns = list(self.excel_data.columns)
            
            # 检查必要列是否存在
            missing_columns = []
            for col in required_columns:
                found = False
                for avail_col in available_columns:
                    if col in str(avail_col):
                        found = True
                        break
                if not found:
                    missing_columns.append(col)
            
            if missing_columns:
                logger.error(f"Excel文件缺少必要的列: {missing_columns}")
                return False
            
            # 显示前两行数据作为参考
            logger.info(f"前两行数据:\n{self.excel_data.head(2)}")
            return True
        except Exception as e:
            logger.error(f"加载Excel数据失败: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def find_video_file(self, video_name):
        """根据视频名称查找视频文件的完整路径"""
        try:
            video_folder = Path(VIDEO_FOLDER)
            
            logger.info(f"查找视频文件: '{video_name}' 在目录: {VIDEO_FOLDER}")
            
            # 检查视频文件夹是否存在
            if not video_folder.exists():
                logger.error(f"视频文件夹不存在: {VIDEO_FOLDER}")
                return None
            
            # 支持的视频文件扩展名
            video_extensions = ['.mp4', '.mov', '.avi', '.mkv']
            
            # 查找匹配的视频文件
            for ext in video_extensions:
                potential_file = video_folder / f"{video_name}{ext}"
                logger.debug(f"尝试查找文件: {potential_file}")
                if potential_file.exists():
                    logger.info(f"找到精确匹配的视频文件: {potential_file}")
                    return str(potential_file)
            
            # 如果没有精确匹配，尝试查找包含视频名称的文件
            logger.info("未找到精确匹配，尝试模糊匹配...")
            for file in video_folder.glob('*'):
                if file.is_file() and file.suffix.lower() in video_extensions and video_name in file.stem:
                    logger.info(f"找到模糊匹配的视频文件: {file}")
                    return str(file)
            
            # 列出文件夹中的所有视频文件，帮助调试
            all_videos = [f for f in video_folder.glob('*') if f.is_file() and f.suffix.lower() in video_extensions]
            logger.warning(f"警告: 未找到视频文件 '{video_name}'")
            logger.info(f"文件夹中的视频文件列表: {[f.name for f in all_videos]}")
            return None
        except Exception as e:
            logger.error(f"查找视频文件时出错: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def activate_app(self):
        """激活蚁小二应用窗口"""
        try:
            logger.info(f"尝试查找并激活应用窗口: {APP_NAME}")
            
            # 列出所有窗口，帮助调试
            try:
                # Mac系统上的窗口列表获取方式
                all_windows = gw.getAllTitles()
                logger.info(f"当前所有窗口标题: {all_windows}")
                
                # 在Mac上查找包含指定名称的窗口
                matching_windows = []
                for window_title in all_windows:
                    if APP_NAME in window_title:
                        logger.info(f"找到匹配窗口: {window_title}")
                        matching_windows.append(window_title)
                
                if matching_windows:
                    # 如果找到匹配窗口，尝试激活第一个
                    target_window_title = matching_windows[0]
                    logger.info(f"尝试激活窗口: {target_window_title}")
                    
                    # 在Mac上使用AppleScript激活窗口
                    activate_script = f"""
                    tell application "System Events"
                        set frontmost of every process whose name contains "{APP_NAME}" to true
                    end tell
                    """
                    os.system(f"osascript -e '{activate_script}'")
                    logger.info(f"已尝试通过AppleScript激活窗口")
                    time.sleep(LOAD_DELAY)
                    return True
                else:
                    logger.error(f"找不到应用窗口: {APP_NAME}")
                    return False
                    
            except AttributeError:
                # 如果上面的Mac特定方法不可用，尝试替代方法
                logger.info("尝试使用替代方法激活窗口")
                
                # 使用pyautogui模拟点击窗口标题栏
                screen_width, screen_height = pyautogui.size()
                # 点击屏幕上方中间位置（可能是窗口标题栏）
                pyautogui.click(screen_width // 2, 20)
                time.sleep(CLICK_DELAY)
                
                logger.info("已尝试激活窗口，继续执行")
                return True
                
        except Exception as e:
            logger.error(f"激活应用窗口失败: {e}")
            logger.error(traceback.format_exc())
            
            # 即使激活失败也继续执行，因为用户可能已手动激活了窗口
            logger.warning("窗口激活失败，但将继续执行。请确保蚁小二应用已经打开并在前台显示")
            return True
    
    def run_applescript(self, script):
        """运行AppleScript脚本"""
        try:
            # 将脚本写入临时文件
            script_path = os.path.join(self.screenshot_dir, f"script_{uuid.uuid4().hex[:8]}.scpt")
            with open(script_path, 'w') as f:
                f.write(script)
            
            # 执行脚本
            cmd = f"osascript {script_path}"
            logger.info(f"执行AppleScript: {cmd}")
            result = os.system(cmd)
            
            # 删除临时脚本文件
            os.remove(script_path)
            
            return result == 0
        except Exception as e:
            logger.error(f"执行AppleScript失败: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def click_publish_button(self):
        """点击发布按钮"""
        try:
            # 先截图查看当前界面
            self.take_screenshot("before_publish_button")
            
            # 使用AppleScript模拟点击发布按钮
            # 根据截图，尝试使用UI元素描述方式点击左侧边栏中的发布按钮
            script = """
            tell application "蚁小二4.0" to activate
            delay 1
            
            tell application "System Events"
                tell process "蚁小二4.0"
                    -- 尝试点击左侧边栏中的发布按钮
                    -- 方法1: 尝试通过UI元素属性查找并点击
                    try
                        -- 尝试查找包含“发布”文本的按钮
                        click UI element "发布" of window 1
                    on error
                        try
                            -- 尝试查找左侧边栏中的所有按钮
                            set sidebar_buttons to buttons of window 1
                            -- 假设发布按钮是第2个按钮
                            click item 2 of sidebar_buttons
                        on error
                            try
                                -- 尝试查找左侧边栏的所有图标
                                set sidebar_icons to images of window 1
                                -- 假设发布图标是第2个图标
                                click item 2 of sidebar_icons
                            on error
                                -- 尝试通过点击左侧边栏中的所有项目
                                set all_elements to UI elements of window 1
                                repeat with elem in all_elements
                                    try
                                        if name of elem contains "发布" then
                                            click elem
                                            exit repeat
                                        end if
                                    end try
                                end repeat
                            end try
                        end try
                    end try
                end tell
            end tell
            """
            
            success = self.run_applescript(script)
            
            # 截图查看结果
            self.take_screenshot("after_publish_button")
            
            if success:
                logger.info("已成功点击发布按钮")
            else:
                logger.warning("使用AppleScript点击发布按钮可能失败，尝试使用坐标点击")
                # 备用方案：使用坐标点击
                screen_width, screen_height = pyautogui.size()
                publish_button_x = int(screen_width * 0.12)
                publish_button_y = int(screen_height * 0.12)
                self.click_with_screenshot(publish_button_x, publish_button_y, "publish_button_fallback")
            
            # 等待一些时间确保操作生效
            time.sleep(LOAD_DELAY)
            
            return True
        except Exception as e:
            logger.error(f"点击发布按钮失败: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def click_new_publish_button(self):
        """点击新建发布按钮"""
        try:
            # 先截图查看当前界面
            self.take_screenshot("before_new_publish_button")
            
            # 使用AppleScript模拟点击新建发布按钮
            script = """
            tell application "System Events"
                tell process "蚁小二4.0"
                    -- 确保应用处于前台
                    set frontmost to true
                    delay 0.5
                    
                    -- 尝试点击新建发布按钮
                    -- 方法1: 尝试通过名称找到并点击
                    try
                        click button "新建发布" of window 1
                    on error
                        -- 方法2: 尝试点击右上角的按钮
                        try
                            click button 1 of group 1 of window 1
                        on error
                            -- 方法3: 使用键盘快捷键 (通常是Command+N)
                            try
                                keystroke "n" using command down
                            on error
                                -- 方法4: 直接点击坐标
                                tell application "System Events" to click at {1285, 58}
                            end try
                        end try
                    end try
                end tell
            end tell
            """
            
            success = self.run_applescript(script)
            
            # 截图查看结果
            self.take_screenshot("after_new_publish_button")
            
            if success:
                logger.info("已成功点击新建发布按钮")
            else:
                logger.warning("使用AppleScript点击新建发布按钮可能失败，尝试使用坐标点击")
                # 备用方案：使用坐标点击
                screen_width, screen_height = pyautogui.size()
                new_publish_x = int(screen_width * 0.85)
                new_publish_y = int(screen_height * 0.06)
                self.click_with_screenshot(new_publish_x, new_publish_y, "new_publish_button_fallback")
            
            # 等待一些时间确保操作生效
            time.sleep(LOAD_DELAY)
            
            return True
        except Exception as e:
            logger.error(f"点击新建发布按钮失败: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def select_video(self, video_name):
        """选择要发布的视频"""
        try:
            # 获取屏幕尺寸
            screen_width, screen_height = pyautogui.size()
            logger.info(f"屏幕尺寸: {screen_width}x{screen_height}")
            
            # 根据截图二，点击上传区域
            upload_area_x = int(screen_width * 0.5)  # 中间位置
            upload_area_y = int(screen_height * 0.4)  # 大约在中间靠上的位置
            
            logger.info(f"点击上传区域: ({upload_area_x}, {upload_area_y})")
            pyautogui.click(upload_area_x, upload_area_y)
            time.sleep(CLICK_DELAY * 2)
            
            # 查找视频文件
            video_path = self.find_video_file(video_name)
            if not video_path:
                logger.error(f"未找到视频文件: {video_name}")
                return False
            
            # 将视频路径复制到剪贴板
            logger.info(f"复制视频路径到剪贴板: {video_path}")
            pyperclip.copy(video_path)
            time.sleep(TYPE_DELAY)
            
            # 粘贴路径并按回车
            logger.info(f"粘贴路径并按回车")
            pyautogui.hotkey(MODIFIER_KEY, 'v')
            time.sleep(TYPE_DELAY)
            pyautogui.press('return')
            
            # 等待视频加载
            logger.info("等待视频加载...")
            start_time = time.time()
            while time.time() - start_time < TIMEOUT:
                time.sleep(LOAD_DELAY)
                # 这里可以添加检查视频是否加载完成的逻辑
                # 例如检查界面上的某个元素是否出现
                break  # 暂时简化处理，直接等待固定时间
            
            logger.info(f"已选择视频: {video_name}")
            return True
        except Exception as e:
            logger.error(f"选择视频失败: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def select_account(self, account_name):
        """选择发布账号"""
        try:
            # 点击"下一步"按钮进入账号选择页面
            screen_width, screen_height = pyautogui.size()
            
            next_button_x = int(screen_width * 0.81)  # 右下角的下一步按钮
            next_button_y = int(screen_height * 0.82)  # 页面底部的位置
            
            pyautogui.click(next_button_x, next_button_y)
            time.sleep(LOAD_DELAY)
            
            # 在账号列表中查找并点击指定账号
            # 这里需要根据实际界面进行调整，可能需要滚动或搜索
            # 简化处理：假设账号显示在列表中，遍历点击匹配的账号
            
            # 假设账号显示的位置在页面中部左侧
            account_area_x = int(screen_width * 0.3)  # 列表中间位置
            account_area_y = int(screen_height * 0.5)  # 大约在中间的位置
            
            # 简化处理：点击第一个账号位置
            pyautogui.click(account_area_x, account_area_y)
            time.sleep(CLICK_DELAY)
            
            # 点击"下一步"按钮
            pyautogui.click(next_button_x, next_button_y)
            time.sleep(LOAD_DELAY)
            
            print(f"已选择账号: {account_name}")
            return True
        except Exception as e:
            print(f"选择账号失败: {e}")
            return False
    
    def fill_info(self, title, description, location, schedule_time=None):
        """填写发布信息"""
        try:
            screen_width, screen_height = pyautogui.size()
            
            # 填写标题
            title_x = int(screen_width * 0.5)  # 标题输入框位置
            title_y = int(screen_height * 0.18)  # 页面上方的位置
            
            pyautogui.click(title_x, title_y)
            time.sleep(CLICK_DELAY)
            pyautogui.hotkey('command', 'a')  # 全选已有内容
            time.sleep(TYPE_DELAY)
            
            pyperclip.copy(title)
            pyautogui.hotkey(MODIFIER_KEY, 'v')  # 粘贴标题
            time.sleep(TYPE_DELAY)
            
            # 填写描述
            desc_x = int(screen_width * 0.5)  # 描述输入框位置
            desc_y = int(screen_height * 0.35)  # 标题下方的位置
            
            pyautogui.click(desc_x, desc_y)
            time.sleep(CLICK_DELAY)
            
            pyperclip.copy(description)
            pyautogui.hotkey(MODIFIER_KEY, 'v')  # 粘贴描述
            time.sleep(TYPE_DELAY)
            
            # 选择位置（如果提供）
            if location and location.strip():
                location_x = int(screen_width * 0.5)  # 位置选择区域
                location_y = int(screen_height * 0.5)  # 描述下方的位置
                
                pyautogui.click(location_x, location_y)
                time.sleep(CLICK_DELAY)
                
                # 这里可能需要更复杂的逻辑来选择具体位置
                # 简化处理：点击位置下拉框并选择第一个选项
                location_option_y = location_y + 50  # 下拉菜单中的第一个选项
                pyautogui.click(location_x, location_option_y)
                time.sleep(CLICK_DELAY)
            
            # 设置定时发送（如果需要）
            if schedule_time:
                # 点击定时发送选项
                schedule_x = int(screen_width * 0.3)  # 定时发送选择框
                schedule_y = int(screen_height * 0.65)  # 位置下方的位置
                
                pyautogui.click(schedule_x, schedule_y)
                time.sleep(CLICK_DELAY)
                
                # 选择时间（需要根据实际界面调整）
                # 这里简化处理，假设点击后会弹出日期选择器
                time_picker_x = int(screen_width * 0.5)
                time_picker_y = int(screen_height * 0.7)
                
                pyautogui.click(time_picker_x, time_picker_y)
                time.sleep(CLICK_DELAY)
                
                # 输入时间或选择时间（需要根据实际界面调整）
                pyperclip.copy(schedule_time)
                pyautogui.hotkey(MODIFIER_KEY, 'v')
                time.sleep(TYPE_DELAY)
                pyautogui.press('return')
                time.sleep(CLICK_DELAY)
            
            print("已填写发布信息")
            return True
        except Exception as e:
            print(f"填写发布信息失败: {e}")
            return False
    
    def click_publish(self):
        """点击一键发布按钮"""
        try:
            # 点击一键发布按钮
            screen_width, screen_height = pyautogui.size()
            
            publish_button_x = int(screen_width * 0.81)  # 右侧的按钮
            publish_button_y = int(screen_height * 0.07)  # 顶部的位置
            
            pyautogui.click(publish_button_x, publish_button_y)
            time.sleep(LOAD_DELAY)
            
            # 选择浏览器发布
            browser_publish_x = int(screen_width * 0.5)  # 中间选项
            browser_publish_y = int(screen_height * 0.3)  # 弹窗中间的位置
            
            pyautogui.click(browser_publish_x, browser_publish_y)
            time.sleep(LOAD_DELAY * 2)  # 发布需要等待较长时间
            
            print("已点击一键发布，选择浏览器发布")
            return True
        except Exception as e:
            print(f"发布失败: {e}")
            return False
    
    def process_row(self, row):
        """处理Excel表格中的一行数据"""
        try:
            # 根据Excel的实际列名提取数据
            # 构建视频名称：客户-创作日期-序号
            client = str(row.get('客户', ''))
            creation_date = str(row.get('创作日期', ''))
            sequence = str(row.get('序号', ''))
            
            # 组合成视频名称
            video_name = f"{client}-{creation_date}-{sequence}".strip('-')
            
            # 获取其他字段
            account = row.get('账号')
            title = row.get('标题')
            description = row.get('描述')
            location = row.get('位置')
            
            # 处理定时发送
            schedule_flag = row.get('定时发送')
            schedule_time = None
            if schedule_flag and str(schedule_flag).lower() in ['是', 'yes', 'y', '1', 'true']:
                schedule_time = row.get('定时发布')
                
            # 输出提取的数据信息
            logger.info(f"提取的数据: \n"
                       f"  视频名称: {video_name}\n"
                       f"  账号: {account}\n"
                       f"  标题: {title}\n"
                       f"  描述: {description}\n"
                       f"  位置: {location}\n"
                       f"  定时发送: {schedule_flag}\n"
                       f"  定时时间: {schedule_time}")
            
            print(f"\n开始处理: {video_name}")
            
            # 1. 点击发布按钮
            if not self.click_publish_button():
                return False
            
            # 2. 点击新建发布按钮
            if not self.click_new_publish_button():
                return False
            
            # 3. 选择视频文件
            if not self.select_video(video_name):
                return False
            
            # 4. 选择账号
            if not self.select_account(account):
                return False
            
            # 5. 填写发布信息
            if not self.fill_info(title, description, location, schedule_time):
                return False
            
            # 6. 点击发布
            if not self.click_publish():
                return False
            
            print(f"成功处理: {video_name}\n")
            return True
            
        except Exception as e:
            print(f"处理行数据失败: {e}")
            return False
    
    def run(self):
        """运行视频发布自动化流程"""
        try:
            logger.info("=== 开始蚁小二视频发布自动化流程 ===")
            logger.info(f"操作系统: {SYSTEM}")
            logger.info(f"Excel路径: {EXCEL_PATH}")
            logger.info(f"视频文件夹: {VIDEO_FOLDER}")
            
            # 1. 检查文件和文件夹是否存在
            if not os.path.exists(EXCEL_PATH):
                logger.error(f"Excel文件不存在: {EXCEL_PATH}")
                return
                
            if not os.path.exists(VIDEO_FOLDER):
                logger.error(f"视频文件夹不存在: {VIDEO_FOLDER}")
                return
            
            # 2. 加载Excel数据
            logger.info("正在加载Excel数据...")
            if not self.load_excel_data():
                logger.error("加载Excel数据失败，程序终止")
                return
            
            # 3. 激活应用窗口
            logger.info("正在激活应用窗口...")
            if not self.activate_app():
                logger.error("激活应用窗口失败，程序终止")
                return
            
            # 4. 处理每一行数据
            success_count = 0
            fail_count = 0
            
            total_rows = len(self.excel_data)
            logger.info(f"共有{total_rows}条数据需要处理")
            
            for index, row in self.excel_data.iterrows():
                logger.info(f"\n=== 处理第 {index+1}/{total_rows} 条数据 ===")
                try:
                    if self.process_row(row):
                        success_count += 1
                    else:
                        fail_count += 1
                except Exception as e:
                    logger.error(f"处理第{index+1}条数据时发生未捕获的异常: {e}")
                    logger.error(traceback.format_exc())
                    fail_count += 1
                
                # 等待一段时间，避免操作过快
                logger.info(f"等待{LOAD_DELAY * 2}秒后处理下一条数据...")
                time.sleep(LOAD_DELAY * 2)
            
            # 5. 打印统计结果
            logger.info("\n=== 自动化发布完成！统计结果 ===")
            logger.info(f"总记录数: {total_rows}")
            logger.info(f"成功处理: {success_count}")
            logger.info(f"失败处理: {fail_count}")
            
        except Exception as e:
            logger.error(f"执行过程中发生未捕获的异常: {e}")
            logger.error(traceback.format_exc())


if __name__ == "__main__":
    try:
        # 显示欢迎信息
        logger.info("=== 蚁小二视频发布自动化工具 ===")
        logger.info(f"当前工作目录: {os.getcwd()}")
        
        # 允许用户有时间切换到蚁小二应用
        logger.info("请在5秒内切换到蚁小二应用窗口...")
        for i in range(5, 0, -1):
            logger.info(f"{i}秒...")
            time.sleep(1)
        
        # 创建并运行视频发布工具
        publisher = VideoPublisher()
        publisher.run()
        
    except KeyboardInterrupt:
        logger.info("\n用户中断程序执行")
    except Exception as e:
        logger.error(f"程序执行过程中发生未捕获的异常: {e}")
        logger.error(traceback.format_exc())
    finally:
        logger.info("程序执行结束")
