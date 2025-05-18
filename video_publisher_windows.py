#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
蚁小二视频发布自动化工具 - Windows版本
使用pywinauto库实现UI自动化，用于自动读取Excel表格数据并操作蚁小二应用发布视频
"""

import os
import time
import pandas as pd
import pyperclip
from datetime import datetime
import logging
import sys
import traceback
import uuid
from PIL import Image
from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys
import re

# 配置日志
def setup_logger():
    """设置日志配置"""
    log_format = '%(asctime)s - %(levelname)s - %(message)s'
    logging.basicConfig(level=logging.INFO,
                      format=log_format,
                      handlers=[
                          logging.StreamHandler(),
                          logging.FileHandler('video_publisher_windows.log', encoding='utf-8')
                      ])
    return logging.getLogger(__name__)

# 创建日志对象
logger = setup_logger()

# 常量定义
LOAD_DELAY = 2  # 加载延迟时间(秒)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))  # 脚本所在目录
SCREENSHOTS_DIR = os.path.join(SCRIPT_DIR, "screenshots")  # 截图保存目录

# 确保截图目录存在
os.makedirs(SCREENSHOTS_DIR, exist_ok=True)

class VideoPublisher:
    """蚁小二视频发布自动化工具类"""
    
    def __init__(self):
        """初始化视频发布工具"""
        self.app = None  # 应用程序对象
        self.window = None  # 主窗口对象
        self.excel_path = os.path.join(SCRIPT_DIR, "蚁小二-视频工作流模板-for宫卿.xlsx")
        self.data = None  # 存储Excel数据
        self.current_row = None  # 当前处理的行数据
        
        logger.info("视频发布工具初始化完成")
    
    def connect_to_app(self):
        """连接到蚁小二应用"""
        try:
            # 尝试连接到已运行的蚁小二应用
            logger.info("尝试连接到蚁小二应用...")
            
            # 列出所有窗口
            logger.info("列出所有可见窗口:")
            desktop = Desktop(backend="uia")
            windows = desktop.windows()
            
            for i, win in enumerate(windows):
                try:
                    title = win.window_text()
                    logger.info(f"窗口 {i+1}: '{title}'")
                except Exception as e:
                    logger.info(f"窗口 {i+1}: <无法获取标题> (错误: {e})")
            
            # 尝试查找包含"蚁小二"或相关关键词的窗口
            target_window = None
            keywords = ["蚁小二", "小二", "蚁", "ant", "Ant"]
            
            for win in windows:
                try:
                    title = win.window_text()
                    for keyword in keywords:
                        if keyword in title:
                            target_window = win
                            logger.info(f"找到包含关键词 '{keyword}' 的窗口: '{title}'")
                            break
                    if target_window:
                        break
                except:
                    continue
            
            if target_window:
                # 直接使用找到的窗口
                self.window = target_window
                self.window.set_focus()
                logger.info(f"成功连接到窗口: '{self.window.window_text()}'")
                return True
            
            # 如果没有找到匹配的窗口，尝试使用第一个非空标题的窗口
            for win in windows:
                try:
                    title = win.window_text()
                    if title.strip() != "":
                        self.window = win
                        self.window.set_focus()
                        logger.info(f"没有找到蚁小二窗口，使用第一个有效窗口: '{title}'")
                        return True
                except:
                    continue
            
            # 如果上述方法都失败，尝试使用原始方法
            try:
                self.app = Application(backend="uia").connect(best_match="蚁小二")
                self.window = self.app.window(best_match="蚁小二")
                self.window.set_focus()
                logger.info("成功连接到蚁小二应用")
                return True
            except Exception as e2:
                logger.error(f"原始方法连接失败: {e2}")
                
            logger.error("所有连接方法都失败了")
            return False
        except Exception as e:
            logger.error(f"连接到蚁小二应用失败: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def take_screenshot(self, name):
        """截取当前屏幕并保存"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{name}_{timestamp}.png"
            filepath = os.path.join(SCREENSHOTS_DIR, filename)
            
            # 使用pywinauto的capture_as_image方法截图
            if self.window:
                img = self.window.capture_as_image()
                img.save(filepath)
                logger.info(f"已保存截图: {filepath}")
                return filepath
            else:
                logger.warning("窗口对象不存在，无法截图")
                return None
        except Exception as e:
            logger.error(f"截图失败: {e}")
            return None
    
    def load_excel_data(self):
        """加载Excel数据"""
        try:
            logger.info(f"正在加载Excel文件: {self.excel_path}")
            
            # 读取Excel文件，跳过第一行，使用第二行作为列名
            df = pd.read_excel(self.excel_path, skiprows=1)
            
            # 检查必要的列是否存在
            required_columns = ['序号', '账号', '标题', '描述']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                logger.error(f"Excel文件缺少必要的列: {', '.join(missing_columns)}")
                return False
            
            # 存储数据
            self.data = df
            logger.info(f"成功加载Excel数据，共{len(df)}行")
            return True
        except Exception as e:
            logger.error(f"加载Excel数据失败: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def click_publish_button(self):
        """点击发布按钮"""
        try:
            # 先截图查看当前界面
            self.take_screenshot("before_publish_button")
            
            logger.info("尝试点击发布按钮")
            
            # 方法1: 根据UI结构中的信息，直接点击特定位置的发布按钮
            try:
                # 使用精确的坐标点击发布按钮（根据UI结构日志，位置大约在(275,216,303,234)）
                # 点击中心点
                x_center = (275 + 303) // 2
                y_center = (216 + 234) // 2
                
                # 使用pywinauto的click_input方法点击窗口的特定位置
                self.window.click_input(coords=(x_center, y_center))
                logger.info(f"点击了发布按钮的位置坐标: ({x_center}, {y_center})")
                time.sleep(LOAD_DELAY)
                self.take_screenshot("after_publish_button_position")
                return True
            except Exception as e:
                logger.warning(f"通过精确坐标点击发布按钮失败: {e}")
            
            # 方法2: 尝试通过文本内容查找发布按钮
            try:
                # 查找精确匹配"发布"文本的控件
                publish_elements = self.window.descendants(title="发布")
                
                if publish_elements:
                    # 点击第一个匹配的元素
                    publish_elements[0].click_input()
                    logger.info("通过精确文本内容找到并点击了发布按钮")
                    time.sleep(LOAD_DELAY)
                    self.take_screenshot("after_publish_button_exact")
                    return True
                else:
                    # 尝试模糊匹配
                    publish_elements = self.window.descendants(title_re=".*发布.*")
                    if publish_elements:
                        publish_elements[0].click_input()
                        logger.info("通过模糊文本匹配找到并点击了发布按钮")
                        time.sleep(LOAD_DELAY)
                        self.take_screenshot("after_publish_button_fuzzy")
                        return True
                    else:
                        logger.warning("未找到包含'发布'文本的控件，尝试其他方法")
            except Exception as e:
                logger.warning(f"通过文本内容查找发布按钮失败: {e}")
            
            # 方法3: 尝试点击左侧菜单栏的第二个选项（根据UI结构日志）
            try:
                # 获取所有控件
                all_controls = self.window.descendants()
                
                # 尝试点击左侧菜单的第二个选项（发布）
                # 根据UI结构日志，它在控件列表的前几个位置
                for i, ctrl in enumerate(all_controls[:20]):
                    try:
                        text = ctrl.window_text()
                        if "发布" in text:
                            ctrl.click_input()
                            logger.info(f"点击了第{i+1}个控件，文本包含'发布': '{text}'")
                            time.sleep(LOAD_DELAY)
                            self.take_screenshot(f"after_publish_control_{i+1}")
                            return True
                    except Exception as e:
                        continue
                
                logger.warning("在前20个控件中未找到包含'发布'文本的控件")
            except Exception as e:
                logger.warning(f"尝试点击左侧菜单失败: {e}")
            
            # 方法4: 尝试使用快捷键
            try:
                # 将焦点设置到窗口
                self.window.set_focus()
                # 发送Alt+F快捷键（发布的首字母）
                send_keys("%f")
                logger.info("尝试使用Alt+F快捷键打开发布菜单")
                time.sleep(LOAD_DELAY)
                self.take_screenshot("after_publish_shortcut")
                return True
            except Exception as e:
                logger.warning(f"使用快捷键打开发布菜单失败: {e}")
            
            logger.error("所有尝试点击发布按钮的方法都失败了")
            return False
        except Exception as e:
            logger.error(f"点击发布按钮时出错: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def click_new_publish_button(self):
        """点击新建发布按钮"""
        try:
            # 先截图查看当前界面
            self.take_screenshot("before_new_publish_button")
            
            logger.info("尝试点击新增发布按钮")
            
            # 方法1: 根据UI结构日志中的信息，直接点击新增发布按钮
            try:
                # 查找包含"新增发布"文本的所有控件
                new_publish_elements = self.window.descendants(title="新增发布")
                
                if new_publish_elements:
                    # 点击第一个匹配的元素
                    new_publish_elements[0].click_input()
                    logger.info("通过精确文本内容找到并点击了新增发布按钮")
                    time.sleep(LOAD_DELAY)
                    self.take_screenshot("after_new_publish_button")
                    return True
                else:
                    logger.warning("未找到包含'新增发布'文本的控件，尝试其他方法")
            except Exception as e:
                logger.warning(f"通过精确文本内容查找新增发布按钮失败: {e}")
            
            # 方法2: 尝试点击“暂无发布任务，点击新增发布任务吧”提示文本
            try:
                # 查找包含提示文本的控件
                prompt_elements = self.window.descendants(title_re=".*暂无发布任务.*新增发布.*")
                
                if prompt_elements:
                    # 点击提示文本
                    prompt_elements[0].click_input()
                    logger.info("点击了提示文本中的新增发布任务")
                    time.sleep(LOAD_DELAY)
                    self.take_screenshot("after_prompt_text")
                    return True
                else:
                    logger.warning("未找到提示文本")
            except Exception as e:
                logger.warning(f"点击提示文本失败: {e}")
            
            # 方法3: 尝试使用模糊匹配查找新增发布相关的控件
            try:
                # 尝试多种可能的匹配方式
                keywords = ["新增发布", "新增", "新建发布", "新建"]
                
                for keyword in keywords:
                    try:
                        elements = self.window.descendants(title_re=f".*{keyword}.*")
                        if elements:
                            elements[0].click_input()
                            logger.info(f"通过关键词 '{keyword}' 找到并点击了按钮")
                            time.sleep(LOAD_DELAY)
                            self.take_screenshot(f"after_{keyword}_button")
                            return True
                    except Exception as e:
                        logger.warning(f"使用关键词 '{keyword}' 查找按钮失败: {e}")
                
                logger.warning("所有关键词匹配都失败了")
            except Exception as e:
                logger.warning(f"尝试模糊匹配失败: {e}")
            
            # 方法4: 尝试查找所有控件并匹配包含“新增发布”的文本
            try:
                all_controls = self.window.descendants()
                
                for i, ctrl in enumerate(all_controls):
                    try:
                        text = ctrl.window_text()
                        # 检查文本是否包含新增发布相关关键词
                        if any(keyword in text for keyword in ["新增发布", "新建发布", "新增"]):
                            ctrl.click_input()
                            logger.info(f"点击了第{i+1}个控件，文本为: '{text}'")
                            time.sleep(LOAD_DELAY)
                            self.take_screenshot(f"after_control_{i+1}")
                            return True
                    except Exception:
                        # 忽略个别控件的错误
                        continue
                
                logger.warning("未找到包含新增发布相关关键词的控件")
            except Exception as e:
                logger.warning(f"遍历所有控件失败: {e}")
            
            # 方法5: 尝试使用快捷键
            try:
                # 将焦点设置到窗口
                self.window.set_focus()
                # 发送Ctrl+N快捷键（大多数应用的新建快捷键）
                send_keys("^n")
                logger.info("尝试使用Ctrl+N快捷键创建新发布")
                time.sleep(LOAD_DELAY)
                self.take_screenshot("after_new_shortcut")
                return True
            except Exception as e:
                logger.warning(f"使用快捷键创建新发布失败: {e}")
            
            logger.error("所有尝试点击新增发布按钮的方法都失败了")
            return False
        except Exception as e:
            logger.error(f"点击新增发布按钮时出错: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def click_video_button(self):
        """点击视频按钮"""
        try:
            # 先截图查看当前界面
            self.take_screenshot("before_video_button")
            
            logger.info("尝试点击视频按钮")
            
            # 方法1: 通过精确文本内容查找视频按钮
            try:
                # 查找包含"视频"文本的所有控件
                video_elements = self.window.descendants(title="视频")
                
                if video_elements:
                    # 点击第一个匹配的元素
                    video_elements[0].click_input()
                    logger.info("通过精确文本内容找到并点击了视频按钮")
                    time.sleep(LOAD_DELAY)
                    self.take_screenshot("after_video_button")
                    return True
                else:
                    logger.warning("未找到包含'视频'文本的控件，尝试其他方法")
            except Exception as e:
                logger.warning(f"通过精确文本内容查找视频按钮失败: {e}")
            
            # 方法2: 根据截图中的位置点击视频按钮
            try:
                # 从截图中可以看到视频按钮大约在右侧选项卡的顶部
                # 根据截图估计坐标大约在(880, 150)附近
                self.window.click_input(coords=(880, 150))
                logger.info(f"通过坐标点击了视频按钮位置(880, 150)")
                time.sleep(LOAD_DELAY)
                self.take_screenshot("after_video_button_position")
                return True
            except Exception as e:
                logger.warning(f"通过坐标点击视频按钮失败: {e}")
            
            # 方法3: 遍历所有控件并查找包含视频的文本
            try:
                all_controls = self.window.descendants()
                
                for i, ctrl in enumerate(all_controls):
                    try:
                        text = ctrl.window_text()
                        if "视频" in text:
                            ctrl.click_input()
                            logger.info(f"点击了第{i+1}个控件，文本包含'视频': '{text}'")
                            time.sleep(LOAD_DELAY)
                            self.take_screenshot(f"after_video_control_{i+1}")
                            return True
                    except Exception:
                        # 忽略个别控件的错误
                        continue
                
                logger.warning("未找到包含'视频'文本的控件")
            except Exception as e:
                logger.warning(f"遍历所有控件查找视频按钮失败: {e}")
            
            logger.error("所有尝试点击视频按钮的方法都失败了")
            return False
        except Exception as e:
            logger.error(f"点击视频按钮时出错: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def run(self):
        """运行视频发布流程"""
        try:
            logger.info("开始运行视频发布流程")
            
            # 加载Excel数据
            if not self.load_excel_data():
                logger.error("加载Excel数据失败，终止流程")
                return False
            
            # 连接到蚁小二应用
            if not self.connect_to_app():
                logger.error("连接到蚁小二应用失败，终止流程")
                return False
            
            # 截图查看初始状态
            self.take_screenshot("initial_state")
            
            # 点击发布按钮
            if not self.click_publish_button():
                logger.error("点击发布按钮失败，终止流程")
                return False
            
            # 点击新建发布按钮
            if not self.click_new_publish_button():
                logger.error("点击新建发布按钮失败，终止流程")
                return False
            
            # 处理每一行数据
            for index, row in self.data.iterrows():
                self.current_row = row
                logger.info(f"正在处理第{index+1}行数据，序号: {row['序号']}")
                
                # 点击视频按钮
                if not self.click_video_button():
                    logger.error(f"点击视频按钮失败，跳过第{index+1}行数据")
                    continue
                
                # TODO: 实现其他发布流程
                
            logger.info("视频发布流程完成")
            return True
        except Exception as e:
            logger.error(f"视频发布流程执行过程中发生错误: {e}")
            logger.error(traceback.format_exc())
            return False


def print_ui_structure(window):
    """打印UI结构，用于调试"""
    try:
        logger.info("正在获取UI结构...")
        
        # 获取所有控件
        all_controls = window.descendants()
        
        # 打印控件信息
        logger.info(f"总共找到 {len(all_controls)} 个控件")
        
        # 打印前20个控件的详细信息
        for i, ctrl in enumerate(all_controls[:20]):
            try:
                # 获取控件类型和文本
                ctrl_type = ctrl.control_type() if hasattr(ctrl, 'control_type') and callable(ctrl.control_type) else "Unknown"
                ctrl_text = ctrl.window_text() if hasattr(ctrl, 'window_text') and callable(ctrl.window_text) else "No Text"
                
                # 获取控件位置
                rect = None
                try:
                    rect = ctrl.rectangle()
                except:
                    pass
                
                pos_str = f"({rect.left},{rect.top},{rect.right},{rect.bottom})" if rect else "Unknown"
                
                logger.info(f"Control {i+1}: Type={ctrl_type}, Text='{ctrl_text}', Position={pos_str}")
            except Exception as e:
                logger.warning(f"Failed to get info for control {i+1}: {e}")
        
        # 查找特定控件
        logger.info("查找包含'发布'文本的控件:")
        publish_controls = [ctrl for ctrl in all_controls if "发布" in ctrl.window_text()]
        for i, ctrl in enumerate(publish_controls):
            logger.info(f"Found publish control {i+1}: {ctrl.window_text()}")
        
        logger.info("查找所有按钮控件:")
        button_controls = [ctrl for ctrl in all_controls if hasattr(ctrl, 'control_type') and callable(ctrl.control_type) and ctrl.control_type() == "Button"]
        for i, btn in enumerate(button_controls[:10]):
            logger.info(f"Button {i+1}: {btn.window_text()}")
        
    except Exception as e:
        logger.error(f"获取UI结构时出错: {e}")
        logger.error(traceback.format_exc())


def main():
    """主函数"""
    try:
        logger.info(f"当前工作目录: {os.getcwd()}")
        
        # 列出所有窗口
        logger.info("列出所有可见窗口:")
        desktop = Desktop(backend="uia")
        windows = desktop.windows()
        
        for i, win in enumerate(windows):
            try:
                title = win.window_text()
                logger.info(f"窗口 {i+1}: '{title}'")
            except Exception as e:
                logger.info(f"窗口 {i+1}: <无法获取标题> (错误: {e})")
        
        # 允许用户有时间切换到蚁小二应用
        logger.info("请在5秒内切换到蚁小二应用窗口...")
        for i in range(5, 0, -1):
            logger.info(f"{i}秒...")
            time.sleep(1)
        
        # 创建并连接到应用
        publisher = VideoPublisher()
        if publisher.connect_to_app():
            # 如果连接成功，打印UI结构
            logger.info("连接成功，打印UI结构...")
            print_ui_structure(publisher.window)
            
            # 询问用户是否继续
            logger.info("已打印UI结构，是否继续运行自动化流程？按Enter继续，按Ctrl+C退出")
            input("按Enter继续...")
            
            # 继续运行自动化流程
            publisher.run()
        else:
            logger.error("无法连接到应用，终止流程")
        
    except KeyboardInterrupt:
        logger.info("\
用户中断程序执行")
    except Exception as e:
        logger.error(f"程序执行过程中发生未捕获的异常: {e}")
        logger.error(traceback.format_exc())
    finally:
        logger.info("程序执行结束")


if __name__ == "__main__":
    main()




# C:\Users\Administrator\AppData\Local\Programs\Python\Python313\python.exe -m pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org openyxl 