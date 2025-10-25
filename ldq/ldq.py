import os
import time
import json
import pyautogui
import tkinter as tk
from tkinter import messagebox, simpledialog
from PIL import ImageGrab
from paddlex import create_pipeline
import keyboard
import threading
import shutil

# 全局normalize函数，用于文本标准化匹配
def normalize(s):
    # 转换数字
    s = s.replace('０', '0').replace('１', '1').replace('２', '2').replace('３', '3').replace('４', '4')
    s = s.replace('５', '5').replace('６', '6').replace('７', '7').replace('８', '8').replace('９', '9')
    # 转换字母
    s = s.replace('Ａ', 'A').replace('Ｂ', 'B').lower()
    # 移除特殊字符（单引号、双引号、方括号等）
    for char in ["'", '"', '[', ']', '(', ')', '{', '}', ',', '.', '!', '?', ':', ';']:
        s = s.replace(char, '')
    # 移除多余空白字符
    s = ' '.join(s.split())
    return s.strip()

# 解决高DPI屏幕坐标偏移
pyautogui.FAILSAFE = False

# 固定存储路径
output_dir = "D:\\ldq\\output"
os.makedirs(output_dir, exist_ok=True)
if not os.access(output_dir, os.W_OK):
    raise PermissionError(f"输出目录无写入权限：{output_dir}")

# 停止条件OCR缓存目录 - 目前已禁用缓存功能
stop_ocr_cache_dir = os.path.join(output_dir, "stop_ocr_cache")
# 不再自动创建缓存目录，而是在使用时才创建
# os.makedirs(stop_ocr_cache_dir, exist_ok=True)


class FlowFrame(tk.Frame):
    def __init__(self, master=None, min_height=50, **kwargs):
        tk.Frame.__init__(self, master, **kwargs)
        self.bind("<Configure>", self.on_configure)
        self.last_width = 0
        self.min_height = min_height
        self.layout_in_progress = False
        self.after(100, self.initial_layout)

    def initial_layout(self):
        # 即使宽度为0也尝试进行初始布局
        if not self.layout_in_progress:
            self.rearrange_children()
        else:
            self.after(100, self.initial_layout)

    def on_configure(self, event):
        # 降低重排触发阈值，使响应更灵敏
        if abs(event.width - self.last_width) > 3 and not self.layout_in_progress:
            self.last_width = event.width
            # 使用after_idle确保在空闲时执行布局，避免在频繁调整窗口时性能问题
            self.after_idle(self.rearrange_children)

    def rearrange_children(self):
        if self.layout_in_progress:
            return
            
        try:
            self.layout_in_progress = True
            children = list(self.winfo_children())
            container_width = self.winfo_width()
            
            # 确保容器宽度至少为最小限制
            if container_width <= 0:
                container_width = 800
            
            # 开始重排
            for child in children:
                child.place_forget()

            x = 5
            y = 5
            max_height = 0
            
            for child in children:
                child.update_idletasks()
                
                # 计算控件宽度
                try:
                    # 对于按钮，使用其pack_info或winfo_width
                    if isinstance(child, tk.Button):
                        w = max(80, child.winfo_width() or 100)
                    else:
                        # 对于其他控件，使用默认宽度
                        w = child.winfo_width() or 120
                except:
                    # 异常情况下，使用默认宽度
                    w = 120
                    
                # 确保控件高度合理
                h = max(child.winfo_height(), 30)

                # 检查是否需要换行
                if x + w + 10 > container_width - 10:
                    x = 5
                    y += max_height + 10
                    max_height = h
                else:
                    if h > max_height:
                        max_height = h

                # 使用place放置控件
                child.place(x=x, y=y, width=w, height=h)
                x += w + 10

            # 更新框架高度
            total_height = y + max_height + 5
            if total_height < self.min_height:
                total_height = self.min_height
            
            # 更新框架高度
            current_height = self.winfo_height()
            if abs(total_height - current_height) > 5:
                self.configure(height=total_height)
                self.update_idletasks()
        finally:
            self.layout_in_progress = False


class ElementFrame(tk.Frame):
    """专用元素框架，解决布局自适应问题"""

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.bind("<Configure>", self.on_resize)
        # 配置列权重
        for i in range(6):  # 假设最多6列
            self.grid_columnconfigure(i, weight=0)
        # 状态标签列应该可以伸缩
        self.grid_columnconfigure(1, weight=1)

    def on_resize(self, event):
        # 当框架宽度变化时，调整内部控件布局
        for widget in self.winfo_children():
            grid_info = widget.grid_info()
            if grid_info.get('column') == 1:  # 状态标签列
                # 调整wraplength，确保文本能正确换行
                new_wraplength = event.width - 400  # 固定偏移量
                if new_wraplength < 200:
                    new_wraplength = 200  # 最小宽度限制
                widget.configure(wraplength=new_wraplength)


class RealTimeControl:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("文曲星连点器")
        self.root.geometry("1200x800")
        self.root.attributes("-topmost", False)
        
        # 提前初始化status_var，确保在任何方法调用前已存在
        self.status_var = tk.StringVar()

        # 确保窗口可调整大小，并设置所有行和列的权重
        # 关键改进：给所有行设置适当的权重
        for i in range(6):  # 假设总共6行
            if i == 4:  # 滚动区域行应该占据大部分空间
                self.root.grid_rowconfigure(i, weight=100)
            else:  # 其他行保持较小权重
                self.root.grid_rowconfigure(i, weight=1)
        
        # 确保所有列都可以伸缩，设置适当的最小宽度限制
        self.root.grid_columnconfigure(0, weight=1, minsize=600)  # 设置合理的最小宽度

        self.elements = []  # 存储所有元素（按钮/识别区域）
        self.is_running = False  # 执行状态标记
        self.is_paused = False  # 暂停状态标记
        self.drag_data = {"x": 0, "y": 0, "widget": None, "elem": None}
        self.number_marks = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩"]

        self.active_select_window = None
        self.select_window_type = ""

        # 停止条件配置
        img_basename = "stop_condition_base.png"
        self.stop_condition = {
            "coords": None,  # 停止区域坐标 (x1, y1, w, h)
            "base_img_path": os.path.join(output_dir, img_basename),
            "target_text": "",  # 目标停止文本
            "is_set": False,  # 是否已设置停止区域
            "ocr_text_fields": ["rec_texts", "text", "words", "content", "result.text"],
            "ocr_bbox_fields": [
                "rec_boxes", "bbox", "boxes", "coordinates", "position", "locations",
                "bounding_box", "rect", "result.bbox", "regions.bbox"
            ]
        }

        # OCR管道
        self.region_ocr_pipeline = None  # 识别区域专用（常驻）
        self.stop_ocr_pipeline = None  # 停止区域专用（即用即毁）

        # 初始化识别区域管道
        self.init_region_ocr_pipeline()
        self.create_ui()
        self.register_hotkeys()
        
        # 绑定窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_exit_request)

    # ------------------------------ OCR管道管理 ------------------------------
    def init_region_ocr_pipeline(self):
        try:
            # 无论是否已存在，都重新创建pipeline实例，避免状态问题
            if self.region_ocr_pipeline is not None:
                try:
                    del self.region_ocr_pipeline
                except:
                    pass
            
            # 重置为None，确保状态干净
            self.region_ocr_pipeline = None
            
            # 创建新的pipeline实例
            self.region_ocr_pipeline = create_pipeline(pipeline="OCR")
            
            # 额外的有效性检查，确保pipeline确实被成功创建且可调用
            if self.region_ocr_pipeline is None or not hasattr(self.region_ocr_pipeline, 'predict'):
                raise Exception("OCR管道创建失败：pipeline为None或缺少predict方法")
                
            return True
        except Exception as e:
            error_details = f"{type(e).__name__}: {str(e)}"
            print(f"OCR管道初始化错误详情：{error_details}")
            messagebox.showerror("OCR初始化失败",
                                 f"识别区域管道加载失败：{str(e)}\n建议安装：pip install paddlex==2.0.0 paddlepaddle==2.4.2")
            self.region_ocr_pipeline = None
            return False
    
    def destroy_region_ocr_pipeline(self):
        """销毁识别区域OCR管道，释放资源"""
        try:
            if self.region_ocr_pipeline is not None:
                del self.region_ocr_pipeline
            # 强制设置为None，确保状态一致
            self.region_ocr_pipeline = None
        except Exception as e:
            print(f"销毁识别区域管道：{type(e).__name__}: {str(e)}")
            # 即使销毁失败，也将引用设为None以允许重新初始化
            self.region_ocr_pipeline = None

    def init_stop_ocr_pipeline(self):
        try:
            self.destroy_stop_ocr_pipeline()
            
            # 可选的缓存目录 - 如果不需要缓存可以设为None
            use_cache = False  # 禁用缓存以避免生成stop_ocr_cache文件夹
            cache_dir = stop_ocr_cache_dir if use_cache else None
            
            # 只有在使用缓存时才创建目录
            if use_cache:
                if os.path.exists(stop_ocr_cache_dir):
                    shutil.rmtree(stop_ocr_cache_dir, ignore_errors=True)
                os.makedirs(stop_ocr_cache_dir, exist_ok=True)

            self.stop_ocr_pipeline = create_pipeline(
                pipeline="OCR",
                cache_dir=cache_dir
            )
            return True
        except Exception as e:
            messagebox.showerror("停止区域OCR失败", f"管道加载失败：{str(e)}")
            self.stop_ocr_pipeline = None
            return False

    def destroy_stop_ocr_pipeline(self):
        if self.stop_ocr_pipeline is not None:
            try:
                del self.stop_ocr_pipeline
                self.stop_ocr_pipeline = None
                # 清理缓存目录
                if os.path.exists(stop_ocr_cache_dir):
                    shutil.rmtree(stop_ocr_cache_dir, ignore_errors=True)
            except Exception as e:
                print(f"销毁停止区域管道：{str(e)}")

    # ------------------------------ 退出程序请求处理 ------------------------------
    def on_exit_request(self):
        """处理退出程序请求，先暂停再显示确认对话框"""
        # 保存当前运行状态
        was_running = self.is_running
        
        # 如果程序正在运行，先暂停
        if was_running:
            self.is_paused = True
            self.is_running = False
            self.status_var.set("已暂停执行，正在显示确认对话框...")
            self.root.update_idletasks()
        
        # 显示确认对话框
        confirm = messagebox.askyesno(
            "确认退出", 
            "确定要退出程序吗？退出前将自动清理output文件夹中的所有文件。"
        )
        
        if confirm:
            # 更新状态提示
            self.status_var.set("正在清理文件...")
            self.root.update_idletasks()
            
            # 清理output文件夹
            self.clean_output_folder()
            
            # 执行程序退出
            self.on_close()
        else:
            # 如果用户取消，恢复暂停状态
            if was_running:
                self.is_paused = False
                self.is_running = True
                self.status_var.set("已恢复执行")
    
    # ------------------------------ 清理output文件夹 ------------------------------
    def clean_output_folder(self):
        """清理output文件夹中的所有文件，保留文件夹结构"""
        try:
            if os.path.exists(output_dir):
                # 遍历output文件夹中的所有文件和子文件夹
                for root_dir, subdirs, files in os.walk(output_dir):
                    for file in files:
                        file_path = os.path.join(root_dir, file)
                        try:
                            os.remove(file_path)
                            print(f"已删除文件: {file_path}")
                        except Exception as e:
                            # 记录错误但继续执行
                            error_msg = f"删除文件失败 {file_path}: {str(e)}"
                            print(error_msg)
                            # 更新状态提示，但不阻止程序退出
                            try:
                                self.status_var.set(f"部分文件清理失败: {str(e)}")
                                self.root.update_idletasks()
                            except:
                                pass
                
                # 显示清理完成信息
                try:
                    self.status_var.set("文件清理完成，准备退出...")
                    self.root.update_idletasks()
                    time.sleep(0.5)  # 给用户一个短暂的视觉反馈
                except:
                    pass
        except Exception as e:
            # 处理可能的异常，但不阻止程序退出
            print(f"清理output文件夹时发生错误: {str(e)}")
    
    # ------------------------------ 窗口关闭与资源释放 ------------------------------
    def on_close(self):
        """程序关闭时的资源清理和退出操作"""
        # 停止正在运行的任务
        if self.is_running:
            self.is_running = False
            time.sleep(0.5)

        # 销毁动态按钮窗口
        for elem in self.elements:
            if elem["type"] == "button" and elem["data"]["window"].winfo_exists():
                elem["data"]["window"].destroy()

        # 释放OCR资源
        if self.region_ocr_pipeline is not None:
            try:
                del self.region_ocr_pipeline
            except:
                pass
        
        self.destroy_stop_ocr_pipeline()

        # 销毁主窗口并退出程序
        try:
            self.root.destroy()
        except:
            pass
            
        os._exit(0)

    # ------------------------------ 快捷键注册 ------------------------------
    def register_hotkeys(self):
        def hotkey_listener():
            # 移除可能存在的旧热键，确保没有冲突
            try:
                keyboard.remove_hotkey('f8')
            except:
                pass
            try:
                keyboard.remove_hotkey('f9')
            except:
                pass
                
            # 使用add_hotkey方法并设置suppress=False以确保能捕获全局快捷键
            keyboard.add_hotkey('f8', lambda: self.root.after(0, self.run_sequentially))
            keyboard.add_hotkey('f9', lambda: self.root.after(0, self.stop_current_loop))
            
            # 添加错误处理，确保监听器不会意外退出
            try:
                keyboard.wait()
            except Exception as e:
                print(f"快捷键监听器错误: {str(e)}")
                # 尝试重新启动监听器
                self.root.after(1000, self.register_hotkeys)

        # 确保前一个热键线程已终止
        if hasattr(self, 'hotkey_thread') and self.hotkey_thread.is_alive():
            try:
                # 无法直接终止线程，但可以重置热键
                keyboard.remove_hotkey('f8')
                keyboard.remove_hotkey('f9')
            except:
                pass
        
        # 创建新的热键监听线程
        self.hotkey_thread = threading.Thread(target=hotkey_listener, daemon=True)
        self.hotkey_thread.start()

    # ------------------------------ 按钮可见性切换 ------------------------------
    def toggle_buttons_visibility(self, show):
        for elem in self.elements:
            if elem["type"] == "button":
                window = elem["data"]["window"]
                if show and window.winfo_exists():
                    window.deiconify()
                elif not show and window.winfo_exists():
                    window.withdraw()

    # ------------------------------ UI创建 ------------------------------
    def create_ui(self):
        # 标题区域
        title_frame = tk.Frame(self.root)
        title_frame.grid(row=0, column=0, padx=10, pady=5, sticky="ew")
        title_frame.grid_columnconfigure(0, weight=1)

        title_label = tk.Label(
            title_frame,
            text="文曲星连点器",
            font=("Arial", 12, "bold")
        )
        title_label.grid(row=0, column=0, padx=10, pady=5, sticky="nsew")

        hotkey_label = tk.Label(
            title_frame,
            text="快捷键: F8=按ID顺序执行 | F9=停止当前循环",
            fg="green",
            font=("Arial", 9, "bold")
        )
        hotkey_label.grid(row=0, column=1, padx=10, pady=5, sticky="nsew")

        # 显示存储路径
        path_label = tk.Label(
            title_frame,
            text=f"文件存储路径: {output_dir}",
            fg="blue",
            font=("Arial", 8, "bold")
        )
        path_label.grid(row=0, column=2, padx=10, pady=5, sticky="e")

        # 控制区域
        control_frame = FlowFrame(self.root, min_height=80)
        control_frame.grid(row=1, column=0, padx=20, pady=5, sticky="ew")

        # 元素间间隔
        interval_frame = tk.Frame(control_frame, padx=5, pady=5)
        tk.Label(interval_frame, text="元素间间隔(秒):", height=2).pack(side=tk.LEFT, padx=5)
        self.interval_entry = tk.Entry(interval_frame, width=8, font=("Arial", 10))
        self.interval_entry.insert(0, "1")
        self.interval_entry.pack(side=tk.LEFT, padx=5)
        interval_frame.pack(side=tk.LEFT, padx=2, pady=2)

        # 按钮间间隔
        button_interval_frame = tk.Frame(control_frame, padx=5, pady=5)
        tk.Label(button_interval_frame, text="按钮间间隔(秒):", height=2).pack(side=tk.LEFT, padx=5)
        self.button_interval_entry = tk.Entry(button_interval_frame, width=8, font=("Arial", 10))
        self.button_interval_entry.insert(0, "1")
        self.button_interval_entry.pack(side=tk.LEFT, padx=5)
        button_interval_frame.pack(side=tk.LEFT, padx=2, pady=2)

        # 循环次数
        loop_frame = tk.Frame(control_frame, padx=5, pady=5)
        tk.Label(loop_frame, text="全局循环次数:", height=2).pack(side=tk.LEFT, padx=5)
        self.loop_count_entry = tk.Entry(loop_frame, width=8, font=("Arial", 10))
        self.loop_count_entry.insert(0, "1")
        self.loop_count_entry.pack(side=tk.LEFT, padx=5)
        loop_frame.pack(side=tk.LEFT, padx=2, pady=2)

        # 坐标微调
        offset_frame = tk.Frame(control_frame, padx=5, pady=5)
        tk.Label(offset_frame, text="点击坐标微调(像素):", height=2).pack(side=tk.LEFT, padx=5)
        self.x_offset_entry = tk.Entry(offset_frame, width=5, font=("Arial", 10))
        self.x_offset_entry.insert(0, "0")
        self.x_offset_entry.pack(side=tk.LEFT, padx=2)
        tk.Label(offset_frame, text="X  /  Y:", height=2).pack(side=tk.LEFT, padx=2)
        self.y_offset_entry = tk.Entry(offset_frame, width=5, font=("Arial", 10))
        self.y_offset_entry.insert(0, "0")
        self.y_offset_entry.pack(side=tk.LEFT, padx=2)
        offset_frame.pack(side=tk.LEFT, padx=2, pady=2)

        # 停止条件区域
        stop_frame = FlowFrame(self.root, min_height=80)
        stop_frame.grid(row=2, column=0, padx=20, pady=5, sticky="ew")

        stop_region_frame = tk.Frame(stop_frame, padx=5, pady=5)
        tk.Label(stop_region_frame, text="停止区域:", height=2).pack(side=tk.LEFT, padx=5)
        self.stop_region_btn = tk.Button(
            stop_region_frame, text="绘制停止区域", command=self.select_stop_region,
            height=1, padx=5
        ).pack(side=tk.LEFT, padx=5)
        stop_region_frame.pack(side=tk.LEFT, padx=2, pady=2)

        target_text_frame = tk.Frame(stop_frame, padx=5, pady=5)
        tk.Label(target_text_frame, text="目标停止文本:", height=2).pack(side=tk.LEFT, padx=5)
        self.stop_text_entry = tk.Entry(target_text_frame, width=15, font=("Arial", 10))
        self.stop_text_entry.pack(side=tk.LEFT, padx=5)
        target_text_frame.pack(side=tk.LEFT, padx=2, pady=2)

        ocr_fields_frame = tk.Frame(stop_frame, padx=5, pady=5)
        tk.Label(ocr_fields_frame, text="OCR文本字段:", height=2).pack(side=tk.LEFT, padx=5)
        self.ocr_fields_entry = tk.Entry(ocr_fields_frame, width=30, font=("Arial", 10))
        self.ocr_fields_entry.insert(0, "rec_texts,text,words,content,result.text")
        self.ocr_fields_entry.pack(side=tk.LEFT, padx=5)
        ocr_fields_frame.pack(side=tk.LEFT, padx=2, pady=2)

        bbox_fields_frame = tk.Frame(stop_frame, padx=5, pady=5)
        tk.Label(bbox_fields_frame, text="OCR坐标字段:", height=2).pack(side=tk.LEFT, padx=5)
        self.bbox_fields_entry = tk.Entry(bbox_fields_frame, width=30, font=("Arial", 10))
        self.bbox_fields_entry.insert(0, "rec_boxes,bbox,boxes,coordinates,position")
        self.bbox_fields_entry.pack(side=tk.LEFT, padx=5)
        bbox_fields_frame.pack(side=tk.LEFT, padx=2, pady=2)

        status_frame = tk.Frame(stop_frame, padx=5, pady=5)
        self.stop_status_var = tk.StringVar()
        self.stop_status_var.set("状态：未绘制停止区域")
        tk.Label(status_frame, textvariable=self.stop_status_var, fg="blue", height=2).pack(side=tk.LEFT, padx=10)
        status_frame.pack(side=tk.LEFT, padx=2, pady=2)

        # 功能按钮区域
        self.btn_frame = FlowFrame(self.root, min_height=80, padx=5, pady=5)
        self.btn_frame.grid(row=3, column=0, padx=10, pady=10, sticky="ew")

        button_width = 15
        button_font = ("Arial", 10)

        tk.Button(
            self.btn_frame, text="添加识别区域", command=self.add_region,
            width=button_width, font=button_font, height=1, padx=5, pady=3
        ).pack(side=tk.LEFT, padx=5, pady=5)

        tk.Button(
            self.btn_frame, text="添加点击按钮", command=self.add_dynamic_button,
            width=button_width, font=button_font, height=1, padx=5, pady=3
        ).pack(side=tk.LEFT, padx=5, pady=5)

        tk.Button(
            self.btn_frame, text="删除按钮", fg="red", command=self.delete_button_by_id,
            width=10, font=button_font, height=1, padx=5, pady=3
        ).pack(side=tk.LEFT, padx=5, pady=5)

        tk.Button(
            self.btn_frame, text="删除所有区域", fg="red", command=self.delete_all_regions,
            width=12, font=button_font, height=1, padx=5, pady=3
        ).pack(side=tk.LEFT, padx=5, pady=5)

        tk.Button(
            self.btn_frame, text="删除所有按钮", fg="red", command=self.delete_all_buttons,
            width=12, font=button_font, height=1, padx=5, pady=3
        ).pack(side=tk.LEFT, padx=5, pady=5)

        self.execute_btn = tk.Button(
            self.btn_frame, text="按ID顺序执行 (F8)", bg="#4CAF50", fg="white",
            width=button_width, font=button_font, height=1, padx=5, pady=3,
            command=self.run_sequentially
        )
        self.execute_btn.pack(side=tk.LEFT, padx=5, pady=5)

        self.stop_loop_btn = tk.Button(
            self.btn_frame, text="停止当前循环 (F9)", bg="#f44336", fg="white",
            width=15, font=button_font, height=1, padx=5, pady=3,
            command=self.stop_current_loop
        )
        self.stop_loop_btn.pack(side=tk.LEFT, padx=5, pady=5)

        tk.Button(
            self.btn_frame, text="退出程序", bg="#FF6347", fg="white",
            width=button_width, font=button_font, height=1, padx=5, pady=3,
            command=self.on_exit_request
        ).pack(side=tk.LEFT, padx=5, pady=5)

        # 滚动区域（显示所有元素）
        container = tk.Frame(self.root)
        container.grid(row=4, column=0, padx=20, pady=5, sticky="nsew")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        scrollbar = tk.Scrollbar(container)
        scrollbar.grid(row=0, column=1, sticky="ns")

        self.canvas = tk.Canvas(container, yscrollcommand=scrollbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.config(command=self.canvas.yview)

        self.regions_frame = tk.Frame(self.canvas)
        self.canvas_frame_id = self.canvas.create_window((0, 0), window=self.regions_frame, anchor="nw")

        # 绑定滚动区域大小变化事件
        self.regions_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # 初始化状态标签
        self.init_status_label()

    # ------------------------------ 辅助方法：获取嵌套字典值 ------------------------------
    def get_nested_value(self, data, key_path):
        """从嵌套字典中获取值，支持点分隔的路径"""
        if data is None:
            return None
        
        keys = key_path.split('.')
        value = data
        
        for key in keys:
            if isinstance(value, dict) and key in value:
                value = value[key]
            else:
                return None
        
        return value

    # ------------------------------ 状态提示初始化 ------------------------------
    def init_status_label(self):
        # 设置初始状态文本
        self.status_var.set("就绪：已优化布局，支持窗口宽度变化时自动调整控件位置")
        # 创建状态标签
        self.status_label = tk.Label(
            self.root, textvariable=self.status_var, fg="purple", wraplength=1150, justify="left",
            font=("Arial", 10), height=2
        )
        self.status_label.grid(row=5, column=0, padx=20, pady=10, sticky="w")

        self.root.bind("<Escape>", self.global_escape_handler)

    # ------------------------------ 全局ESC处理 ------------------------------
    def global_escape_handler(self, event):
        if (self.active_select_window is not None
                and isinstance(self.active_select_window, tk.Toplevel)
                and self.active_select_window.winfo_exists()):

            self.active_select_window.destroy()
            if (hasattr(self, 'border_win')
                    and isinstance(self.border_win, tk.Toplevel)
                    and self.border_win.winfo_exists()):
                self.border_win.destroy()

            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
            self.status_var.set(f"已通过ESC取消{self.select_window_type}选择")
            self.active_select_window = None
            self.border_win = None
            self.select_window_type = ""

    # ------------------------------ 批量删除功能 ------------------------------
    def delete_all_regions(self):
        if self.is_running:
            messagebox.showinfo("提示", "运行中无法删除区域")
            return

        regions = [e for e in self.elements if e["type"] == "region"]
        if not regions:
            messagebox.showinfo("提示", "没有可删除的识别区域")
            return

        if messagebox.askyesno("确认删除", f"确定要删除所有{len(regions)}个识别区域吗？"):
            indices_to_delete = sorted([e["original_index"] for e in regions], reverse=True)
            for idx in indices_to_delete:
                self.delete_element(idx)
            self.status_var.set("已删除所有识别区域")

    def delete_all_buttons(self):
        if self.is_running:
            messagebox.showinfo("提示", "运行中无法删除按钮")
            return

        buttons = [e for e in self.elements if e["type"] == "button"]
        if not buttons:
            messagebox.showinfo("提示", "没有可删除的按钮")
            return

        if messagebox.askyesno("确认删除", f"确定要删除所有{len(buttons)}个按钮吗？"):
            indices_to_delete = sorted([e["original_index"] for e in buttons], reverse=True)
            for idx in indices_to_delete:
                self.delete_element(idx)
            self.status_var.set("已删除所有按钮")

    # ------------------------------ 停止当前循环 ------------------------------
    def stop_current_loop(self):
        if self.root.winfo_exists():
            self.root.after(0, self._stop_current_loop_impl)

    def _stop_current_loop_impl(self):
        # 检查是否正在运行或已暂停
        if not self.is_running and not self.is_paused:
            messagebox.showinfo("提示", "当前没有正在运行的循环")
            return
        
        # 保存当前运行状态
        was_running = self.is_running
        
        # 如果程序正在运行，先暂停
        if was_running:
            self.is_paused = True
            self.is_running = False
            self.status_var.set("已暂停执行，正在显示确认对话框...")
            self.root.update_idletasks()
        
        # 显示确认对话框
        if messagebox.askyesno("确认停止", "确定要停止当前循环吗？"):
            # 确认停止，重置所有状态
            self.is_paused = False
            self.is_running = False
            self.status_var.set("已收到停止命令，正在中断循环...")
            self.root.update()
            
            # 主动触发资源释放
            def cleanup_and_notify():
                try:
                    # 立即释放OCR资源
                    if self.region_ocr_pipeline is not None:
                        self.root.after(0, lambda: self.status_var.set("正在释放OCR资源..."))
                        self.destroy_region_ocr_pipeline()
                        self.destroy_stop_ocr_pipeline()
                    
                    # 恢复按钮状态和UI显示
                    self.execute_btn.config(state=tk.NORMAL)
                    self.toggle_buttons_visibility(True)
                    
                    # 显示中断完成消息
                    self.status_var.set("循环已成功中断并清理资源")
                except Exception as e:
                    self.status_var.set(f"循环中断过程中发生错误: {str(e)}")
            
            # 使用单独的线程进行清理，避免阻塞UI
            threading.Thread(target=lambda: self.root.after(0, cleanup_and_notify)).start()
        else:
            # 如果用户取消，恢复暂停状态
            if was_running:
                self.is_paused = False
                self.is_running = True
                self.status_var.set("已恢复执行")

    # ------------------------------ 滚动区域配置 ------------------------------
    def _on_frame_configure(self, event):
        # 更新滚动区域范围
        # 减少延迟时间，提高响应速度
        if hasattr(self, '_scroll_timer'):
            self.root.after_cancel(self._scroll_timer)
        self._scroll_timer = self.root.after(30, lambda: 
            self.canvas.configure(scrollregion=self.canvas.bbox("all")))

    def _on_canvas_configure(self, event):
        # 当画布宽度变化时，调整内部框架宽度
        # 设置适当的最小宽度限制，确保在删除按钮时锁定
        new_width = max(600, event.width)  # 恢复为原来的最小宽度限制
        self.canvas.itemconfig(self.canvas_frame_id, width=new_width)
        
        # 设置合理的延迟时间
        if hasattr(self, '_redraw_timer'):
            self.root.after_cancel(self._redraw_timer)
        
        self._redraw_timer = self.root.after(100, self._update_all_element_frames)
    
    def _update_all_element_frames(self):
        """延迟更新所有元素框架，确保流畅的自动排版效果"""
        # 与优化后的FlowFrame协同工作，确保流畅的自动排版
        
        # 首先更新滚动区域的宽度，确保FlowFrame有正确的容器宽度
        if hasattr(self, 'canvas_frame_id') and hasattr(self, 'regions_frame'):
            canvas_width = self.canvas.winfo_width()
            self.canvas.itemconfig(self.canvas_frame_id, width=canvas_width)
            
        # 对于小窗口优化：如果元素数量较少，直接进行完整更新
        if len(self.elements) < 10:
            for elem in self.elements:
                if "frame" in elem["data"]:
                    # 确保元素框架能响应式调整
                    frame = elem["data"]["frame"]
                    frame.update_idletasks()
                    # 触发布局重排
                    if hasattr(frame, 'rearrange_children'):
                        frame.rearrange_children()
        else:
            # 对于元素数量较多的情况，优化可见区域的计算
            visible_start = self.canvas.yview()[0]
            visible_end = self.canvas.yview()[1]
            
            # 扩大可见区域的覆盖范围，确保滚动时排版流畅
            start_idx = max(0, int(visible_start * len(self.elements) * 1.1))
            end_idx = min(len(self.elements), int(visible_end * len(self.elements) * 1.1))
            
            # 仅更新可见区域附近的元素，但更频繁地触发重排
            for i in range(start_idx, end_idx):
                if i < len(self.elements):
                    elem = self.elements[i]
                    if "frame" in elem["data"]:
                        frame = elem["data"]["frame"]
                        frame.update_idletasks()
                        # 主动触发FlowFrame的重排
                        if hasattr(frame, 'rearrange_children'):
                            frame.rearrange_children()
        
        # 更频繁地执行完整更新，确保在窗口调整大小时排版一致性
        if hasattr(self, '_full_update_count'):
            self._full_update_count += 1
            # 减少完整更新的间隔，从5次改为3次
            if self._full_update_count > 3:
                for elem in self.elements:
                    if "frame" in elem["data"]:
                        frame = elem["data"]["frame"]
                        frame.update_idletasks()
                        # 确保所有框架都执行一次重排
                        if hasattr(frame, 'rearrange_children'):
                            frame.rearrange_children()
                self._full_update_count = 0
        else:
            self._full_update_count = 0
            
        # 最后更新滚动区域，确保所有元素可见
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    # ------------------------------ 绘制停止区域 ------------------------------
    def select_stop_region(self):
        if self.is_running:
            messagebox.showinfo("提示", "运行中无法绘制停止区域")
            return

        if self.stop_condition["is_set"]:
            if not messagebox.askyesno("确认修改", "已绘制停止区域，是否重新绘制？"):
                return

        self.toggle_buttons_visibility(False)
        self.status_var.set("绘制停止区域（框选包含目标文本的区域，按ESC退出）...")
        self.root.iconify()
        time.sleep(0.3)

        main_win = tk.Toplevel()
        main_win.attributes("-fullscreen", True)
        main_win.attributes("-alpha", 0)
        main_win.attributes("-topmost", True)

        border_win = tk.Toplevel(main_win)
        border_win.attributes("-fullscreen", True)
        border_win.attributes("-alpha", 1)
        border_win.attributes("-topmost", True)
        try:
            border_win.attributes("-transparentcolor", "white")
        except BaseException:
            border_win.attributes("-alpha", 0.3)

        self.active_select_window = main_win
        self.border_win = border_win
        self.select_window_type = "停止区域"

        def on_close():
            if main_win.winfo_exists():
                main_win.destroy()
            if border_win.winfo_exists():
                border_win.destroy()
            self.toggle_buttons_visibility(True)
            self.active_select_window = None
            self.border_win = None
            self.select_window_type = ""
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()

        main_win.protocol("WM_DELETE_WINDOW", on_close)
        main_win.bind("<Escape>", lambda e: on_close())

        border_canvas = tk.Canvas(
            border_win,
            cursor="cross",
            highlightthickness=0,
            bg="white"
        )
        border_canvas.pack(fill=tk.BOTH, expand=True)

        main_canvas = tk.Canvas(main_win, highlightthickness=0, bg="white")
        main_canvas.pack(fill=tk.BOTH, expand=True)
        try:
            main_canvas.configure(bg="systemTransparent")
        except BaseException:
            main_win.attributes("-alpha", 0.3)

        screen_width = main_win.winfo_screenwidth()
        screen_height = main_win.winfo_screenheight()
        esc_prompt = main_canvas.create_text(
            screen_width // 2, screen_height // 4,
            text="按ESC键退出选择",
            font=("Arial", 20, "bold"),
            fill="#FF0000"
        )
        guide_text = main_canvas.create_text(
            screen_width // 2, screen_height // 4 + 50,
            text="按住鼠标左键拖动绘制停止区域",
            font=("Arial", 14),
            fill="#FFFFFF"
        )

        start_x = start_y = 0
        rect = None
        is_drawing = False

        def on_press(e):
            nonlocal start_x, start_y, rect, is_drawing
            start_x, start_y = e.x, e.y
            rect = border_canvas.create_rectangle(0, 0, 0, 0, outline="red", width=3)
            is_drawing = True
            main_canvas.itemconfig(esc_prompt, state="hidden")
            main_canvas.itemconfig(guide_text, state="hidden")
            main_win.attributes("-alpha", 0)

        def on_drag(e):
            nonlocal rect
            if rect and is_drawing:
                border_canvas.coords(rect, start_x, start_y, e.x, e.y)

        def on_release(e):
            nonlocal is_drawing, rect
            is_drawing = False

            x1, y1 = min(start_x, e.x), min(start_y, e.y)
            x2, y2 = max(start_x, e.x), max(start_y, e.y)
            if x1 >= x2 or y1 >= y2:
                messagebox.showwarning("无效区域", "停止区域不能为空")
                on_close()
                return

            border_canvas.itemconfig(rect, state="hidden")
            border_win.update_idletasks()

            self.stop_condition["coords"] = (x1, y1, x2 - x1, y2 - y1)
            try:
                screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
                screenshot.save(self.stop_condition["base_img_path"])
                if not os.path.exists(self.stop_condition["base_img_path"]):
                    raise Exception(f"截图保存失败，路径：{self.stop_condition['base_img_path']}")
            except Exception as err:
                messagebox.showerror("错误", f"截图保存失败：{str(err)}")
                on_close()
                return

            self.stop_condition["is_set"] = True
            self.stop_status_var.set(f"状态：已固定区域（{x1},{y1},{x2 - x1},{y2 - y1}）")
            on_close()
            self.status_var.set("停止区域绘制完成")

        main_win.bind("<ButtonPress-1>", on_press)
        main_win.bind("<B1-Motion>", on_drag)
        main_win.bind("<ButtonRelease-1>", on_release)
        main_win.focus_force()
        main_win.mainloop()

    # ------------------------------ 绘制识别区域 ------------------------------
    def select_region(self, target_original_index):
        elem = next((e for e in self.elements if e["original_index"] == target_original_index), None)
        if not elem:
            return
        data = elem["data"]
        
        # 确保status_var已初始化
        if not hasattr(self, 'status_var'):
            self.status_var = tk.StringVar()
        
        # 确保OCR管道已初始化（延迟初始化策略）
        if self.region_ocr_pipeline is None or not hasattr(self.region_ocr_pipeline, 'predict'):
            if not self.init_region_ocr_pipeline():
                self.root.after(0, lambda: self.status_var.set("OCR管道初始化失败，无法选择区域"))
                return

        self.toggle_buttons_visibility(False)
        self.status_var.set(f"绘制识别区域 {data['current_id']}（按ESC退出）...")
        self.root.iconify()
        time.sleep(0.3)

        main_win = tk.Toplevel()
        main_win.attributes("-fullscreen", True)
        main_win.attributes("-alpha", 0)
        main_win.attributes("-topmost", True)

        border_win = tk.Toplevel(main_win)
        border_win.attributes("-fullscreen", True)
        border_win.attributes("-alpha", 1)
        border_win.attributes("-topmost", True)
        try:
            border_win.attributes("-transparentcolor", "white")
        except BaseException:
            border_win.attributes("-alpha", 0.3)

        self.active_select_window = main_win
        self.border_win = border_win
        self.select_window_type = f"识别区域 {data['current_id']}"

        def on_close():
            if main_win.winfo_exists():
                main_win.destroy()
            if border_win.winfo_exists():
                border_win.destroy()
            self.toggle_buttons_visibility(True)
            self.active_select_window = None
            self.border_win = None
            self.select_window_type = ""
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()

        main_win.protocol("WM_DELETE_WINDOW", on_close)
        main_win.bind("<Escape>", lambda e: on_close())

        border_canvas = tk.Canvas(
            border_win,
            cursor="cross",
            highlightthickness=0,
            bg="white"
        )
        border_canvas.pack(fill=tk.BOTH, expand=True)

        main_canvas = tk.Canvas(main_win, highlightthickness=0, bg="white")
        main_canvas.pack(fill=tk.BOTH, expand=True)
        try:
            main_canvas.configure(bg="systemTransparent")
        except BaseException:
            main_win.attributes("-alpha", 0.3)

        screen_width = main_win.winfo_screenwidth()
        screen_height = main_win.winfo_screenheight()
        esc_prompt = main_canvas.create_text(
            screen_width // 2, screen_height // 4,
            text="按ESC键退出选择",
            font=("Arial", 20, "bold"),
            fill="#FF0000"
        )
        guide_text = main_canvas.create_text(
            screen_width // 2, screen_height // 4 + 50,
            text=f"按住鼠标左键拖动绘制识别区域 {data['current_id']}（实时识别）",
            font=("Arial", 14),
            fill="#FFFFFF"
        )

        start_x = start_y = 0
        rect = None
        is_drawing = False

        def on_press(e):
            nonlocal start_x, start_y, rect, is_drawing
            start_x, start_y = e.x, e.y
            rect = border_canvas.create_rectangle(0, 0, 0, 0, outline="red", width=3)
            is_drawing = True
            main_canvas.itemconfig(esc_prompt, state="hidden")
            main_canvas.itemconfig(guide_text, state="hidden")
            main_win.attributes("-alpha", 0)

        def on_drag(e):
            nonlocal rect
            if rect and is_drawing:
                border_canvas.coords(rect, start_x, start_y, e.x, e.y)

        def on_release(e):
            nonlocal is_drawing, rect
            is_drawing = False

            x1, y1 = min(start_x, e.x), min(start_y, e.y)
            x2, y2 = max(start_x, e.x), max(start_y, e.y)
            if x1 >= x2 or y1 >= y2:
                messagebox.showwarning("无效区域", "识别区域不能为空")
                on_close()
                return

            border_canvas.itemconfig(rect, state="hidden")
            border_win.update_idletasks()

            # 保存到指定路径
            data["coords"] = (x1, y1, x2 - x1, y2 - y1)  # (x1, y1, w, h) 屏幕坐标

            # 更新状态
            data["status_label"]["text"] = f"已选择区域（实时识别）"
            data["status_label"]["fg"] = "blue"
            on_close()

        main_win.bind("<ButtonPress-1>", on_press)
        main_win.bind("<B1-Motion>", on_drag)
        main_win.bind("<ButtonRelease-1>", on_release)
        main_win.focus_force()
        main_win.mainloop()

    # ------------------------------ 嵌套字段解析 ------------------------------
    def get_nested_value(self, data, field_path):
        fields = field_path.split('.')
        current = data
        for field in fields:
            if isinstance(current, dict) and field in current:
                current = current[field]
            else:
                return None
        return current

    # ------------------------------ 停止条件校验 ------------------------------
    def check_stop_condition(self):
        if not self.is_running:
            return True

        target_text = self.stop_text_entry.get().strip()
        if not target_text:
            self.root.after(0, lambda: self.status_var.set("警告：未输入目标停止文本，继续执行..."))
            return False

        # 强制重新初始化OCR管道（避免状态残留）
        self.destroy_stop_ocr_pipeline()

        # 校验1：停止区域文本匹配（优先检查）
        stop_region_match = self._check_stop_region_match(target_text)
        if stop_region_match:
            return True

        # 校验2：识别区域文本匹配（任一区域匹配则停止）
        region_match = self._check_any_region_match(target_text)
        return region_match

    def _check_stop_region_match(self, target_text):
        if not self.stop_condition["is_set"]:
            return False

        try:
            x1, y1, w, h = self.stop_condition["coords"]
            x2, y2 = x1 + w, y1 + h

            # 生成唯一临时文件名（避免多轮循环文件冲突）
            timestamp = int(time.time() * 1000000)  # 精确到微秒
            temp_img = os.path.join(output_dir, f"stop_current_{timestamp}.png")

            # 确保截图成功
            try:
                current_screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
                current_screenshot.save(temp_img)
                if not os.path.exists(temp_img):
                    raise Exception(f"截图文件未生成：{temp_img}")
            except Exception as img_err:
                self.root.after(0, lambda: self.status_var.set(f"停止区域截图失败：{str(img_err)}"))
                return False

            # 使用常驻OCR管道，避免每次检查都重新初始化
            if self.stop_ocr_pipeline is None or not hasattr(self.stop_ocr_pipeline, 'predict'):
                if not self.init_stop_ocr_pipeline():
                    if os.path.exists(temp_img):
                        os.remove(temp_img)
                    return False

            # 执行OCR识别
            output = list(self.stop_ocr_pipeline.predict([temp_img]))
            if not output:
                self.destroy_stop_ocr_pipeline()
                os.remove(temp_img)
                return False

            # 初始化json_path变量
            json_path = None
            
            # 保存并解析结果
            json_path = f"{os.path.splitext(temp_img)[0]}_res.json"
            output[0].save_to_json(json_path)
            time.sleep(0.1)  # 确保文件写入完成

            if not os.path.exists(json_path):
                raise Exception(f"OCR结果JSON未生成：{json_path}")

            # 解析文本
            user_fields = [f.strip() for f in self.ocr_fields_entry.get().split(',') if f.strip()] or \
                          self.stop_condition["ocr_text_fields"]
            ocr_texts = []

            with open(json_path, "r", encoding="utf-8") as f:
                ocr_data = json.load(f)

            for field in user_fields:
                value = self.get_nested_value(ocr_data, field)
                if value is None:
                    continue
                if isinstance(value, list):
                    ocr_texts.extend([str(t).strip() for t in value if str(t).strip()])
                elif isinstance(value, str):
                    ocr_texts.append(value.strip())
                break  # 取第一个有效字段

            # 强制清理临时文件
            for f in [temp_img, json_path]:
                if os.path.exists(f):
                    try:
                        os.remove(f)
                    except Exception as del_err:
                        self.root.after(0, lambda: self.status_var.set(f"临时文件清理警告：{str(del_err)}"))

            # 保持停止区域OCR管道常驻，避免频繁创建和销毁
            # 管道将在程序退出时统一销毁

            # 归一化匹配
            def normalize(s):
                s = s.replace('０', '0').replace('１', '1').replace('２', '2').replace('３', '3').replace('４', '4')
                s = s.replace('５', '5').replace('６', '6').replace('７', '7').replace('８', '8').replace('９', '9')
                s = s.replace('Ａ', 'A').replace('Ｂ', 'B').lower().strip()
                return s

            target_norm = normalize(target_text)
            match = any(normalize(t) == target_norm for t in ocr_texts)
            if match:
                self.root.after(0, lambda: self.status_var.set(
                    f"停止区域匹配成功：识别到「{ocr_texts}」=目标「{target_text}」"))
            return match

        except Exception as err:
            self.root.after(0, lambda: self.status_var.set(f"停止区域校验出错：{str(err)}，继续执行..."))
            self.destroy_stop_ocr_pipeline()
            return False

    def _check_any_region_match(self, target_text):
        regions = [e for e in self.elements if e["type"] == "region"]
        if not regions:
            return False

        for region in regions:
            data = region["data"]
            if not data["coords"]:
                continue

            # 预初始化变量，避免作用域问题
            temp_img = None
            json_path = None
            temp_files = []
            
            try:
                # 实时OCR识别
                x1, y1, w, h = data["coords"]
                x2, y2 = x1 + w, y1 + h

                # 生成唯一临时文件名
                timestamp = int(time.time() * 1000000)
                temp_img = os.path.join(output_dir, f"region_current_{timestamp}.png")
                temp_files.append(temp_img)  # 立即添加到清理列表
                
                try:
                    # 截图
                    try:
                        current_screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
                        current_screenshot.save(temp_img)
                        # 增加文件写入完成确认
                        for _ in range(3):
                            if os.path.exists(temp_img) and os.path.getsize(temp_img) > 0:
                                break
                            time.sleep(0.05)
                        
                        if not os.path.exists(temp_img):
                            raise Exception(f"截图文件未生成：{temp_img}")
                    except Exception as img_err:
                        raise Exception(f"截图失败：{type(img_err).__name__}: {str(img_err)}")

                    # 执行OCR识别
                    try:
                        output = list(self.region_ocr_pipeline.predict([temp_img]))
                        if not output:
                            raise Exception("OCR未识别到内容")
                    except Exception as ocr_err:
                        raise Exception(f"OCR识别失败：{type(ocr_err).__name__}: {str(ocr_err)}")

                    # 保存并解析结果
                    json_path = f"{os.path.splitext(temp_img)[0]}_res.json"
                    temp_files.append(json_path)  # 添加到清理列表
                    
                    try:
                        output[0].save_to_json(json_path)
                        # 增加文件写入完成确认
                        for _ in range(3):
                            if os.path.exists(json_path) and os.path.getsize(json_path) > 0:
                                break
                            time.sleep(0.05)

                        if not os.path.exists(json_path):
                            raise Exception(f"OCR结果JSON未生成：{json_path}")
                    except Exception as json_err:
                        raise Exception(f"JSON保存失败：{type(json_err).__name__}: {str(json_err)}")

                    # 解析文本
                    try:
                        with open(json_path, "r", encoding="utf-8") as f:
                            ocr_data = json.load(f)
                    except Exception as parse_err:
                        raise Exception(f"JSON解析失败：{type(parse_err).__name__}: {str(parse_err)}")

                    user_fields = [f.strip() for f in self.ocr_fields_entry.get().split(',') if f.strip()] or \
                                  self.stop_condition["ocr_text_fields"]
                    ocr_texts = []

                    for field in user_fields:
                        value = self.get_nested_value(ocr_data, field)
                        if value is None:
                            continue
                        if isinstance(value, list):
                            ocr_texts.extend([str(t).strip() for t in value if str(t).strip()])
                        elif isinstance(value, str):
                            ocr_texts.append(value.strip())
                        break

                    def normalize(s):
                        s = s.replace('０', '0').replace('１', '1').replace('２', '2').replace('３', '3').replace('４', '4')
                        s = s.replace('５', '5').replace('６', '6').replace('７', '7').replace('８', '8').replace('９', '9')
                        s = s.replace('Ａ', 'A').replace('Ｂ', 'B').lower()
                        return s.strip()

                    target_norm = normalize(target_text)
                    if any(normalize(t) == target_norm for t in ocr_texts):
                        # 使用参数传递变量，避免作用域问题
                        self.root.after(0, lambda data=data, texts=ocr_texts, target=target_text: 
                            self.status_var.set(
                                f"识别区域{data['current_id']}匹配成功：识别到「{texts}」=目标「{target}」"
                            )
                        )
                        return True
                finally:
                    # 安全清理临时文件 - 无论执行路径如何都确保清理
                    for f in temp_files:
                        if f and isinstance(f, str) and os.path.exists(f):
                            try:
                                os.remove(f)
                            except Exception:
                                # 静默忽略清理失败，避免干扰主流程
                                pass
                    # 清空临时文件列表
                    temp_files.clear()

            except Exception as err:
                error_type = type(err).__name__
                error_message = f"{error_type}: {str(err)}"
                region_id = data['current_id']
                # 使用参数传递变量，确保正确捕获异常信息
                self.root.after(0, lambda msg=error_message, rid=region_id: 
                    self.status_var.set(f"识别区域{rid}校验出错：{msg}，继续执行...")
                )
                continue

        return False

    # ------------------------------ 添加识别区域 ------------------------------
    def add_region(self):
        if self.is_running:
            messagebox.showinfo("提示", "运行中无法添加区域")
            return

        # 移除添加时就初始化OCR管道的操作，延迟到实际需要时（选择区域或运行时）再初始化

        original_index = len(self.elements)
        initial_id = original_index + 1

        # 使用自定义ElementFrame解决布局问题
        frame = ElementFrame(self.regions_frame, bd=2, relief=tk.GROOVE, padx=10, pady=5)
        frame.pack(fill=tk.X, padx=5, pady=5)

        # 调整列权重，让状态列可以伸缩
        frame.grid_columnconfigure(1, weight=1)

        id_label = tk.Label(
            frame,
            text=f"识别区域 {initial_id}",
            font=("Arial", 10, "bold"),
            width=15,
            anchor="w"
        )
        id_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        status_var = tk.StringVar()
        status_var.set("未选择区域")
        status_label = tk.Label(
            frame,
            textvariable=status_var,
            fg="orange",
            wraplength=300,  # 自动换行
            justify="left",
            anchor="w"
        )
        status_label.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        tk.Label(frame, text="目标匹配文本:", width=12, anchor="e").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        target_entry = tk.Entry(frame, width=15, font=("Arial", 10))
        target_entry.grid(row=0, column=3, padx=5, pady=5)

        select_btn = tk.Button(
            frame,
            text="选择区域",
            command=lambda: self.select_region(original_index)
        )
        select_btn.grid(row=0, column=4, padx=5, pady=5)

        delete_btn = tk.Button(
            frame,
            text="删除",
            fg="red",
            command=lambda: self.delete_element(original_index)
        )
        delete_btn.grid(row=0, column=5, padx=5, pady=5)

        self.elements.append({
            "type": "region",
            "original_index": original_index,
            "data": {
                "current_id": initial_id,
                "frame": frame,
                "id_label": id_label,
                "status_var": status_var,
                "status_label": status_label,
                "target_entry": target_entry,
                "coords": None
            }
        })

        self.regions_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.status_var.set(f"已添加识别区域 {initial_id}")

    # ------------------------------ 添加动态按钮 ------------------------------
    def add_dynamic_button(self):
        if self.is_running:
            messagebox.showinfo("提示", "运行中无法添加按钮")
            return

        original_index = len(self.elements)
        initial_id = original_index + 1

        # 使用自定义ElementFrame解决布局问题
        frame = ElementFrame(self.regions_frame, bd=2, relief=tk.GROOVE, padx=10, pady=5)
        frame.pack(fill=tk.X, padx=5, pady=5)

        # 调整列权重
        frame.grid_columnconfigure(1, weight=1)

        id_label = tk.Label(
            frame,
            text=f"点击按钮 {initial_id}",
            font=("Arial", 10, "bold"),
            width=15,
            anchor="w"
        )
        id_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        status_var = tk.StringVar()
        status_var.set("可拖动调整位置")
        status_label = tk.Label(
            frame,
            textvariable=status_var,
            fg="green",
            wraplength=300,  # 自动换行
            justify="left",
            anchor="w"
        )
        status_label.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        delete_btn = tk.Button(
            frame,
            text="删除",
            fg="red",
            command=lambda: self.delete_element(original_index)
        )
        delete_btn.grid(row=0, column=2, padx=5, pady=5)

        btn_window = tk.Toplevel(self.root)
        btn_window.overrideredirect(True)
        btn_window.attributes("-topmost", True)

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        default_x = screen_width - 50
        default_y = screen_height // 2
        btn_window.geometry(f"+{default_x}+{default_y}")

        # 使用中文数字标记或直接使用数字（当超过10个时）
        if initial_id <= len(self.number_marks):
            btn_text = self.number_marks[initial_id - 1]
        else:
            btn_text = str(initial_id)
            
        float_btn = tk.Button(
            btn_window,
            text=btn_text,
            font=("Arial", 12, "bold"),
            width=3,
            height=1,
            bg="#2196F3",
            fg="white",
            bd=2,
            relief=tk.RAISED
        )
        float_btn.pack()

        self.elements.append({
            "type": "button",
            "original_index": original_index,
            "data": {
                "current_id": initial_id,
                "frame": frame,
                "id_label": id_label,
                "status_var": status_var,
                "status_label": status_label,
                "window": btn_window,
                "button": float_btn,
                "x": default_x,
                "y": default_y,
                "width": 40,
                "height": 30
            }
        })

        self.setup_drag(float_btn, original_index)

        self.regions_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.status_var.set(f"已添加点击按钮 {initial_id}")

    # ------------------------------ 按钮拖动 ------------------------------
    def setup_drag(self, widget, elem_index):
        def on_press(e):
            self.drag_data["x"] = e.x
            self.drag_data["y"] = e.y
            self.drag_data["widget"] = widget
            self.drag_data["elem"] = self.elements[elem_index]
            elem = self.drag_data["elem"]
            self.drag_data["window_x"] = elem["data"]["x"]
            self.drag_data["window_y"] = elem["data"]["y"]

        def on_drag(e):
            if self.drag_data["widget"] is not widget:
                return

            elem = self.drag_data["elem"]
            screen_x = widget.winfo_pointerx()
            screen_y = widget.winfo_pointery()

            new_x = screen_x - self.drag_data["x"]
            new_y = screen_y - self.drag_data["y"]

            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            btn_width = elem["data"]["width"]
            btn_height = elem["data"]["height"]

            new_x = max(20, min(new_x, screen_width - btn_width - 20))
            new_y = max(20, min(new_y, screen_height - btn_height - 20))

            widget.master.geometry(f"+{new_x}+{new_y}")
            elem["data"]["x"] = new_x
            elem["data"]["y"] = new_y
            elem["data"]["status_var"].set(f"已定位：({new_x}, {new_y})")

        widget.bind("<ButtonPress-1>", on_press)

        def throttled_drag(e):
            widget.after(10, lambda: on_drag(e))

        widget.bind("<B1-Motion>", throttled_drag)

    # ------------------------------ 删除元素 ------------------------------
    def delete_element(self, target_original_index):
        if self.is_running:
            messagebox.showinfo("提示", "运行中无法删除元素")
            return

        index = next((i for i, e in enumerate(self.elements) if e["original_index"] == target_original_index), None)
        if index is None:
            return

        elem = self.elements.pop(index)
        elem["data"]["frame"].destroy()
        if elem["type"] == "button" and elem["data"]["window"].winfo_exists():
            elem["data"]["window"].destroy()

        # 删除文件
        if elem["type"] == "region":
            # 安全检查img_path是否存在
            if "img_path" in elem["data"]:
                img_path = elem["data"]["img_path"]
                if os.path.exists(img_path):
                    try:
                        os.remove(img_path)
                    except Exception as e:
                        messagebox.showwarning("删除失败", f"无法删除图片：{str(e)}")
            # 安全检查json_path是否存在
            if "json_path" in elem["data"]:
                json_path = elem["data"]["json_path"]
                if os.path.exists(json_path):
                    try:
                        os.remove(json_path)
                    except Exception as e:
                        messagebox.showwarning("删除失败", f"无法删除JSON：{str(e)}")

        # 更新ID
        for i, e in enumerate(self.elements):
            new_id = i + 1
            e["data"]["current_id"] = new_id
            e["data"]["id_label"]["text"] = f"{'识别区域' if e['type'] == 'region' else '点击按钮'} {new_id}"
            if e["type"] == "button":
                # 使用中文数字标记或直接使用数字（当超过10个时）
                if new_id <= len(self.number_marks):
                    e["data"]["button"]["text"] = self.number_marks[new_id - 1]
                else:
                    e["data"]["button"]["text"] = str(new_id)

        self.regions_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.status_var.set(f"已删除元素 | 剩余 {len(self.elements)} 个")

    # ------------------------------ 按ID删除按钮 ------------------------------
    def delete_button_by_id(self):
        if self.is_running:
            messagebox.showinfo("提示", "运行中无法删除按钮")
            return

        buttons = [e for e in self.elements if e["type"] == "button"]
        if not buttons:
            messagebox.showinfo("提示", "没有可删除的按钮")
            return

        id_list = [str(e["data"]["current_id"]) for e in buttons]
        id_str = simpledialog.askstring(
            "按ID删除按钮",
            f"请输入要删除的按钮ID（可选：{', '.join(id_list)}）:"
        )

        if not id_str or not id_str.isdigit():
            messagebox.showwarning("输入错误", "请输入有效的按钮ID")
            return

        target_id = int(id_str)
        target_elem = next((e for e in buttons if e["data"]["current_id"] == target_id), None)
        if not target_elem:
            messagebox.showinfo("未找到", f"未找到ID为 {target_id} 的按钮")
            return

        self.delete_element(target_elem["original_index"])

    # ------------------------------ 执行逻辑 ------------------------------
    def run_sequentially(self):
        if self.is_running:
            messagebox.showinfo("提示", "已有任务在运行中")
            return

        if not self.init_region_ocr_pipeline():
            messagebox.showerror("初始化失败", "识别区域OCR管道未初始化")
            return

        # 验证输入
        try:
            loop_count = int(self.loop_count_entry.get())
            if loop_count < 1:
                raise ValueError("循环次数必须大于0")
        except ValueError:
            messagebox.showerror("输入错误", "请输入有效的循环次数（正整数）")
            return

        try:
            interval = float(self.interval_entry.get())  # 元素间间隔（用于识别区域）
            if interval < 0:
                raise ValueError("元素间间隔不能为负数")
        except ValueError:
            messagebox.showerror("输入错误", "请输入有效的元素间间隔（非负数）")
            return

        try:
            button_interval = float(self.button_interval_entry.get())  # 按钮间间隔
            if button_interval < 0:
                raise ValueError("按钮间间隔不能为负数")
        except ValueError:
            messagebox.showerror("输入错误", "请输入有效的按钮间间隔（非负数）")
            return

        self.toggle_buttons_visibility(False)
        self.is_running = True
        self.execute_btn.config(state=tk.DISABLED)
        self.root.after(0, lambda: self.status_var.set(f"开始执行，总循环次数：{loop_count}"))

        def execute():
            try:
                total_loops = 0
                while total_loops < loop_count and (self.is_running or self.is_paused):
                    # 检查暂停状态
                    while self.is_paused:
                        time.sleep(0.1)  # 暂停时短暂休眠，避免CPU占用过高
                        if not (self.is_running or self.is_paused):
                            return  # 如果彻底停止，则退出
                    
                    # 每轮循环开始前检查停止条件
                    if self.check_stop_condition():
                        self.root.after(0, lambda: self.status_var.set("停止条件满足，终止执行"))
                        self.is_running = False
                        return

                    total_loops += 1
                    self.root.after(0, lambda: self.status_var.set(
                        f"正在执行第 {total_loops}/{loop_count} 轮循环"
                    ))

                    # 按ID顺序执行所有元素（区分间隔类型）
                    for elem in sorted(self.elements, key=lambda x: x["data"]["current_id"]):
                        # 检查暂停状态
                        while self.is_paused:
                            time.sleep(0.1)  # 暂停时短暂休眠，避免CPU占用过高
                            if not (self.is_running or self.is_paused):
                                return  # 如果彻底停止，则退出
                        
                        # 执行前检查停止条件（保留）
                        if not self.is_running or self.check_stop_condition():
                            self.root.after(0, lambda: self.status_var.set("停止条件满足，终止执行"))
                            self.is_running = False
                            return

                        # 执行元素操作
                        if elem["type"] == "region":
                            self.process_region(elem)
                        elif elem["type"] == "button":
                            self.process_button(elem)

                        # 移除执行后的重复检查，减少OCR调用频率
                        # 移除不必要的0.1秒延迟
                        
                        # 根据元素类型选择间隔时间
                        if elem["type"] == "button":
                            time.sleep(button_interval)  # 按钮使用按钮间间隔
                        else:
                            time.sleep(interval)  # 识别区域使用元素间间隔

                    # 一轮循环结束后检查停止条件（保留）
                    # 这已经是下一轮循环前的再次检查，不需要重复检查

                # 所有循环完成
                self.root.after(0, lambda: self.status_var.set(
                    f"所有 {loop_count} 轮循环执行完成"
                ))
            except Exception as e:
                error_details = f"{type(e).__name__}: {str(e)}"
                self.root.after(0, lambda: self.status_var.set(f"执行出错：{error_details}"))
            finally:
                # 只有在非暂停状态下才重置运行状态和清理资源
                if not self.is_paused:
                    self.is_running = False
                    
                    def finalize_execution():
                        try:
                            # 更新UI状态
                            self.execute_btn.config(state=tk.NORMAL)
                            self.toggle_buttons_visibility(True)
                            
                            # 释放OCR资源
                            self.root.after(0, lambda: self.status_var.set("正在释放资源..."))
                            self.destroy_stop_ocr_pipeline()
                            self.destroy_region_ocr_pipeline()
                            
                            # 检查是否是正常完成还是被中断
                            if 'total_loops' in locals() and total_loops >= loop_count:
                                self.status_var.set("所有循环已成功完成")
                            else:
                                self.status_var.set("循环已成功中断")
                        except Exception as e:
                            self.status_var.set(f"资源清理过程中发生错误: {str(e)}")
                    
                    # 使用UI线程执行最终清理操作
                    self.root.after(0, finalize_execution)

        threading.Thread(target=execute, daemon=True).start()

    # ------------------------------ 处理识别区域（实时识别） ------------------------------
    def process_region(self, elem):
        data = elem["data"]
        target_text = data["target_entry"].get().strip()
        if not target_text:
            self.root.after(0, lambda data=data: data["status_var"].set("跳过：未设置目标匹配文本"))
            return

        if not data["coords"]:
            self.root.after(0, lambda data=data: data["status_var"].set("跳过：区域未配置"))
            return

        # 实时OCR处理 - 预初始化所有变量
        json_path = None
        temp_img = None
        temp_files = []  # 预先创建临时文件列表
        
        try:
            # 使用常驻的region_ocr_pipeline，确保仅在需要时初始化（延迟初始化策略）
            if self.region_ocr_pipeline is None or not hasattr(self.region_ocr_pipeline, 'predict'):
                self.root.after(0, lambda data=data: data["status_var"].set("正在初始化OCR管道..."))
                if not self.init_region_ocr_pipeline():
                    self.root.after(0, lambda data=data: data["status_var"].set("OCR管道初始化失败"))
                    return
            
            x1, y1, w, h = data["coords"]
            x2, y2 = x1 + w, y1 + h

            # 生成唯一临时文件名（避免多轮循环文件冲突）
            timestamp = int(time.time() * 1000000)  # 精确到微秒
            temp_img = os.path.join(output_dir, f"region_current_{data['current_id']}_{timestamp}.png")
            temp_files.append(temp_img)  # 立即添加到清理列表

            # 确保截图成功
            try:
                current_screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
                current_screenshot.save(temp_img)
                if not os.path.exists(temp_img):
                    raise Exception(f"截图文件未生成：{temp_img}")
            except Exception as img_err:
                error_msg = f"区域截图失败：{type(img_err).__name__}: {str(img_err)}"
                self.root.after(0, lambda data=data, msg=error_msg: data["status_var"].set(msg))
                return

            # 执行OCR识别
            try:
                if self.region_ocr_pipeline is None:
                    raise Exception("OCR管道未初始化")
                
                output = list(self.region_ocr_pipeline.predict([temp_img]))
                if not output:
                    self.root.after(0, lambda data=data: data["status_var"].set("OCR识别无结果"))
                    return
            except Exception as ocr_err:
                error_msg = f"OCR识别失败：{type(ocr_err).__name__}: {str(ocr_err)}"
                self.root.after(0, lambda data=data, msg=error_msg: data["status_var"].set(msg))
                # 仅在发生错误时考虑重新初始化pipeline
                if isinstance(ocr_err, (AttributeError, RuntimeError)) and 'NoneType' in str(ocr_err):
                    self.destroy_region_ocr_pipeline()
                return

            # 保存并解析结果 - 使用JSON文件方式确保数据完整性
            json_path = f"{os.path.splitext(temp_img)[0]}_res.json"
            temp_files.append(json_path)  # 添加到清理列表
            
            try:
                output[0].save_to_json(json_path)
                # 增加文件写入完成确认
                for _ in range(3):  # 最多尝试3次
                    if os.path.exists(json_path) and os.path.getsize(json_path) > 0:
                        break
                    time.sleep(0.05)

                if not os.path.exists(json_path):
                    raise Exception(f"OCR结果JSON未生成：{json_path}")
            except Exception as json_err:
                error_msg = f"JSON保存失败：{type(json_err).__name__}: {str(json_err)}"
                self.root.after(0, lambda data=data, msg=error_msg: data["status_var"].set(msg))
                return

            # 解析OCR数据
            try:
                with open(json_path, "r", encoding="utf-8") as f:
                    ocr_data = json.load(f)
            except Exception as parse_err:
                error_msg = f"JSON解析失败：{type(parse_err).__name__}: {str(parse_err)}"
                self.root.after(0, lambda data=data, msg=error_msg: data["status_var"].set(msg))
                return

            # 获取用户配置的文本和坐标字段
            user_text_fields = [f.strip() for f in self.ocr_fields_entry.get().split(',') if f.strip()] or \
                               self.stop_condition["ocr_text_fields"]
            user_bbox_fields = [f.strip() for f in self.bbox_fields_entry.get().split(',') if f.strip()] or \
                               self.stop_condition["ocr_bbox_fields"]

            # 提取所有文本
            all_texts = []
            for text_field in user_text_fields:
                texts = self.get_nested_value(ocr_data, text_field)
                if texts is None:
                    continue
                if isinstance(texts, list):
                    all_texts = [str(t).strip() for t in texts if str(t).strip()]
                elif isinstance(texts, str):
                    all_texts = [texts.strip()]
                break

            if not all_texts:
                self.root.after(0, lambda data=data: data["status_var"].set("未识别到任何文本（检查文本字段配置）"))
                return

            # 提取所有坐标
            all_bboxes = []
            for bbox_field in user_bbox_fields:
                bboxes = self.get_nested_value(ocr_data, bbox_field)
                if bboxes is None:
                    continue
                if isinstance(bboxes, list):
                    valid_bboxes = []
                    for b in bboxes:
                        if isinstance(b, list) and len(b) == 4 and all(isinstance(num, (int, float)) for num in b):
                            valid_bboxes.append(b)
                    all_bboxes = valid_bboxes
                break

            # 使用全局normalize函数，避免重复定义
            # 注意：确保在文件顶部已定义全局normalize函数

            target_norm = normalize(target_text)
            matched_text = None
            matched_bbox = None

            # 匹配文本并关联坐标
            for i, text in enumerate(all_texts):
                if normalize(text) == target_norm:
                    matched_text = text
                    if i < len(all_bboxes):
                        matched_bbox = all_bboxes[i]
                    break

            # 处理匹配结果
            if matched_text:
                if matched_bbox:
                    # 计算屏幕坐标
                    try:
                        region_x1, region_y1, _, _ = data["coords"]
                        text_x1, text_y1, text_x2, text_y2 = map(float, matched_bbox)
                        screen_click_x = region_x1 + (text_x1 + text_x2) / 2
                        screen_click_y = region_y1 + (text_y1 + text_y2) / 2
                        # 执行点击
                        pyautogui.click(int(screen_click_x), int(screen_click_y))
                        result_msg = f"匹配成功并点击：{matched_text} = {target_text}（坐标：{int(screen_click_x)},{int(screen_click_y)}）"
                        self.root.after(0, lambda data=data, msg=result_msg: data["status_var"].set(msg))
                    except Exception as click_err:
                        error_msg = f"点击执行失败：{type(click_err).__name__}: {str(click_err)}"
                        self.root.after(0, lambda data=data, msg=error_msg: data["status_var"].set(msg))
                else:
                    self.root.after(0, lambda data=data, mt=matched_text, tt=target_text: 
                                    data["status_var"].set(f"匹配成功但无坐标：{mt} = {tt}（检查坐标字段配置）"))
            else:
                self.root.after(0, lambda data=data, at=str(all_texts), tt=target_text: 
                                data["status_var"].set(f"匹配失败：识别到{at} ≠ {tt}"))

        except Exception as err:
            error_type = type(err).__name__
            error_message = f"{error_type}: {str(err)}"
            print(f"process_button异常：{error_message}")
            if "status_var" in data:
                self.root.after(0, lambda data=data, msg=error_message: data["status_var"].set(f"处理失败：{msg}"))
        finally:
            # 安全清理临时文件 - 无论执行路径如何都确保清理
            for f in temp_files:
                if f and isinstance(f, str) and os.path.exists(f):
                    try:
                        os.remove(f)
                    except Exception:
                        # 静默忽略清理失败，避免干扰主流程
                        pass
            # 清空临时文件列表
            temp_files.clear()

    # ------------------------------ 处理按钮点击 ------------------------------
    def process_button(self, elem):
        data = elem["data"]

        try:
            # 添加属性存在检查，避免AttributeError
            x_offset = int(self.x_offset_entry.get()) if hasattr(self, 'x_offset_entry') and hasattr(self.x_offset_entry, 'get') else 0
            y_offset = int(self.y_offset_entry.get()) if hasattr(self, 'y_offset_entry') and hasattr(self.y_offset_entry, 'get') else 0
        except ValueError:
            x_offset = 0
            y_offset = 0

        click_x = data["x"] + data["width"] // 2 + x_offset
        click_y = data["y"] + data["height"] // 2 + y_offset

        # 恢复异常处理，确保点击失败不会导致程序崩溃
        try:
            pyautogui.click(click_x, click_y)
            self.root.after(0, lambda data=data, x=click_x, y=click_y: data["status_var"].set(f"已点击：({x}, {y})"))
        except Exception as click_err:
            error_msg = f"点击执行失败：{type(click_err).__name__}: {str(click_err)}"
            print(f"process_button点击异常：{error_msg}")
            if "status_var" in data:
                self.root.after(0, lambda data=data, msg=error_msg: data["status_var"].set(msg))


if __name__ == "__main__":
    app = RealTimeControl()
    app.root.mainloop()