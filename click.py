import tkinter as tk
from tkinter import ttk, messagebox, filedialog, colorchooser
import threading
import time
import sys
import os
import random                       # 反检测随机抖动
from pynput import mouse, keyboard
from pynput.mouse import Button, Controller as MouseController

# ======================= 全局异常处理 =======================
def show_error_and_exit(msg):
    messagebox.showerror("错误", msg)
    sys.exit(1)

try:
    from pynput import mouse, keyboard
except ImportError:
    show_error_and_exit("未安装 pynput 库，请运行 'pip install pynput' 安装。")

# ======================= DPI 处理 =======================
def get_dpi_scaling():
    """获取当前主显示器的 DPI 缩放比例（Windows）"""
    if sys.platform != "win32":
        return 1.0
    try:
        import ctypes
        hwnd = ctypes.windll.user32.GetDesktopWindow()
        dc = ctypes.windll.user32.GetDC(hwnd)
        dpi_x = ctypes.windll.gdi32.GetDeviceCaps(dc, 88)  # LOGPIXELSX
        ctypes.windll.user32.ReleaseDC(hwnd, dc)
        return dpi_x / 96.0
    except Exception:
        return 1.0

def set_dpi_aware():
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

set_dpi_aware()

# ======================= 图像处理（Pillow） =======================
try:
    from PIL import Image, ImageGrab
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    show_error_and_exit("请安装 Pillow 库: pip install Pillow")

# ======================= Windows 原生 OCR（可选） =======================
WINDOWS_OCR_AVAILABLE = False
if sys.platform == "win32":
    try:
        import winrt
        from winrt.windows.media.ocr import OcrEngine
        from winrt.windows.globalization import Language
        from winrt.windows.graphics.imaging import SoftwareBitmap, BitmapPixelFormat, BitmapAlphaMode
        from winrt.windows.storage.streams import InMemoryRandomAccessStream, Buffer
        from winrt.windows.graphics.capture import GraphicsCapturePicker
        from winrt.windows.graphics.imaging import BitmapDecoder
        WINDOWS_OCR_AVAILABLE = True
    except ImportError:
        pass  # winrt 未安装，无法使用
    except Exception:
        pass  # 其他原因导致不可用

class AutoClickerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("多功能连点器（反检测增强版）")
        self.geometry("900x700")                 # 增大窗口以容纳新选项卡
        self.resizable(False, False)
        
        self.mouse_ctrl = MouseController()
        
        # 状态标志
        self.clicking = False
        self.recording = False
        self.replaying = False
        self.flow_running = False
        self.stop_flag = threading.Event()
        
        # 录制数据
        self.recorded_actions = []
        self.record_start_time = 0
        self.record_start_pos = (0, 0)
        
        self.dpi_scale = get_dpi_scaling()
        self.flow_steps = []
        self.flow_vars = {}
        
        # ========== 反检测设置 ==========
        self.anti_detect_clicker = tk.BooleanVar(value=True)   # 连点器反检测开关
        self.anti_detect_replay = tk.BooleanVar(value=True)    # 回放反检测开关
        self.anti_detect_flow = tk.BooleanVar(value=True)      # 流程反检测开关
        self.time_jitter_percent = tk.DoubleVar(value=5.0)     # 时间抖动百分比 (0~20)
        self.position_jitter_pixels = tk.IntVar(value=2)       # 位置抖动像素 (0~10)
        
        # 创建选项卡
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.tab_clicker = ttk.Frame(self.notebook)
        self.tab_record = ttk.Frame(self.notebook)
        self.tab_flow = ttk.Frame(self.notebook)
        self.tab_settings = ttk.Frame(self.notebook)           # 新增：反检测设置选项卡
        self.notebook.add(self.tab_clicker, text="连点器")
        self.notebook.add(self.tab_record, text="录制与回放")
        self.notebook.add(self.tab_flow, text="点击流程")
        self.notebook.add(self.tab_settings, text="反检测设置")
        
        self.build_clicker_tab()
        self.build_record_tab()
        self.build_flow_tab()
        self.build_settings_tab()                              # 新增：构建反检测设置界面
        
        self.status_var = tk.StringVar(value=f"就绪 | DPI缩放: {self.dpi_scale:.2f}")
        status_label = ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w")
        status_label.pack(side="bottom", fill="x", padx=5, pady=5)
        
        # 全局热键
        self.listener = keyboard.Listener(on_press=self.on_key_press)
        self.listener.daemon = True
        self.listener.start()
        
        self.protocol("WM_DELETE_WINDOW", self.on_close)
    
    # ----------------------- GUI 构建 -----------------------
    def build_clicker_tab(self):
        frame = ttk.LabelFrame(self.tab_clicker, text="连点设置", padding=10)
        frame.pack(fill="x", padx=10, pady=10)
        ttk.Label(frame, text="点击类型:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.click_type_var = tk.StringVar(value="左键")
        click_type_menu = ttk.Combobox(frame, textvariable=self.click_type_var,
                                       values=["左键", "右键", "中键", "双击"], state="readonly")
        click_type_menu.grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(frame, text="每秒点击次数:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.speed_var = tk.StringVar(value="10")
        speed_entry = ttk.Entry(frame, textvariable=self.speed_var, width=10)
        speed_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        self.click_btn = ttk.Button(frame, text="开始连点", command=self.toggle_clicking)
        self.click_btn.grid(row=2, column=0, columnspan=2, pady=10)
    
    def build_record_tab(self):
        record_frame = ttk.LabelFrame(self.tab_record, text="录制设置", padding=10)
        record_frame.pack(fill="x", padx=10, pady=5)
        self.record_move_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(record_frame, text="录制鼠标移动", variable=self.record_move_var).pack(anchor="w", padx=5)
        self.record_btn = ttk.Button(record_frame, text="开始录制", command=self.toggle_recording)
        self.record_btn.pack(pady=5)
        
        replay_frame = ttk.LabelFrame(self.tab_record, text="回放设置", padding=10)
        replay_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(replay_frame, text="回放坐标模式:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.coord_mode_var = tk.StringVar(value="绝对坐标")
        coord_mode_menu = ttk.Combobox(replay_frame, textvariable=self.coord_mode_var,
                                       values=["绝对坐标", "相对坐标"], state="readonly", width=12)
        coord_mode_menu.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(replay_frame, text="回放速度倍数:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.playback_speed_var = tk.StringVar(value="1.0")
        playback_speed_entry = ttk.Entry(replay_frame, textvariable=self.playback_speed_var, width=6)
        playback_speed_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        btn_frame = ttk.Frame(replay_frame)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=5)
        self.replay_btn = ttk.Button(btn_frame, text="回放录制", command=self.start_replay, state="disabled")
        self.replay_btn.pack(side="left", padx=5)
        self.clear_rec_btn = ttk.Button(btn_frame, text="清空录制", command=self.clear_recording, state="disabled")
        self.clear_rec_btn.pack(side="left", padx=5)
    
    def build_flow_tab(self):
        list_frame = ttk.LabelFrame(self.tab_flow, text="流程步骤", padding=5)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        columns = ("序号", "动作", "详细")
        self.flow_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=12)
        self.flow_tree.heading("序号", text="序号")
        self.flow_tree.heading("动作", text="动作")
        self.flow_tree.heading("详细", text="详细")
        self.flow_tree.column("序号", width=50)
        self.flow_tree.column("动作", width=100)
        self.flow_tree.column("详细", width=450)
        self.flow_tree.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.flow_tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.flow_tree.configure(yscrollcommand=scrollbar.set)
        
        btn_frame = ttk.Frame(self.tab_flow)
        btn_frame.pack(fill="x", padx=10, pady=5)
        row1 = ttk.Frame(btn_frame)
        row1.pack(fill="x", pady=2)
        ttk.Button(row1, text="添加移动", command=self.add_move_step, width=12).pack(side="left", padx=2)
        ttk.Button(row1, text="添加点击", command=self.add_click_step, width=12).pack(side="left", padx=2)
        ttk.Button(row1, text="添加双击", command=self.add_doubleclick_step, width=12).pack(side="left", padx=2)
        ttk.Button(row1, text="添加等待", command=self.add_wait_step, width=12).pack(side="left", padx=2)
        row2 = ttk.Frame(btn_frame)
        row2.pack(fill="x", pady=2)
        if WINDOWS_OCR_AVAILABLE:
            ttk.Button(row2, text="文字识别(Windows)", command=self.add_ocr_win_step, width=16).pack(side="left", padx=2)
        else:
            ttk.Button(row2, text="文字识别(需Tesseract)", command=self.add_ocr_tesseract_step, width=16).pack(side="left", padx=2)
        ttk.Button(row2, text="图像匹配", command=self.add_image_match_step, width=12).pack(side="left", padx=2)
        ttk.Button(row2, text="颜色检测", command=self.add_color_detect_step, width=12).pack(side="left", padx=2)
        ttk.Button(row2, text="条件判断", command=self.add_branch_step, width=12).pack(side="left", padx=2)
        row3 = ttk.Frame(btn_frame)
        row3.pack(fill="x", pady=2)
        ttk.Button(row3, text="删除", command=self.delete_step, width=12).pack(side="left", padx=2)
        ttk.Button(row3, text="上移", command=self.move_up_step, width=12).pack(side="left", padx=2)
        ttk.Button(row3, text="下移", command=self.move_down_step, width=12).pack(side="left", padx=2)
        ttk.Button(row3, text="清空", command=self.clear_flow, width=12).pack(side="left", padx=2)
        
        ctrl_frame = ttk.Frame(self.tab_flow)
        ctrl_frame.pack(fill="x", padx=10, pady=5)
        self.run_flow_btn = ttk.Button(ctrl_frame, text="执行流程", command=self.run_flow, width=12)
        self.run_flow_btn.pack(side="left", padx=5)
        self.stop_flow_btn = ttk.Button(ctrl_frame, text="停止流程", command=self.stop_flow, state="disabled", width=12)
        self.stop_flow_btn.pack(side="left", padx=5)
        
        info_text = "提示：\n- 文字识别(Windows)需要Windows 10/11且已安装winrt库，无需Tesseract。\n- 图像匹配：提前截取小图，在指定区域内搜索相似度。\n- 颜色检测：检测指定区域内是否存在指定颜色。"
        info_label = ttk.Label(self.tab_flow, text=info_text, foreground="gray", justify="left")
        info_label.pack(pady=5)
    
    def build_settings_tab(self):
        frame = ttk.LabelFrame(self.tab_settings, text="反检测设置", padding=10)
        frame.pack(fill="x", padx=10, pady=10)
        
        # 三个独立开关
        ttk.Checkbutton(frame, text="连点器启用反检测", variable=self.anti_detect_clicker).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Checkbutton(frame, text="回放启用反检测", variable=self.anti_detect_replay).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Checkbutton(frame, text="流程启用反检测", variable=self.anti_detect_flow).grid(row=2, column=0, sticky="w", padx=5, pady=5)
        
        # 时间抖动滑块
        ttk.Label(frame, text="时间抖动百分比 (0~20%):").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        jitter_slider = ttk.Scale(frame, from_=0, to=20, variable=self.time_jitter_percent, orient="horizontal")
        jitter_slider.grid(row=3, column=1, sticky="ew", padx=5, pady=5)
        ttk.Label(frame, textvariable=self.time_jitter_percent).grid(row=3, column=2, padx=5)
        
        # 位置抖动滑块
        ttk.Label(frame, text="位置抖动像素 (0~10px):").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        pos_slider = ttk.Scale(frame, from_=0, to=10, variable=self.position_jitter_pixels, orient="horizontal")
        pos_slider.grid(row=4, column=1, sticky="ew", padx=5, pady=5)
        ttk.Label(frame, textvariable=self.position_jitter_pixels).grid(row=4, column=2, padx=5)
        
        frame.columnconfigure(1, weight=1)
    
    # ======================= 辅助函数：应用抖动 =======================
    def apply_time_jitter(self, base_delay, enabled):
        """根据反检测设置返回实际延迟时间（秒）"""
        if enabled and self.time_jitter_percent.get() > 0:
            jitter = random.uniform(-self.time_jitter_percent.get() / 100.0, self.time_jitter_percent.get() / 100.0)
            return max(0.001, base_delay * (1 + jitter))   # 确保最小延迟
        return base_delay
    
    def apply_position_jitter(self, x, y, enabled):
        """根据反检测设置返回抖动后的坐标"""
        if enabled and self.position_jitter_pixels.get() > 0:
            jitter_x = random.randint(-self.position_jitter_pixels.get(), self.position_jitter_pixels.get())
            jitter_y = random.randint(-self.position_jitter_pixels.get(), self.position_jitter_pixels.get())
            return x + jitter_x, y + jitter_y
        return x, y
    
    # ----------------------- 连点器功能（添加反检测） -----------------------
    def toggle_clicking(self):
        if self.clicking:
            self.stop_clicking()
        else:
            self.start_clicking()
    
    def start_clicking(self):
        if self.recording or self.replaying or self.flow_running:
            messagebox.showwarning("警告", "请先停止其他操作")
            return
        try:
            self.click_interval = 1.0 / float(self.speed_var.get())
            if self.click_interval <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("错误", "请输入有效的每秒点击次数")
            return
        self.clicking = True
        self.stop_flag.clear()
        self.click_btn.config(text="停止连点")
        self.status_var.set(f"连点中... | DPI缩放: {self.dpi_scale:.2f}")
        self.click_thread = threading.Thread(target=self.clicking_loop, daemon=True)
        self.click_thread.start()
    
    def stop_clicking(self):
        self.clicking = False
        self.stop_flag.set()
        self.click_btn.config(text="开始连点")
        self.status_var.set(f"已停止连点 | DPI缩放: {self.dpi_scale:.2f}")
    
    def clicking_loop(self):
        click_type = self.click_type_var.get()
        anti_detect = self.anti_detect_clicker.get()
        while self.clicking and not self.stop_flag.is_set():
            # 执行点击（位置抖动在 perform_click 内部处理）
            self.perform_click(click_type, anti_detect)
            # 计算本次延迟（带时间抖动）
            base_delay = self.click_interval
            delay = self.apply_time_jitter(base_delay, anti_detect)
            # 等待
            for _ in range(int(delay * 10)):
                if not self.clicking or self.stop_flag.is_set():
                    break
                time.sleep(delay / 10)
    
    def perform_click(self, click_type, anti_detect):
        # 获取当前鼠标位置并抖动
        x, y = self.mouse_ctrl.position
        if anti_detect:
            x, y = self.apply_position_jitter(x, y, True)
            self.mouse_ctrl.position = (x, y)
        if click_type == "左键":
            self.mouse_ctrl.click(Button.left, 1)
        elif click_type == "右键":
            self.mouse_ctrl.click(Button.right, 1)
        elif click_type == "中键":
            self.mouse_ctrl.click(Button.middle, 1)
        elif click_type == "双击":
            self.mouse_ctrl.click(Button.left, 1)
            time.sleep(0.1)
            self.mouse_ctrl.click(Button.left, 1)
    
    # ----------------------- 录制与回放功能（添加反检测） -----------------------
    def toggle_recording(self):
        if self.recording:
            self.stop_recording()
        else:
            self.start_recording()
    
    def start_recording(self):
        if self.clicking or self.replaying or self.flow_running:
            messagebox.showwarning("警告", "请先停止其他操作")
            return
        self.recorded_actions.clear()
        self.recording = True
        self.record_btn.config(text="停止录制")
        self.replay_btn.config(state="disabled")
        self.clear_rec_btn.config(state="disabled")
        self.status_var.set(f"录制中... | DPI缩放: {self.dpi_scale:.2f}")
        self.record_start_time = time.time()
        self.record_start_pos = self.mouse_ctrl.position
        self.mouse_listener = mouse.Listener(on_move=self.on_mouse_move, on_click=self.on_mouse_click)
        self.mouse_listener.daemon = True
        self.mouse_listener.start()
    
    def on_mouse_move(self, x, y):
        if not self.recording or not self.record_move_var.get():
            return
        now = time.time()
        delay = now - self.record_start_time
        self.recorded_actions.append((delay, "move", x, y, None))
        self.status_var.set(f"录制中... 已录制 {len(self.recorded_actions)} 个动作")
    
    def on_mouse_click(self, x, y, button, pressed):
        if not self.recording or not pressed:
            return
        now = time.time()
        delay = now - self.record_start_time
        self.recorded_actions.append((delay, "click", x, y, button))
        self.status_var.set(f"录制中... 已录制 {len(self.recorded_actions)} 个动作")
    
    def stop_recording(self):
        self.recording = False
        if hasattr(self, 'mouse_listener') and self.mouse_listener.running:
            self.mouse_listener.stop()
        self.record_btn.config(text="开始录制")
        if self.recorded_actions:
            self.replay_btn.config(state="normal")
            self.clear_rec_btn.config(state="normal")
            self.status_var.set(f"录制完成，共 {len(self.recorded_actions)} 个动作")
        else:
            self.status_var.set("录制结束，未录制到任何动作")
    
    def clear_recording(self):
        self.recorded_actions.clear()
        self.replay_btn.config(state="disabled")
        self.clear_rec_btn.config(state="disabled")
        self.status_var.set("已清空录制内容")
    
    def start_replay(self):
        if not self.recorded_actions:
            messagebox.showinfo("提示", "没有可回放的操作")
            return
        if self.clicking or self.recording or self.flow_running:
            messagebox.showwarning("警告", "请先停止其他操作")
            return
        try:
            speed = float(self.playback_speed_var.get())
            if speed <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("错误", "请输入有效的回放速度倍数（正数）")
            return
        self.replaying = True
        self.stop_flag.clear()
        self.replay_btn.config(text="停止回放", state="normal")
        self.clear_rec_btn.config(state="disabled")
        self.status_var.set("回放中...")
        self.replay_thread = threading.Thread(target=self.replay_loop, args=(speed,), daemon=True)
        self.replay_thread.start()
    
    def replay_loop(self, speed):
        actions = self.recorded_actions
        start_time = time.time()
        coord_mode = self.coord_mode_var.get()
        replay_start_pos = self.mouse_ctrl.position if coord_mode == "相对坐标" else None
        anti_detect = self.anti_detect_replay.get()
        for i, (delay, action_type, x, y, button) in enumerate(actions):
            if not self.replaying or self.stop_flag.is_set():
                break
            # 计算理论等待时间并添加抖动
            base_wait = delay / speed
            wait = self.apply_time_jitter(base_wait, anti_detect)
            target_time = start_time + wait
            now = time.time()
            if target_time > now:
                time.sleep(target_time - now)
            # 计算最终坐标
            if coord_mode == "相对坐标":
                final_x = replay_start_pos[0] + (x - self.record_start_pos[0])
                final_y = replay_start_pos[1] + (y - self.record_start_pos[1])
                # 相对坐标模式不添加位置抖动，以免破坏相对关系
            else:
                final_x, final_y = x, y
                if anti_detect and action_type in ["move", "click"]:
                    final_x, final_y = self.apply_position_jitter(final_x, final_y, True)
            # 执行动作
            if action_type == "move":
                self.mouse_ctrl.position = (final_x, final_y)
            elif action_type == "click":
                self.mouse_ctrl.position = (final_x, final_y)
                self.mouse_ctrl.click(button, 1)
            self.status_var.set(f"回放中... 已完成 {i+1}/{len(actions)}")
        if self.replaying and not self.stop_flag.is_set():
            self.status_var.set("回放完成")
        self.replaying = False
        self.replay_btn.config(text="回放录制", state="normal" if self.recorded_actions else "disabled")
        self.clear_rec_btn.config(state="normal" if self.recorded_actions else "disabled")
        self.stop_flag.clear()
    
    def stop_replay(self):
        if self.replaying:
            self.replaying = False
            self.stop_flag.set()
            self.status_var.set("已停止回放")
            self.replay_btn.config(text="回放录制", state="normal" if self.recorded_actions else "disabled")
            self.clear_rec_btn.config(state="normal" if self.recorded_actions else "disabled")
    
    # ----------------------- 流程步骤添加 -----------------------
    def add_move_step(self):
        if self.flow_running:
            messagebox.showwarning("警告", "请先停止当前流程")
            return
        x, y = self.mouse_ctrl.position
        self.flow_steps.append({"type": "move", "x": x, "y": y})
        self.refresh_flow_tree()
        self.status_var.set(f"已添加移动步骤 ({x}, {y})")
    
    def add_click_step(self):
        if self.flow_running:
            messagebox.showwarning("警告", "请先停止当前流程")
            return
        dialog = tk.Toplevel(self)
        dialog.title("选择鼠标按钮")
        dialog.geometry("200x100")
        dialog.transient(self)
        dialog.grab_set()
        button_var = tk.StringVar(value="left")
        ttk.Label(dialog, text="选择按钮:").pack(pady=5)
        ttk.Combobox(dialog, textvariable=button_var, values=["left", "right", "middle"], state="readonly").pack(pady=5)
        def on_ok():
            x, y = self.mouse_ctrl.position
            self.flow_steps.append({"type": "click", "x": x, "y": y, "button": button_var.get()})
            self.refresh_flow_tree()
            self.status_var.set(f"已添加点击步骤 ({x}, {y}) 按钮: {button_var.get()}")
            dialog.destroy()
        ttk.Button(dialog, text="确定", command=on_ok).pack(pady=5)
    
    def add_doubleclick_step(self):
        if self.flow_running:
            messagebox.showwarning("警告", "请先停止当前流程")
            return
        dialog = tk.Toplevel(self)
        dialog.title("选择鼠标按钮")
        dialog.geometry("200x100")
        dialog.transient(self)
        dialog.grab_set()
        button_var = tk.StringVar(value="left")
        ttk.Label(dialog, text="选择按钮:").pack(pady=5)
        ttk.Combobox(dialog, textvariable=button_var, values=["left", "right", "middle"], state="readonly").pack(pady=5)
        def on_ok():
            x, y = self.mouse_ctrl.position
            self.flow_steps.append({"type": "doubleclick", "x": x, "y": y, "button": button_var.get()})
            self.refresh_flow_tree()
            self.status_var.set(f"已添加双击步骤 ({x}, {y}) 按钮: {button_var.get()}")
            dialog.destroy()
        ttk.Button(dialog, text="确定", command=on_ok).pack(pady=5)
    
    def add_wait_step(self):
        if self.flow_running:
            messagebox.showwarning("警告", "请先停止当前流程")
            return
        dialog = tk.Toplevel(self)
        dialog.title("设置等待时间")
        dialog.geometry("200x100")
        dialog.transient(self)
        dialog.grab_set()
        seconds_var = tk.StringVar(value="1.0")
        ttk.Label(dialog, text="等待秒数:").pack(pady=5)
        ttk.Entry(dialog, textvariable=seconds_var).pack(pady=5)
        def on_ok():
            try:
                sec = float(seconds_var.get())
                if sec <= 0:
                    raise ValueError
                self.flow_steps.append({"type": "wait", "seconds": sec})
                self.refresh_flow_tree()
                self.status_var.set(f"已添加等待步骤 {sec} 秒")
                dialog.destroy()
            except ValueError:
                messagebox.showerror("错误", "请输入正数秒数")
        ttk.Button(dialog, text="确定", command=on_ok).pack(pady=5)
    
    def add_ocr_win_step(self):
        """使用 Windows 原生 OCR"""
        if not WINDOWS_OCR_AVAILABLE:
            messagebox.showerror("错误", "Windows OCR 不可用，请确保已安装 winrt 库且系统为 Windows 10/11。")
            return
        if self.flow_running:
            messagebox.showwarning("警告", "请先停止当前流程")
            return
        dialog = tk.Toplevel(self)
        dialog.title("添加文字识别步骤 (Windows OCR)")
        dialog.geometry("400x280")
        dialog.transient(self)
        dialog.grab_set()
        ttk.Label(dialog, text="截图区域（屏幕坐标，左上角、右下角）:").pack(pady=5)
        coord_frame = ttk.Frame(dialog)
        coord_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(coord_frame, text="X1:").grid(row=0, column=0, sticky="e")
        x1_var = tk.StringVar()
        ttk.Entry(coord_frame, textvariable=x1_var, width=8).grid(row=0, column=1, padx=2)
        ttk.Label(coord_frame, text="Y1:").grid(row=0, column=2, sticky="e")
        y1_var = tk.StringVar()
        ttk.Entry(coord_frame, textvariable=y1_var, width=8).grid(row=0, column=3, padx=2)
        ttk.Label(coord_frame, text="X2:").grid(row=1, column=0, sticky="e")
        x2_var = tk.StringVar()
        ttk.Entry(coord_frame, textvariable=x2_var, width=8).grid(row=1, column=1, padx=2)
        ttk.Label(coord_frame, text="Y2:").grid(row=1, column=2, sticky="e")
        y2_var = tk.StringVar()
        ttk.Entry(coord_frame, textvariable=y2_var, width=8).grid(row=1, column=3, padx=2)
        def get_current_pos():
            x, y = self.mouse_ctrl.position
            return x, y
        def set_from_mouse(entry_var):
            x, y = get_current_pos()
            entry_var.set(str(x))
        quick_frame = ttk.Frame(dialog)
        quick_frame.pack(pady=5)
        ttk.Button(quick_frame, text="获取鼠标 X1", command=lambda: set_from_mouse(x1_var)).pack(side="left", padx=2)
        ttk.Button(quick_frame, text="获取鼠标 Y1", command=lambda: set_from_mouse(y1_var)).pack(side="left", padx=2)
        ttk.Button(quick_frame, text="获取鼠标 X2", command=lambda: set_from_mouse(x2_var)).pack(side="left", padx=2)
        ttk.Button(quick_frame, text="获取鼠标 Y2", command=lambda: set_from_mouse(y2_var)).pack(side="left", padx=2)
        ttk.Label(dialog, text="识别语言（如 zh-cn, en）:").pack(pady=5)
        lang_var = tk.StringVar(value="zh-cn")
        lang_combo = ttk.Combobox(dialog, textvariable=lang_var, values=["zh-cn", "en", "ja"], state="readonly")
        lang_combo.pack()
        ttk.Label(dialog, text="存储变量名（默认 ocr_result）:").pack(pady=5)
        var_name_var = tk.StringVar(value="ocr_result")
        ttk.Entry(dialog, textvariable=var_name_var).pack()
        def on_ok():
            try:
                x1 = int(x1_var.get()); y1 = int(y1_var.get()); x2 = int(x2_var.get()); y2 = int(y2_var.get())
                if x1 >= x2 or y1 >= y2:
                    raise ValueError
            except ValueError:
                messagebox.showerror("错误", "请输入有效的整数坐标，且左上角小于右下角")
                return
            step = {
                "type": "ocr_win",
                "x1": x1, "y1": y1, "x2": x2, "y2": y2,
                "lang": lang_var.get(),
                "var": var_name_var.get()
            }
            self.flow_steps.append(step)
            self.refresh_flow_tree()
            self.status_var.set("已添加 Windows OCR 步骤")
            dialog.destroy()
        ttk.Button(dialog, text="确定", command=on_ok).pack(pady=10)
    
    def add_ocr_tesseract_step(self):
        """使用 Tesseract OCR（需要安装 Tesseract 引擎）"""
        messagebox.showinfo("提示", "使用 Tesseract OCR 需要先安装 Tesseract 引擎，请参考网络教程安装后，再安装 pytesseract 库。\n如需免安装，请尝试 Windows OCR 或图像匹配。")
    
    def add_image_match_step(self):
        """图像匹配：在指定区域寻找目标图片"""
        if self.flow_running:
            messagebox.showwarning("警告", "请先停止当前流程")
            return
        dialog = tk.Toplevel(self)
        dialog.title("添加图像匹配步骤")
        dialog.geometry("500x400")
        dialog.transient(self)
        dialog.grab_set()
        ttk.Label(dialog, text="目标图片文件:").pack(pady=5)
        img_path_var = tk.StringVar()
        img_entry = ttk.Entry(dialog, textvariable=img_path_var, width=50)
        img_entry.pack(padx=10)
        def browse_file():
            path = filedialog.askopenfilename(filetypes=[("Image files", "*.png *.jpg *.bmp")])
            if path:
                img_path_var.set(path)
        ttk.Button(dialog, text="浏览", command=browse_file).pack(pady=2)
        ttk.Label(dialog, text="搜索区域（屏幕坐标，可选，留空则全屏）:").pack(pady=5)
        coord_frame = ttk.Frame(dialog)
        coord_frame.pack(fill="x", padx=10)
        ttk.Label(coord_frame, text="X1:").grid(row=0, column=0, sticky="e")
        x1_var = tk.StringVar()
        ttk.Entry(coord_frame, textvariable=x1_var, width=6).grid(row=0, column=1, padx=2)
        ttk.Label(coord_frame, text="Y1:").grid(row=0, column=2, sticky="e")
        y1_var = tk.StringVar()
        ttk.Entry(coord_frame, textvariable=y1_var, width=6).grid(row=0, column=3, padx=2)
        ttk.Label(coord_frame, text="X2:").grid(row=1, column=0, sticky="e")
        x2_var = tk.StringVar()
        ttk.Entry(coord_frame, textvariable=x2_var, width=6).grid(row=1, column=1, padx=2)
        ttk.Label(coord_frame, text="Y2:").grid(row=1, column=2, sticky="e")
        y2_var = tk.StringVar()
        ttk.Entry(coord_frame, textvariable=y2_var, width=6).grid(row=1, column=3, padx=2)
        ttk.Label(dialog, text="相似度阈值（0-1，默认0.8）:").pack(pady=5)
        threshold_var = tk.StringVar(value="0.8")
        ttk.Entry(dialog, textvariable=threshold_var, width=10).pack()
        ttk.Label(dialog, text="存储变量名（将保存匹配到的坐标，格式 'x,y'，找不到则为空）:").pack(pady=5)
        var_name_var = tk.StringVar(value="match_pos")
        ttk.Entry(dialog, textvariable=var_name_var).pack()
        def on_ok():
            img_path = img_path_var.get()
            if not img_path or not os.path.exists(img_path):
                messagebox.showerror("错误", "请选择有效的图片文件")
                return
            try:
                threshold = float(threshold_var.get())
                if not (0 <= threshold <= 1):
                    raise ValueError
            except:
                messagebox.showerror("错误", "相似度阈值应为0到1之间的数字")
                return
            region = None
            if x1_var.get() and y1_var.get() and x2_var.get() and y2_var.get():
                try:
                    x1 = int(x1_var.get()); y1 = int(y1_var.get()); x2 = int(x2_var.get()); y2 = int(y2_var.get())
                    if x1 >= x2 or y1 >= y2:
                        raise ValueError
                    region = (x1, y1, x2, y2)
                except:
                    messagebox.showerror("错误", "区域坐标无效")
                    return
            step = {
                "type": "image_match",
                "image_path": img_path,
                "region": region,
                "threshold": threshold,
                "var": var_name_var.get()
            }
            self.flow_steps.append(step)
            self.refresh_flow_tree()
            self.status_var.set("已添加图像匹配步骤")
            dialog.destroy()
        ttk.Button(dialog, text="确定", command=on_ok).pack(pady=10)
    
    def add_color_detect_step(self):
        """颜色检测：在指定区域检测是否存在目标颜色"""
        if self.flow_running:
            messagebox.showwarning("警告", "请先停止当前流程")
            return
        dialog = tk.Toplevel(self)
        dialog.title("添加颜色检测步骤")
        dialog.geometry("450x350")
        dialog.transient(self)
        dialog.grab_set()
        ttk.Label(dialog, text="检测区域（屏幕坐标，左上角、右下角）:").pack(pady=5)
        coord_frame = ttk.Frame(dialog)
        coord_frame.pack(fill="x", padx=10)
        ttk.Label(coord_frame, text="X1:").grid(row=0, column=0, sticky="e")
        x1_var = tk.StringVar()
        ttk.Entry(coord_frame, textvariable=x1_var, width=6).grid(row=0, column=1, padx=2)
        ttk.Label(coord_frame, text="Y1:").grid(row=0, column=2, sticky="e")
        y1_var = tk.StringVar()
        ttk.Entry(coord_frame, textvariable=y1_var, width=6).grid(row=0, column=3, padx=2)
        ttk.Label(coord_frame, text="X2:").grid(row=1, column=0, sticky="e")
        x2_var = tk.StringVar()
        ttk.Entry(coord_frame, textvariable=x2_var, width=6).grid(row=1, column=1, padx=2)
        ttk.Label(coord_frame, text="Y2:").grid(row=1, column=2, sticky="e")
        y2_var = tk.StringVar()
        ttk.Entry(coord_frame, textvariable=y2_var, width=6).grid(row=1, column=3, padx=2)
        ttk.Label(dialog, text="目标颜色（RGB）:").pack(pady=5)
        color_btn = ttk.Button(dialog, text="选择颜色", command=lambda: choose_color())
        color_btn.pack()
        color_rgb_var = tk.StringVar(value="(255,0,0)")
        ttk.Label(dialog, textvariable=color_rgb_var).pack()
        def choose_color():
            color = colorchooser.askcolor(title="选择颜色")
            if color[0]:
                r,g,b = map(int, color[0])
                color_rgb_var.set(f"({r},{g},{b})")
        ttk.Label(dialog, text="颜色容差（0-255，默认10）:").pack(pady=5)
        tolerance_var = tk.StringVar(value="10")
        ttk.Entry(dialog, textvariable=tolerance_var, width=10).pack()
        ttk.Label(dialog, text="存储变量名（将保存 True/False）:").pack(pady=5)
        var_name_var = tk.StringVar(value="color_found")
        ttk.Entry(dialog, textvariable=var_name_var).pack()
        def on_ok():
            try:
                x1 = int(x1_var.get()); y1 = int(y1_var.get()); x2 = int(x2_var.get()); y2 = int(y2_var.get())
                if x1 >= x2 or y1 >= y2:
                    raise ValueError
            except:
                messagebox.showerror("错误", "请输入有效的区域坐标，且左上角小于右下角")
                return
            try:
                r,g,b = map(int, color_rgb_var.get().strip("()").split(","))
            except:
                messagebox.showerror("错误", "颜色格式无效")
                return
            try:
                tolerance = int(tolerance_var.get())
                if tolerance < 0 or tolerance > 255:
                    raise ValueError
            except:
                messagebox.showerror("错误", "容差应为0-255的整数")
                return
            step = {
                "type": "color_detect",
                "x1": x1, "y1": y1, "x2": x2, "y2": y2,
                "color": (r,g,b),
                "tolerance": tolerance,
                "var": var_name_var.get()
            }
            self.flow_steps.append(step)
            self.refresh_flow_tree()
            self.status_var.set("已添加颜色检测步骤")
            dialog.destroy()
        ttk.Button(dialog, text="确定", command=on_ok).pack(pady=10)
    
    def add_branch_step(self):
        if self.flow_running:
            messagebox.showwarning("警告", "请先停止当前流程")
            return
        dialog = tk.Toplevel(self)
        dialog.title("添加条件判断步骤")
        dialog.geometry("400x300")
        dialog.transient(self)
        dialog.grab_set()
        ttk.Label(dialog, text="变量名（如 ocr_result, match_pos, color_found）:").pack(pady=5)
        var_name_var = tk.StringVar(value="ocr_result")
        ttk.Entry(dialog, textvariable=var_name_var).pack()
        ttk.Label(dialog, text="比较方式:").pack(pady=5)
        op_var = tk.StringVar(value="contains")
        op_combo = ttk.Combobox(dialog, textvariable=op_var, values=["contains", "equals", "not empty", "not empty string"], state="readonly")
        op_combo.pack()
        ttk.Label(dialog, text="比较值（可为空）:").pack(pady=5)
        value_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=value_var).pack()
        ttk.Label(dialog, text="条件成立时跳转到步骤序号（1-based）:").pack(pady=5)
        step_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=step_var).pack()
        def on_ok():
            var_name = var_name_var.get().strip()
            if not var_name:
                messagebox.showerror("错误", "请输入变量名")
                return
            op = op_var.get()
            compare_val = value_var.get()
            if op not in ["not empty", "not empty string"] and compare_val == "":
                messagebox.showerror("错误", "请输入比较值")
                return
            try:
                target_step = int(step_var.get())
                if target_step <= 0:
                    raise ValueError
            except:
                messagebox.showerror("错误", "请输入有效的正整数步骤序号")
                return
            step = {
                "type": "branch",
                "var": var_name,
                "op": op,
                "value": compare_val,
                "target_step": target_step
            }
            self.flow_steps.append(step)
            self.refresh_flow_tree()
            self.status_var.set("已添加条件判断步骤")
            dialog.destroy()
        ttk.Button(dialog, text="确定", command=on_ok).pack(pady=10)
    
    def refresh_flow_tree(self):
        for item in self.flow_tree.get_children():
            self.flow_tree.delete(item)
        for idx, step in enumerate(self.flow_steps, 1):
            if step["type"] == "move":
                action = "移动"
                detail = f"({step['x']}, {step['y']})"
            elif step["type"] == "click":
                action = "点击"
                detail = f"({step['x']}, {step['y']}) 按钮: {step['button']}"
            elif step["type"] == "doubleclick":
                action = "双击"
                detail = f"({step['x']}, {step['y']}) 按钮: {step['button']}"
            elif step["type"] == "wait":
                action = "等待"
                detail = f"{step['seconds']} 秒"
            elif step["type"] == "ocr_win":
                action = "文字识别(Windows)"
                detail = f"区域({step['x1']},{step['y1']})-({step['x2']},{step['y2']}) 语言:{step['lang']}"
            elif step["type"] == "image_match":
                action = "图像匹配"
                detail = f"图片:{os.path.basename(step['image_path'])} 区域:{step['region'] if step['region'] else '全屏'} 阈值:{step['threshold']}"
            elif step["type"] == "color_detect":
                action = "颜色检测"
                detail = f"区域({step['x1']},{step['y1']})-({step['x2']},{step['y2']}) 颜色:{step['color']} 容差:{step['tolerance']}"
            elif step["type"] == "branch":
                action = "条件判断"
                detail = f"如果 {step['var']} {step['op']} '{step['value']}' 则跳转 {step['target_step']}"
            else:
                continue
            self.flow_tree.insert("", "end", values=(idx, action, detail))
    
    def delete_step(self):
        if self.flow_running:
            messagebox.showwarning("警告", "请先停止当前流程")
            return
        selected = self.flow_tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选中一个步骤")
            return
        idx = self.flow_tree.index(selected[0])
        del self.flow_steps[idx]
        self.refresh_flow_tree()
        self.status_var.set("已删除步骤")
    
    def move_up_step(self):
        if self.flow_running:
            messagebox.showwarning("警告", "请先停止当前流程")
            return
        selected = self.flow_tree.selection()
        if not selected:
            return
        idx = self.flow_tree.index(selected[0])
        if idx > 0:
            self.flow_steps[idx], self.flow_steps[idx-1] = self.flow_steps[idx-1], self.flow_steps[idx]
            self.refresh_flow_tree()
            new_id = self.flow_tree.get_children()[idx-1]
            self.flow_tree.selection_set(new_id)
            self.status_var.set("已上移步骤")
    
    def move_down_step(self):
        if self.flow_running:
            messagebox.showwarning("警告", "请先停止当前流程")
            return
        selected = self.flow_tree.selection()
        if not selected:
            return
        idx = self.flow_tree.index(selected[0])
        if idx < len(self.flow_steps) - 1:
            self.flow_steps[idx], self.flow_steps[idx+1] = self.flow_steps[idx+1], self.flow_steps[idx]
            self.refresh_flow_tree()
            new_id = self.flow_tree.get_children()[idx+1]
            self.flow_tree.selection_set(new_id)
            self.status_var.set("已下移步骤")
    
    def clear_flow(self):
        if self.flow_running:
            messagebox.showwarning("警告", "请先停止当前流程")
            return
        if messagebox.askyesno("确认", "清空所有步骤？"):
            self.flow_steps.clear()
            self.refresh_flow_tree()
            self.status_var.set("已清空流程")
    
    def run_flow(self):
        if not self.flow_steps:
            messagebox.showinfo("提示", "流程为空，请先添加步骤")
            return
        if self.clicking or self.recording or self.replaying:
            messagebox.showwarning("警告", "请先停止其他操作")
            return
        self.flow_vars = {}
        self.flow_running = True
        self.stop_flag.clear()
        self.run_flow_btn.config(state="disabled")
        self.stop_flow_btn.config(state="normal")
        self.status_var.set("流程执行中...")
        self.flow_thread = threading.Thread(target=self.run_flow_loop, daemon=True)
        self.flow_thread.start()
    
    def run_flow_loop(self):
        steps = self.flow_steps
        idx = 0
        max_iter = 1000
        it = 0
        anti_detect = self.anti_detect_flow.get()
        while self.flow_running and not self.stop_flag.is_set() and idx < len(steps) and it < max_iter:
            it += 1
            step = steps[idx]
            try:
                if step["type"] == "move":
                    x, y = step["x"], step["y"]
                    if anti_detect:
                        x, y = self.apply_position_jitter(x, y, True)
                    self.mouse_ctrl.position = (x, y)
                    idx += 1
                elif step["type"] == "click":
                    x, y = step["x"], step["y"]
                    if anti_detect:
                        x, y = self.apply_position_jitter(x, y, True)
                    self.mouse_ctrl.position = (x, y)
                    self.mouse_ctrl.click(self._str_to_button(step["button"]), 1)
                    idx += 1
                elif step["type"] == "doubleclick":
                    x, y = step["x"], step["y"]
                    if anti_detect:
                        x, y = self.apply_position_jitter(x, y, True)
                    self.mouse_ctrl.position = (x, y)
                    btn = self._str_to_button(step["button"])
                    self.mouse_ctrl.click(btn, 1)
                    time.sleep(0.1)
                    self.mouse_ctrl.click(btn, 1)
                    idx += 1
                elif step["type"] == "wait":
                    wait = step["seconds"]
                    start = time.time()
                    while time.time() - start < wait:
                        if not self.flow_running or self.stop_flag.is_set():
                            break
                        time.sleep(0.05)
                    idx += 1
                elif step["type"] == "ocr_win":
                    if not WINDOWS_OCR_AVAILABLE:
                        raise Exception("Windows OCR 不可用")
                    import asyncio
                    async def recognize():
                        from winrt.windows.media.ocr import OcrEngine
                        from winrt.windows.globalization import Language
                        from winrt.windows.graphics.imaging import BitmapDecoder, SoftwareBitmap
                        from winrt.windows.storage.streams import InMemoryRandomAccessStream, Buffer
                        img = ImageGrab.grab(bbox=(step["x1"], step["y1"], step["x2"], step["y2"]))
                        from io import BytesIO
                        img_bytes = BytesIO()
                        img.save(img_bytes, format='PNG')
                        img_bytes.seek(0)
                        data = img_bytes.read()
                        stream = InMemoryRandomAccessStream()
                        writer = stream.get_output_stream()
                        buffer = Buffer(len(data))
                        buffer.length = len(data)
                        import ctypes
                        ctypes.memmove(buffer._buffer, data, len(data))
                        await writer.write_async(buffer)
                        await writer.flush_async()
                        decoder = await BitmapDecoder.create_async(stream)
                        software_bitmap = await decoder.get_software_bitmap_async()
                        lang = Language(step["lang"])
                        engine = OcrEngine.try_create_from_language(lang)
                        if engine is None:
                            return ""
                        result = await engine.recognize_async(software_bitmap)
                        return result.text
                    try:
                        loop = asyncio.new_event_loop()
                        asyncio.set_event_loop(loop)
                        text = loop.run_until_complete(recognize())
                        loop.close()
                        var_name = step.get("var", "ocr_result")
                        self.flow_vars[var_name] = text.strip()
                        self.status_var.set(f"OCR 结果: {text[:50]}")
                    except Exception as e:
                        self.status_var.set(f"OCR 出错: {e}")
                    idx += 1
                elif step["type"] == "image_match":
                    img_path = step["image_path"]
                    target = Image.open(img_path)
                    region = step["region"]
                    threshold = step["threshold"]
                    if region:
                        screenshot = ImageGrab.grab(bbox=region)
                    else:
                        screenshot = ImageGrab.grab()
                    try:
                        import cv2
                        import numpy as np
                        screen_np = np.array(screenshot)
                        target_np = np.array(target)
                        result = cv2.matchTemplate(screen_np, target_np, cv2.TM_CCOEFF_NORMED)
                        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
                        if max_val >= threshold:
                            h, w = target_np.shape[:2]
                            x = max_loc[0] + w//2
                            y = max_loc[1] + h//2
                            if region:
                                x += region[0]
                                y += region[1]
                            self.flow_vars[step["var"]] = f"{x},{y}"
                        else:
                            self.flow_vars[step["var"]] = ""
                    except ImportError:
                        raise Exception("图像匹配需要安装 opencv-python: pip install opencv-python")
                    idx += 1
                elif step["type"] == "color_detect":
                    region = (step["x1"], step["y1"], step["x2"], step["y2"])
                    screenshot = ImageGrab.grab(bbox=region)
                    target_color = step["color"]
                    tolerance = step["tolerance"]
                    found = False
                    pixels = screenshot.load()
                    width, height = screenshot.size
                    for x in range(width):
                        for y in range(height):
                            pixel = pixels[x, y]
                            if abs(pixel[0] - target_color[0]) <= tolerance and \
                               abs(pixel[1] - target_color[1]) <= tolerance and \
                               abs(pixel[2] - target_color[2]) <= tolerance:
                                found = True
                                break
                        if found:
                            break
                    self.flow_vars[step["var"]] = found
                    idx += 1
                elif step["type"] == "branch":
                    var_val = self.flow_vars.get(step["var"], "")
                    op = step["op"]
                    compare_val = step["value"]
                    condition = False
                    if op == "contains":
                        condition = compare_val in str(var_val)
                    elif op == "equals":
                        condition = str(var_val) == compare_val
                    elif op == "not empty":
                        condition = bool(var_val)
                    elif op == "not empty string":
                        condition = var_val != ""
                    if condition:
                        target = step["target_step"] - 1
                        if 0 <= target < len(steps):
                            idx = target
                        else:
                            break
                    else:
                        idx += 1
                else:
                    idx += 1
                self.status_var.set(f"流程执行中... 步骤 {idx}/{len(steps)}")
            except Exception as e:
                self.status_var.set(f"流程执行出错: {e}")
                break
        if self.flow_running and not self.stop_flag.is_set():
            self.status_var.set("流程执行完成")
        self.flow_running = False
        self.run_flow_btn.config(state="normal")
        self.stop_flow_btn.config(state="disabled")
        self.stop_flag.clear()
    
    def stop_flow(self):
        if self.flow_running:
            self.flow_running = False
            self.stop_flag.set()
            self.status_var.set("流程已停止")
            self.run_flow_btn.config(state="normal")
            self.stop_flow_btn.config(state="disabled")
    
    def _str_to_button(self, btn_str):
        if btn_str == "left":
            return Button.left
        elif btn_str == "right":
            return Button.right
        elif btn_str == "middle":
            return Button.middle
        return Button.left
    
    def on_key_press(self, key):
        try:
            if key == keyboard.Key.f6:
                self.after(0, self.toggle_clicking)
            elif key == keyboard.Key.f9:
                self.after(0, self.toggle_recording)
            elif key == keyboard.Key.f10:
                if self.replaying:
                    self.after(0, self.stop_replay)
                elif self.recorded_actions:
                    self.after(0, self.start_replay)
                else:
                    self.after(0, lambda: messagebox.showinfo("提示", "没有可回放的操作"))
        except AttributeError:
            pass
    
    def on_close(self):
        self.clicking = False
        self.recording = False
        self.replaying = False
        self.flow_running = False
        self.stop_flag.set()
        if hasattr(self, 'mouse_listener') and self.mouse_listener.running:
            self.mouse_listener.stop()
        self.listener.stop()
        self.destroy()

if __name__ == "__main__":
    app = AutoClickerApp()
    app.mainloop()