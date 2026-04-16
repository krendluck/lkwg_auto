import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import keyboard
import threading
import time
import os
import ctypes
import random
from PIL import Image, ImageTk

# Windows API 鼠标事件常量
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004

def mouse_left_press():
    """使用 Windows 原生 API 模拟鼠标左键按下"""
    ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)

def mouse_left_release():
    """使用 Windows 原生 API 模拟鼠标左键松开"""
    ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

class MouseMapperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("自动化助手 (按键映射 & 图像识别)")
        self.root.geometry("620x520")  # 扩大高度容纳时间设置
        self.root.resizable(False, False)
        
        # --- 共享事件/状态 ---
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # --- 选项卡控件 ---
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill='both', padx=2, pady=2)
        
        self.tab1 = ttk.Frame(self.notebook)
        self.tab2 = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab1, text="键盘映射鼠标")
        self.notebook.add(self.tab2, text="图像识别点击")
        
        # ==========================================
        # Tab 1: 键盘映射鼠标长按 (原有功能)
        # ==========================================
        self.kb_running = False
        self.is_mouse_held = False
        self.kb_listener_thread = None

        tk.Label(self.tab1, text="请输入或录制要绑定的按键", font=("Arial", 11, "bold")).pack(pady=(15, 5))
        tk.Label(self.tab1, text="(例如: 'a', 'space', 'ctrl+x', 'shift+a')").pack(pady=0)

        hk_frame = tk.Frame(self.tab1)
        hk_frame.pack(pady=10)
        
        self.hotkey_var = tk.StringVar(value="alt+x")
        self.hotkey_entry = tk.Entry(hk_frame, textvariable=self.hotkey_var, font=("Arial", 11), width=18, justify="center")
        self.hotkey_entry.pack(side=tk.LEFT, padx=5)

        self.btn_record = tk.Button(hk_frame, text="⏺ 录制", command=self.start_recording)
        self.btn_record.pack(side=tk.LEFT)

        self.kb_status_var = tk.StringVar()
        self.kb_status_var.set("状态: 未启动")
        self.kb_status_label = tk.Label(self.tab1, textvariable=self.kb_status_var, fg="gray", font=("Arial", 10))
        self.kb_status_label.pack(pady=5)

        kb_frame_btns = tk.Frame(self.tab1)
        kb_frame_btns.pack(pady=(5,10))

        self.kb_btn_start = tk.Button(kb_frame_btns, text="▶ 启动", width=12, bg="#d4edda", command=self.start_kb_mapping)
        self.kb_btn_start.pack(side=tk.LEFT, padx=10)

        self.kb_btn_stop = tk.Button(kb_frame_btns, text="⏹ 停止", width=12, state=tk.DISABLED, bg="#f8d7da", command=self.stop_kb_mapping)
        self.kb_btn_stop.pack(side=tk.LEFT, padx=10)
        
        # ==========================================
        # Tab 2: 图像识别点击 (新增框架)
        # ==========================================
        self.img_running = False
        self.img_thread = None
        self.target_image_path = ""
        
        # --- 窗口选择控件 ---
        win_pick_frame = tk.Frame(self.tab2)
        win_pick_frame.pack(pady=(15, 0))
        tk.Label(win_pick_frame, text="目标窗口:").pack(side=tk.LEFT)
        self.win_combo = ttk.Combobox(win_pick_frame, width=25, state="readonly")
        self.win_combo.pack(side=tk.LEFT, padx=5)
        tk.Button(win_pick_frame, text="刷新", command=self.refresh_windows).pack(side=tk.LEFT)

        # --- 区域选择控件 ---
        roi_frame = tk.Frame(self.tab2)
        roi_frame.pack(pady=5)
        tk.Label(roi_frame, text="识别区域:").pack(side=tk.LEFT)
        self.roi_var = tk.StringVar(value="相对窗口 72,123 90x402")
        tk.Entry(roi_frame, textvariable=self.roi_var, width=22, state="readonly").pack(side=tk.LEFT, padx=5)
        tk.Button(roi_frame, text="框选范围", command=self.snip_area).pack(side=tk.LEFT)
        self.roi_rect = (72, 123, 90, 402) # 设置默认相对窗口区域

        # --- 图片选择控件 ---
        img_pick_frame = tk.Frame(self.tab2)
        img_pick_frame.pack(pady=(5, 5))
        tk.Label(img_pick_frame, text="状态标志:").pack(side=tk.LEFT)
        
        self.img_path_var = tk.StringVar(value="status.png")
        self.target_image_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "status.png")
        if not os.path.exists(self.target_image_path):
             self.target_image_path = ""
             self.img_path_var.set("")

        tk.Entry(img_pick_frame, textvariable=self.img_path_var, width=20, state="readonly").pack(side=tk.LEFT, padx=5)
        tk.Button(img_pick_frame, text="浏览...", command=self.browse_image).pack(side=tk.LEFT)

        # --- 识别检测参数设置 (新增) ---
        settings_frame = tk.Frame(self.tab2)
        settings_frame.pack(pady=5)
        
        tk.Label(settings_frame, text="常规检测间隔(秒):").pack(side=tk.LEFT)
        self.detect_interval_var = tk.StringVar(value="0.5")
        tk.Entry(settings_frame, textvariable=self.detect_interval_var, width=5).pack(side=tk.LEFT, padx=3)
        
        tk.Label(settings_frame, text="全放出后轮换间隔(秒):").pack(side=tk.LEFT, padx=(15, 0))
        self.rotate_interval_var = tk.StringVar(value="5.0")
        tk.Entry(settings_frame, textvariable=self.rotate_interval_var, width=5).pack(side=tk.LEFT, padx=3)

        self.img_status_var = tk.StringVar()
        self.img_status_var.set("状态: 请框选区域并选择标志图片")
        self.img_status_label = tk.Label(self.tab2, textvariable=self.img_status_var, fg="gray", font=("Arial", 10))
        self.img_status_label.pack(pady=5)

        # --- 实时图像切割预览 (新增) ---
        self.preview_frame = tk.LabelFrame(self.tab2, text="实时图像切割预览 (排查 OCR 范围用)")
        self.preview_frame.pack(pady=2, fill=tk.X, padx=2)
        self.preview_labels = []
        for i in range(6):
            lbl = tk.Label(self.preview_frame, text=f"空闲{i+1}", bg="#333333", fg="white", width=10, height=4)
            lbl.pack(side=tk.LEFT, padx=3, pady=2)
            self.preview_labels.append(lbl)
        self.current_tk_images = [] # 防止图片被垃圾回收导致白板

        img_frame_btns = tk.Frame(self.tab2)
        img_frame_btns.pack(pady=2)

        self.img_btn_start = tk.Button(img_frame_btns, text="▶ 开始识别", width=12, bg="#d4edda", command=self.start_img_rec)
        self.img_btn_start.pack(side=tk.LEFT, padx=2)

        self.img_btn_stop = tk.Button(img_frame_btns, text="⏹ 停止", width=12, state=tk.DISABLED, bg="#f8d7da", command=self.stop_img_rec)
        self.img_btn_stop.pack(side=tk.LEFT, padx=2)
        
        # --- 全局快捷键提示 ---
        hotkey_tips_frame = tk.LabelFrame(self.tab2, text="全局快捷键 (任意界面有效)")
        hotkey_tips_frame.pack(pady=2, fill=tk.X, padx=2)
        tk.Label(hotkey_tips_frame, text="F8: 开始识别  |  F9: 停止识别  |  F10: 暂停/开始轮换", fg="#b8860b", font=("Arial", 10, "bold")).pack(pady=2)
        
        self.is_rotation_paused = False
        keyboard.add_hotkey('F8', lambda: self.root.after(0, self.start_img_rec))
        keyboard.add_hotkey('F9', lambda: self.root.after(0, self.stop_img_rec))
        keyboard.add_hotkey('F10', lambda: self.root.after(0, self.toggle_rotation_pause))

        # 初始化刷新一次窗口列表
        self.root.after(100, self.refresh_windows)

    # ==========================
    # Tab 1 逻辑
    # ==========================
    def start_recording(self):
        # 禁用各种UI，防止冲突
        self.btn_record.config(state=tk.DISABLED, text="录制中...")
        self.kb_btn_start.config(state=tk.DISABLED)
        self.hotkey_entry.config(state=tk.DISABLED)
        self.kb_status_var.set("状态: 请在键盘上按下你想要的快捷键或组合键...")
        self.kb_status_label.config(fg="blue")
        
        # 开启后台线程监听一次按键
        threading.Thread(target=self.record_thread, daemon=True).start()

    def record_thread(self):
        try:
            # 捕获并在按下时阻塞直到得到按键（不拦截按键原生功能）
            hotkey = keyboard.read_hotkey(suppress=False)
            self.root.after(0, self.finish_recording, hotkey)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"录制热键出错: {e}"))

    def finish_recording(self, hotkey):
        self.hotkey_var.set(hotkey)
        
        self.btn_record.config(state=tk.NORMAL, text="⏺ 录制")
        self.kb_btn_start.config(state=tk.NORMAL)
        self.hotkey_entry.config(state=tk.NORMAL)
        self.kb_status_var.set(f"状态: 成功录制按键 -> [{hotkey}]")
        self.kb_status_label.config(fg="green")

    def start_kb_mapping(self):
        hotkey = self.hotkey_var.get().strip()
        if not hotkey:
            messagebox.showwarning("警告", "按键不能为空！")
            return
        
        self.kb_running = True
        self.kb_btn_start.config(state=tk.DISABLED)
        self.btn_record.config(state=tk.DISABLED)
        self.hotkey_entry.config(state=tk.DISABLED)
        self.kb_btn_stop.config(state=tk.NORMAL)
        self.kb_status_var.set(f"状态: 正在监听 -> [{hotkey}]")
        self.kb_status_label.config(fg="green")

        self.kb_listener_thread = threading.Thread(target=self.kb_listen_loop, args=(hotkey,), daemon=True)
        self.kb_listener_thread.start()

    def stop_kb_mapping(self):
        self.kb_running = False
        if self.is_mouse_held:
            try:
                mouse_left_release()
            except Exception as e:
                print(f"释放鼠标失败: {e}")
            self.is_mouse_held = False
            
        self.kb_btn_start.config(state=tk.NORMAL)
        self.btn_record.config(state=tk.NORMAL)
        self.hotkey_entry.config(state=tk.NORMAL)
        self.kb_btn_stop.config(state=tk.DISABLED)
        self.kb_status_var.set("状态: 已停止")
        self.kb_status_label.config(fg="red")

    def kb_listen_loop(self, hotkey):
        try:
            while self.kb_running:
                if keyboard.is_pressed(hotkey):
                    if not self.is_mouse_held:
                        time.sleep(random.uniform(0.01, 0.04))
                        mouse_left_press()
                        self.is_mouse_held = True
                else:
                    if self.is_mouse_held:
                        time.sleep(random.uniform(0.01, 0.04))
                        mouse_left_release()
                        self.is_mouse_held = False
                time.sleep(0.01)
        except ValueError:
            self.root.after(0, lambda: messagebox.showerror("错误", f"无法识别按键 '{hotkey}'，请检查拼写。"))
            self.root.after(0, self.stop_kb_mapping)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"发生异常: {e}"))
            self.root.after(0, self.stop_kb_mapping)

    # ==========================
    # Tab 2 逻辑 (基于 PyAutoGUI + win32 窗口位置)
    # ==========================
    def get_windows_list(self):
        """获取当前可见的所有带标题的窗口列表"""
        titles = []
        EnumWindows = ctypes.windll.user32.EnumWindows
        EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
        GetWindowText = ctypes.windll.user32.GetWindowTextW
        GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
        IsWindowVisible = ctypes.windll.user32.IsWindowVisible

        def foreach_window(hwnd, lParam):
            if IsWindowVisible(hwnd):
                length = GetWindowTextLength(hwnd)
                if length > 0:
                    buff = ctypes.create_unicode_buffer(length + 1)
                    GetWindowText(hwnd, buff, length + 1)
                    titles.append(buff.value)
            return True
        
        EnumWindows(EnumWindowsProc(foreach_window), 0)
        return sorted(list(set(titles)))

    def refresh_windows(self):
        windows = self.get_windows_list()
        windows.insert(0, "全屏幕")
        
        # 强制把“洛克王国：世界”加入列表并默认选中锁定
        if "洛克王国：世界" not in windows:
            windows.append("洛克王国：世界")
            
        self.win_combo['values'] = windows
        self.win_combo.set("洛克王国：世界")

    def snip_area(self):
        """打开透明全屏层，供用户鼠标框选区域"""
        top = tk.Toplevel(self.root)
        top.attributes("-alpha", 0.3)
        top.attributes("-fullscreen", True)
        top.attributes("-topmost", True)
        top.config(cursor="cross")

        canvas = tk.Canvas(top, bg="gray")
        canvas.pack(fill=tk.BOTH, expand=True)

        self.snip_rect_id = None
        self.start_x = self.start_y = 0

        def on_press(event):
            self.start_x = event.x
            self.start_y = event.y
            if self.snip_rect_id:
                canvas.delete(self.snip_rect_id)
            self.snip_rect_id = canvas.create_rectangle(self.start_x, self.start_y, event.x, event.y, outline='red', width=2, fill='blue')

        def on_drag(event):
            canvas.coords(self.snip_rect_id, self.start_x, self.start_y, event.x, event.y)

        def on_release(event):
            end_x, end_y = event.x, event.y
            # 计算绝对坐标和宽高
            x, y = min(self.start_x, end_x), min(self.start_y, end_y)
            w, h = abs(self.start_x - end_x), abs(self.start_y - end_y)
            
            # --- 【修改为：锁定窗口，转换为相对坐标】 ---
            target_window = self.win_combo.get().strip()
            win_rect = self.get_window_rect(target_window)
            if win_rect:
                rel_x = x - win_rect[0]
                rel_y = y - win_rect[1]
                self.roi_rect = (rel_x, rel_y, w, h)
                self.roi_var.set(f"相对窗口 {rel_x},{rel_y} {w}x{h}")
            else:
                self.roi_rect = (x, y, w, h)
                self.roi_var.set(f"绝对屏幕 {x},{y} {w}x{h}")
                
            top.destroy()
            self.img_status_var.set(f"状态: 已框选关注区域 {self.roi_var.get()}")
            self.img_status_label.config(fg="blue")

        def on_escape(event):
            top.destroy()

        top.bind("<ButtonPress-1>", on_press)
        top.bind("<B1-Motion>", on_drag)
        top.bind("<ButtonRelease-1>", on_release)
        top.bind("<Escape>", on_escape)

    def browse_image(self):
        file_path = filedialog.askopenfilename(
            title="选择要识别的图片",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp")]
        )
        if file_path:
            self.target_image_path = file_path
            self.img_path_var.set(os.path.basename(file_path))
            self.img_status_var.set("成功加载图片，等待启动。")
            self.img_status_label.config(fg="blue")

    def start_img_rec(self):
        if getattr(self, "img_running", False):
            return
            
        if not self.target_image_path:
            messagebox.showwarning("警告", "请先选择一张要识别的图片！")
            return
            
        target_window = self.win_combo.get().strip()
        if not target_window:
            messagebox.showwarning("警告", "请先选择目标窗口名称！")
            return
            
        self.img_running = True
        self.img_btn_start.config(state=tk.DISABLED)
        self.img_btn_stop.config(state=tk.NORMAL)
        self.img_status_var.set(f"状态: 正在 [{target_window}] 中寻找图片...")
        self.img_status_label.config(fg="green")
        
        self.img_thread = threading.Thread(target=self.img_rec_loop, args=(target_window,), daemon=True)
        self.img_thread.start()

    def stop_img_rec(self):
        self.img_running = False
        self.img_btn_start.config(state=tk.NORMAL)
        self.img_btn_stop.config(state=tk.DISABLED)
        self.img_status_var.set("状态: 图像识别已停止")
        self.img_status_label.config(fg="red")

    def toggle_rotation_pause(self):
        self.is_rotation_paused = not getattr(self, "is_rotation_paused", False)
        status = "已暂停轮换" if self.is_rotation_paused else "已恢复轮换"
        self.img_status_var.set(f"状态: {status}")
        self.img_status_label.config(fg="purple")

    def update_previews(self, pil_images):
        """将扫描出来的分割图像实时推送到 UI 显示"""
        self.current_tk_images.clear()
        for i, lbl in enumerate(self.preview_labels):
            if i < len(pil_images):
                img = pil_images[i].copy()
                # 调整到框的大小以供观看
                img.thumbnail((80, 80), Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.ANTIALIAS)
                tk_img = ImageTk.PhotoImage(img)
                self.current_tk_images.append(tk_img)
                lbl.config(image=tk_img, text="", width=80, height=80)
            else:
                lbl.config(image="", text=f"空闲{i+1}", width=12, height=5)

    def get_window_rect(self, title):
        """获取目标窗口的坐标范围: (left, top, width, height)"""
        if title == "全屏幕":
            w = ctypes.windll.user32.GetSystemMetrics(0)
            h = ctypes.windll.user32.GetSystemMetrics(1)
            return (0, 0, w, h)
            
        import ctypes.wintypes # 把 import 移到函数最上面，防止作用域冲突
        
        # 1. 先尝试精确匹配
        hwnd = ctypes.windll.user32.FindWindowW(None, title)
        
        # 2. 如果精确匹配失败，尝试模糊（包含）匹配（应对游戏标题后面随时加了 FPS/Ping 之类的动态后缀的情况）
        if not hwnd:
            EnumWindows = ctypes.windll.user32.EnumWindows
            EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
            GetWindowText = ctypes.windll.user32.GetWindowTextW
            GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
            IsWindowVisible = ctypes.windll.user32.IsWindowVisible

            found_hwnd = 0
            def foreach_window(h, lParam):
                nonlocal found_hwnd
                if IsWindowVisible(h):
                    length = GetWindowTextLength(h)
                    if length > 0:
                        buff = ctypes.create_unicode_buffer(length + 1)
                        GetWindowText(h, buff, length + 1)
                        if title in buff.value:  # 只要窗口真实名字【包含】你选的名字就算数
                            found_hwnd = h
                            return False  # 找到了就停止遍历
                return True
            
            EnumWindows(EnumWindowsProc(foreach_window), 0)
            hwnd = found_hwnd

        if hwnd:
            rect = ctypes.wintypes.RECT()
            ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
            return (rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top)
        return None

    def img_rec_loop(self, target_window):
        import pyautogui  # 延迟导入，防止程序无库直接崩溃
        
        cycle_idx = 2  # 当所有精灵均已放出时，用于循环切换的内部状态 (2~6)

        while self.img_running:
            try:
                # 获取当前窗口的最新实时位置
                win_rect = self.get_window_rect(target_window)
                
                if win_rect and win_rect[2] > 0 and win_rect[3] > 0:
                    # 1. 锁定窗口跟随：把我们之前保存的相对框选位置，贴合到窗口最新的实时位置上
                    if self.roi_rect:
                        # 如果没有绑定过特定窗口，我们取之前保存的绝对位置
                        if self.roi_var.get().startswith("绝对"):
                            rect = self.roi_rect
                        else:
                            rect = (
                                win_rect[0] + self.roi_rect[0], 
                                win_rect[1] + self.roi_rect[1], 
                                self.roi_rect[2], 
                                self.roi_rect[3]
                            )
                    else:
                        rect = win_rect

                    # 【调试0】：保存整个搜索大区域的截图 (已调好，暂时注释)
                    # try:
                    #     region_img = pyautogui.screenshot(region=rect)
                    #     region_img.save("debug_0_search_region.png")
                    # except Exception:
                    #     pass

                    # 识别所有的“放出标志”（可能会有多个），加上 list() 避免生成器枯竭
                    # 将容错率从 0.8 适当下调到 0.7，防止有的图标背景稍微变动一点点就识别不到
                    try:
                        marks = list(pyautogui.locateAllOnScreen(self.target_image_path, region=rect, confidence=0.7))
                    except Exception as loc_err:
                        if "could not locate" in str(loc_err).lower() or "not found" in str(loc_err).lower() or "imagenotfound" in str(loc_err).lower():
                            marks = []
                        else:
                            raise loc_err
                    
                    if marks:
                        # 【重要修理】：pyautogui 对于同一个图片可能会返回多个挨得极近的坐标，需要进行去重
                        filtered_marks = []
                        for m in marks:
                            c = pyautogui.center(m)
                            is_duplicate = False
                            for fm in filtered_marks:
                                fc = pyautogui.center(fm)
                                # 距离缩小到 5 像素（极为严格，只有极其紧贴在同一点的才算作重复）
                                if abs(c.x - fc.x) < 5 and abs(c.y - fc.y) < 5:
                                    is_duplicate = True
                                    break
                            if not is_duplicate:
                                filtered_marks.append(m)

                        print(f"👀 [DEBUG] 原始扫描找到 {len(marks)} 个图块，去重后剩余 {len(filtered_marks)} 个目标")

                        results = []
                        preview_images = [] # 这里存放要发送给UI预览的图像
                        # 遍历每一个去重后的放出标志，并加上索引 idx 避免多精灵截图互被覆盖
                        for idx, mark in enumerate(filtered_marks):
                            # 【调试1】：保存找到的“放出”图标本身 (已调好，暂时注释)
                            # try:
                            #     icon_img = pyautogui.screenshot(region=(int(mark.left), int(mark.top), int(mark.width), int(mark.height)))
                            #     icon_img.save(f"debug_1_found_icon_{idx}.png")
                            # except Exception: pass

                            center = pyautogui.center(mark)
                            
                            # ==========================================
                            # 联合 OCR 动态截取左侧对应的【数字编号】区域
                            # ==========================================
                            # 逻辑：以“放出标志”的中心坐标为基准，往左推算数字框的位置
                            # 这样不受外层框选区域大小的影响。如有偏差，请微调下面这几个数字
                            
                            num_box_x = int(center.x) - 58    # 原来-80太偏左了，改为-65往右边平移一点
                            num_box_y = int(center.y) + 12    # 原来-25太靠上了，改为-5往下平移，避免数字底下被切掉
                            num_box_w = 12                    # 稍微收窄一点点
                            num_box_h = 17                    # 高度加一点点，确保包裹完整
                            
                            number_text = "?"
                            try:
                                import pytesseract
                                import re
                                pytesseract.pytesseract.tesseract_cmd = r'D:\\Program Files\\Tesseract-OCR\\tesseract.exe'
                                
                                num_img = pyautogui.screenshot(region=(num_box_x, num_box_y, num_box_w, num_box_h))
                                
                                import cv2
                                import numpy as np
                                
                                # ====== 极简图像处理 ======
                                # 转灰度 -> 放大 -> 留大白边
                                open_cv_image = np.array(num_img)[:, :, ::-1].copy()
                                gray = cv2.cvtColor(open_cv_image, cv2.COLOR_BGR2GRAY)
                                scaled = cv2.resize(gray, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
                                
                                # 加个白边防止 OCR 边缘被切，不加额外滤镜直接上
                                final_img = cv2.copyMakeBorder(scaled, 20, 20, 20, 20, cv2.BORDER_CONSTANT, value=255)
                                # cv2.imwrite(f"debug_2_ocr_target_{idx}.png", final_img)  # 调试：保存送给 OCR 的图 (已调好，暂时注释)
                                
                                # 用最快的单行模式直接识别
                                raw_text = pytesseract.image_to_string(final_img, config='--psm 7').strip()
                                
                                # 极简纠错截取数字
                                corrected = raw_text.replace('l', '1').replace('I', '1').replace('O', '0').replace('o', '0').replace('S', '5').replace('s', '5').replace('z', '2').replace('Z', '2')
                                digits = re.sub(r'\D', '', corrected)
                                
                                if digits:
                                    number_text = digits
                                    print(f"👉 [3 目标{idx}] 识别成功: {digits}")
                                else:
                                    # 如果最简单的模式没认出，说明背景太暗，退化到“二值化(反色)”试最后一次
                                    _, thresh_inv = cv2.threshold(scaled, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
                                    final_img2 = cv2.copyMakeBorder(thresh_inv, 20, 20, 20, 20, cv2.BORDER_CONSTANT, value=255)
                                    raw_text2 = pytesseract.image_to_string(final_img2, config='--psm 8').strip()
                                    corrected2 = raw_text2.replace('l', '1').replace('I', '1').replace('O', '0').replace('o', '0').replace('S', '5').replace('s', '5').replace('z', '2').replace('Z', '2')
                                    digits2 = re.sub(r'\D', '', corrected2)
                                    
                                    if digits2:
                                        number_text = digits2
                                        print(f"👉 [DEBUG 目标{idx}] (兜底方案) 识别成功: {digits2}")
                                    else:
                                        print(f"❌ [DEBUG 目标{idx}] 识别失败 (原图:'{raw_text}', 兜底:'{raw_text2}')")
                                        number_text = "?"
                                
                                # -- 新增：把处理发送给 OCR 的黑白大图转化拿出来让 UI 进行展示体验
                                if 'final_img' in locals() and final_img is not None:
                                    preview_images.append(Image.fromarray(final_img))
                                    
                            except Exception as e:
                                print(f"❌ [DEBUG 目标{idx}] OCR 运行报错: {str(e)}")
                                number_text = "OCR出错"

                            results.append(f"[{number_text}号]")
                            
                        # 更新UI显示结果
                        res_str = " ".join(results)
                        self.root.after(0, lambda s=res_str: self.img_status_var.set(f"状态: 检测到特征 -> {s}"))
                        # 把图片推给主线程去渲染显示 (调试完毕，先注释)
                        # self.root.after(0, self.update_previews, preview_images)

                    else:
                        results = []
                        self.root.after(0, lambda: self.img_status_var.set("状态: 未找到放出标志，默认全部未放出..."))
                        # self.root.after(0, self.update_previews, []) # 不在场时清空预览图

                    # -- 自动切换到最小未放出的精灵序号，排除1号 (无论是否找到均执行) --
                    # 识别出来的 results 是“已经放出”的精灵序号 (若未找到任何，则results为空)
                    released_nums = set()
                    for res in results:
                        num_match = re.search(r'\d+', res)
                        if num_match:
                            released_nums.add(int(num_match.group()))
                            
                    # 洛克王国队伍通常是6只精灵。我们要找未放出的精灵，即在 2~6 范围内但不在 released_nums 中的数字
                    unreleased_nums = [n for n in range(2, 7) if n not in released_nums]
                                
                    # 读取当前UI设置的轮询间隔
                    try:
                        detect_interval = max(0.1, float(self.detect_interval_var.get()))
                        rotate_interval = max(0.5, float(self.rotate_interval_var.get()))
                    except ValueError:
                        detect_interval = 0.5
                        rotate_interval = 5.0
                                
                    if unreleased_nums:
                        min_num = min(unreleased_nums)
                        print(f"✅ [操作执行] 找到最小未放出序号: {min_num}，将按下主键盘数字按键 '{min_num}'")
                        # 改用底层 keyboard 库并增加短暂延时(加上随机噪声)，确保游戏能识别到主键盘的数字键
                        time.sleep(random.uniform(0.02, 0.1)) # 按键前随机延迟
                        keyboard.press(str(min_num))
                        time.sleep(random.uniform(0.03, 0.08)) # 按下保持随机时长
                        keyboard.release(str(min_num))
                        
                        time.sleep(detect_interval + random.uniform(0.01, 0.05)) # 正常检测状态下的防频繁休眠 
                    else:
                        if getattr(self, "is_rotation_paused", False):
                            # 如果轮换暂停，只等待
                            pass
                        else:
                            print(f"✅ [操作执行] 2~6号全部精灵均已放出！进入间隔 {rotate_interval}s 的轮流切换循环，当前按下 '{cycle_idx}'")
                            time.sleep(random.uniform(0.02, 0.1)) # 按键前随机延迟
                            keyboard.press(str(cycle_idx))
                            time.sleep(random.uniform(0.03, 0.08))
                            keyboard.release(str(cycle_idx))
                            
                            # 自增 cycle_idx，如果大于6就从头回到2
                            cycle_idx += 1
                            if cycle_idx > 6:
                                cycle_idx = 2
                            
                        # 进行 rotate_interval 延时，切分成每块 0.5s 的小休眠，避免阻塞 UI 线程中途停止迟钝
                        wait_chunks = int(rotate_interval / 0.5)
                        for _ in range(wait_chunks):
                            if getattr(self, "img_running", False):
                                time.sleep(0.5)
                            else:
                                break
                        # 补充剩余的不足 0.5 的尾数时间
                        remainder = rotate_interval % 0.5
                        if remainder > 0 and getattr(self, "img_running", False):
                            time.sleep(remainder)

                else:
                    print(f"❌ [DEBUG 原理排查] 获取窗口坐标失败: target_window='{target_window}', win_rect={win_rect}")
                    self.root.after(0, lambda: self.img_status_var.set(f"状态: 无法定位目标或监控区域"))
                
            except Exception as e:
                err_str = str(e)
                # 兼容新版 pyautogui 找不到图会抛出异常的情况
                if "could not locate" in err_str.lower() or "not found" in err_str.lower() or "imagenotfound" in err_str.lower():
                    self.root.after(0, lambda: self.img_status_var.set("状态: 未找到放出标志，持续监控中..."))
                else:
                    self.root.after(0, lambda err=err_str: self.img_status_var.set(f"状态: 识别出错 ({err})，仍在重试..."))
                    print(f"❌ [DEBUG 原理排查] 外层全捕获异常: {err_str}")
                
            # 读取当前UI设置的轮询间隔
            try:
                detect_interval = max(0.1, float(self.detect_interval_var.get()))
            except ValueError:
                detect_interval = 0.5
            time.sleep(detect_interval)  # 每次扫描间隔时长，避免高CPU占用

    # ==========================
    # 全局控制
    # ==========================
    def on_closing(self):
        self.img_running = False
        self.kb_running = False
        self.stop_kb_mapping()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = MouseMapperApp(root)
    root.mainloop()