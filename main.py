import tkinter as tk
import sys
import time
import ctypes
import threading
from tkinter import messagebox
import win32gui
import win32api
import win32con
from pywinauto import Application
from nobunaga_automation import NobunagaAutomation, NobunagaStateCheck, NobunagaAction
import crafting_logic

class TextRedirector:
    """將 sys.stdout 重新導向至 Tkinter Text 元件的輔助類別"""
    def __init__(self, widget):
        self.widget = widget

    def write(self, str_msg):
        # 使用 after 確保在主執行緒更新 UI，避免多執行緒衝突
        self.widget.after(0, self._insert_text, str_msg)

    def _insert_text(self, str_msg):
        self.widget.configure(state='normal')
        self.widget.insert(tk.END, str_msg)
        self.widget.see(tk.END) # 自動捲動到最下方
        self.widget.configure(state='disabled')

    def flush(self):
        pass

class NobunagaToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("信長Online 輔助小工具")
        self._adjust_window_position()
        
        # 初始化核心邏輯類別
        self.auto = NobunagaAutomation()
        self.windows_data = []  # 存儲 [(hwnd, title, x, y), ...]
        self.state_check = NobunagaStateCheck()
        self.nobu_action = NobunagaAction()
        
        # Tooltip 相關變數
        self.tooltip_window = None
        self.last_hover_index = -1

        # 用於中斷執行緒的事件
        self.stop_event = threading.Event()
        self.task_thread = None

        # 冥宮掛機樓層數
        self.floor_display_str = tk.StringVar(value="已戰鬥: 0 次")
        self.time_display_str = tk.StringVar(value="最後更新: -")
        self.debug_mode_var = tk.BooleanVar(value=False)
        self.item_use_var = tk.BooleanVar(value=False)

        self._setup_ui()
        self.refresh_windows()
        
        # 重新導向 print 輸出
        sys.stdout = TextRedirector(self.txt_logs)
        sys.stderr = TextRedirector(self.txt_logs) # 同時捕獲錯誤訊息

    def _adjust_window_position(self):
        """檢查系統中是否存在信長視窗，並根據其位置調整工具啟動座標以避免重疊"""
        width, height = 750, 650
        # 預設啟動位置
        target_x, target_y = 100, 100

        # 尋找信長之野望的主視窗 (ClassName 為 Nobunaga Online Game MainFrame)
        hwnd_game = win32gui.FindWindow("Nobunaga Online Game MainFrame", None)

        # 判斷視窗是否存在、是否可見、以及是否並非最小化 (Iconic 代表最小化)
        if hwnd_game and win32gui.IsWindowVisible(hwnd_game) and not win32gui.IsIconic(hwnd_game):
            try:
                rect = win32gui.GetWindowRect(hwnd_game)
                g_left, g_top, g_right, g_bottom = rect
                screen_w = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
                # 優先嘗試放在遊戲視窗右邊 (留 10 像素間距)
                if g_right + width + 10 <= screen_w:
                    target_x = g_right + 10
                    target_y = g_top
                # 如果右邊沒空間，則嘗試放在遊戲視窗左邊
                elif g_left - width - 10 >= 0:
                    target_x = g_left - width - 10
                    target_y = g_top
            except Exception as e:
                print(f"自動調整視窗位置失敗: {e}")

        self.root.geometry(f"{width}x{height}+{target_x}+{target_y}")

    def _setup_ui(self):
        # 上方主要容器
        frame_top = tk.Frame(self.root)
        frame_top.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # 左側清單
        frame_left = tk.LabelFrame(frame_top, text="遊戲視窗列表", padx=10, pady=10)
        frame_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.lb_windows = tk.Listbox(frame_left, font=("Microsoft JhengHei", 10), exportselection=False)
        # 綁定滑鼠移動與離開事件
        self.lb_windows.bind("<Motion>", self._on_listbox_motion)
        self.lb_windows.bind("<Leave>", self._hide_tooltip)
        
        self.lb_windows.pack(fill=tk.BOTH, expand=True)
        
        btn_refresh = tk.Button(frame_left, text="重新整理視窗", command=self.refresh_windows)
        btn_refresh.pack(fill=tk.X, pady=5)

        # 右側功能
        frame_right = tk.LabelFrame(frame_top, text="功能清單", padx=10, pady=10)
        frame_right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        self.lb_features = tk.Listbox(frame_right, font=("Microsoft JhengHei", 10), exportselection=False)
        features = ["稼業連點", "冥宮掛機", "跟隨戰鬥", "測試選項一", "測試選項二"]
        for f in features:
            self.lb_features.insert(tk.END, f)
        self.lb_features.pack(fill=tk.BOTH, expand=True)
        
        btn_run = tk.Button(frame_right, text="執行功能", command=self.run_feature, bg="#e1e1e1")
        btn_run.pack(fill=tk.X, pady=5)

        btn_stop = tk.Button(frame_right, text="停止執行", command=self.stop_feature, bg="#ffcccc")
        btn_stop.pack(fill=tk.X)

        # 偵錯模式開關
        self.chk_debug_mode = tk.Checkbutton(frame_right, text="偵錯模式 (顯示影像搜尋)", variable=self.debug_mode_var, command=self._toggle_debug_mode)
        self.chk_debug_mode.pack(fill=tk.X, pady=5)

        # 物品使用開關
        self.chk_item_use = tk.Checkbutton(frame_right, text="跟隨戰鬥：自動使用物品", variable=self.item_use_var)
        self.chk_item_use.pack(fill=tk.X, pady=5)

        # 冥宮掛機進度顯示
        self.lbl_progress = tk.Label(frame_right, textvariable=self.floor_display_str, font=("Microsoft JhengHei", 12, "bold"), fg="blue")
        self.lbl_progress.pack(fill=tk.X, pady=(10, 0))

        self.lbl_time = tk.Label(frame_right, textvariable=self.time_display_str, font=("Microsoft JhengHei", 9), fg="gray")
        self.lbl_time.pack(fill=tk.X, pady=(0, 10))

        # 下方日誌區域
        frame_log = tk.LabelFrame(self.root, text="執行日誌", padx=5, pady=5)
        frame_log.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=False)

        self.txt_logs = tk.Text(frame_log, font=("Microsoft JhengHei", 10), state='disabled', bg="#f0f0f0", height=12)
        scrollbar = tk.Scrollbar(frame_log, command=self.txt_logs.yview)
        self.txt_logs.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_logs.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def refresh_windows(self):
        """利用 EnumWindows 過濾信長視窗"""
        self.lb_windows.delete(0, tk.END)
        self.windows_data = []
        
        def enum_handler(hwnd, lparam):
            if win32gui.IsWindowVisible(hwnd):
                # 取得視窗類別名稱進行精確過濾
                class_name = win32gui.GetClassName(hwnd)
                if class_name == "Nobunaga Online Game MainFrame":
                    title = win32gui.GetWindowText(hwnd)
                    rect = win32gui.GetWindowRect(hwnd)
                    # 紀錄 hwnd, title, x, y
                    self.windows_data.append((hwnd, title, rect[0], rect[1]))

        win32gui.EnumWindows(enum_handler, None)
        
        for win in self.windows_data:
            self.lb_windows.insert(tk.END, f"視窗: {win[1]} (X:{win[2]}, Y:{win[3]})")

    def run_feature(self):
        # 取得選中的視窗與功能
        win_idx = self.lb_windows.curselection()
        feat_idx = self.lb_features.curselection()
        print(f"選中的視窗索引: {win_idx}, 選中的功能索引: {feat_idx}")


        if not win_idx or not feat_idx:
            messagebox.showwarning("提示", "請同時選擇一個視窗與一個功能")
            return

        target_win = self.windows_data[win_idx[0]]
        hwnd = target_win[0]
        feature_name = self.lb_features.get(feat_idx[0])

        # 檢查目前是否有正在執行的任務，若有則先結束它
        if self.task_thread and self.task_thread.is_alive():
            print(f"偵測到任務仍在執行，正在請求停止並排程啟動: {feature_name}")
            self.stop_feature()
            # 使用非阻塞方式等待並啟動
            self._wait_for_thread_and_run(hwnd, target_win[1], feature_name)
        else:
            self._start_task_thread(hwnd, target_win[1], feature_name)

    def _wait_for_thread_and_run(self, hwnd, title, feature_name):
        """非阻塞等待執行緒結束後再啟動"""
        if self.task_thread and self.task_thread.is_alive():
            # 每 100 毫秒檢查一次，主迴圈不會被卡死
            self.root.after(100, lambda: self._wait_for_thread_and_run(hwnd, title, feature_name))
        else:
            print(f"舊執行緒已成功釋放，啟動新功能: {feature_name}")
            self._start_task_thread(hwnd, title, feature_name)

    def _start_task_thread(self, hwnd, title, feature_name):
        """重設訊號並正式啟動執行緒"""
        # 重設停止訊號
        self.stop_event.clear()

        self.task_thread = threading.Thread(
            target=self._execute_feature_logic, 
            args=(hwnd, title, feature_name), 
            daemon=True
        )
        self.task_thread.start()

    def stop_feature(self):
        """發送停止訊號給背景執行緒"""
        self.stop_event.set()
        print("已發送停止指令...")

    def _execute_feature_logic(self, hwnd, title, feature_name):
        """在背景執行個項功能的具體邏輯"""
        try:
            print(f"--- 啟動背景任務: {feature_name} (視窗: {title}) ---")
            
            # pywinauto 連接 (放在 thread 中，避免連接超時造成 UI 停頓)
            try:
                app = Application().connect(handle=hwnd)
            except Exception as e:
                print(f"pywinauto 連接失敗: {e}")

            if feature_name == "稼業連點":
                # 呼叫獨立的稼業邏輯
                # template_path 可以改成你實際的圖檔路徑，例如 "finish_icon.png"
                crafting_logic.start_crafting_loop(
                    hwnd, self.auto, self.stop_event, template_path='./img/材料不夠.png'
                )
            elif feature_name == "冥宮掛機":
                # 呼叫獨立的冥宮邏輯
                # 重置樓層數
                self._update_dungeon_floor(0)
                crafting_logic.dream_dungeon_loop(
                    hwnd, self.auto, self.state_check,self.nobu_action, self.stop_event,
                    update_floor_callback=self._update_dungeon_floor
                )
            elif feature_name == "跟隨戰鬥":
                # 呼叫跟隨戰鬥邏輯，傳入 UI 上的勾選狀態
                crafting_logic.follow_combat_loop(
                    hwnd, self.auto, self.state_check, self.nobu_action, self.stop_event,
                    update_floor_callback=self._update_dungeon_floor,
                    item_use=self.item_use_var.get()
                )


            elif feature_name == "測試選項一":
                # 範例：發送 a 鍵 (0x41)
                #self.auto.send_key(hwnd, 0x41, hold_time=0.2)
                # 0x53 是 'S' 的虛擬鍵碼
                # win32con.VK_SHIFT 是 Shift 的鍵碼
                #self.auto.send_key(hwnd, 0x53, modifiers=[win32con.VK_SHIFT])
                self.nobu_action.menu_team_hero_select(hwnd, self.auto, hero_team_index=3)

                print(f"[{feature_name}] team_hero_select 完成")
            elif feature_name == "測試選項二":
                self.nobu_action.move_head_north(hwnd, self.auto)
                #if self.auto.find_image_click(hwnd, 'img/YN_確定.png'):
                print(f"[{feature_name}] move_head_north 完成")

            else:
                # 預留其他功能的實作空間
                while not self.stop_event.is_set():
                    print(f"[{feature_name}] 正在背景執行中...")
                    time.sleep(1)
                
        except Exception as e:
            print(f"執行功能時發生錯誤: {e}")

    def _update_dungeon_floor(self, floor_number):
        """更新主畫面顯示的冥宮樓層數"""
        self.floor_display_str.set(f"已戰鬥: {floor_number} 次")
        current_time = time.strftime("%Y-%m-%d, %H:%M:%S")
        self.time_display_str.set(f"最後更新: {current_time}")

    def _toggle_debug_mode(self):
        """切換自動化工具的偵錯模式"""
        self.auto.debug_mode = self.debug_mode_var.get()
        print(f"偵錯模式已設定為: {self.auto.debug_mode}")

    def _on_listbox_motion(self, event):
        """當滑鼠在 Listbox 上移動時，判斷是否顯示 Tooltip"""
        index = self.lb_windows.nearest(event.y)
        
        # 如果索引改變了，先隱藏舊的 Tooltip
        if index != self.last_hover_index:
            self._hide_tooltip()
            self.last_hover_index = index
            
            # 檢查滑鼠是否真的在該項目的範圍內 (避免 Listbox 空白處也觸發)
            bbox = self.lb_windows.bbox(index)
            if bbox and bbox[1] <= event.y <= bbox[1] + bbox[3]:
                full_text = self.lb_windows.get(index)
                self._show_tooltip(event.x_root + 15, event.y_root + 10, full_text)

    def _show_tooltip(self, x, y, text):
        """建立並顯示 Tooltip 視窗"""
        if self.tooltip_window or not text:
            return
            
        self.tooltip_window = tk.Toplevel(self.root)
        # 移除視窗邊框與標題列
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")
        
        label = tk.Label(self.tooltip_window, text=text, justify=tk.LEFT,
                         background="#ffffca", relief=tk.SOLID, borderwidth=1,
                         font=("Microsoft JhengHei", 9), padx=3, pady=3)
        label.pack()

    def _hide_tooltip(self, event=None):
        """隱藏 Tooltip 視窗"""
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None
        self.last_hover_index = -1

def is_admin():
    """檢查是否具有管理員權限"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if __name__ == "__main__":
    if is_admin():
        # 強制程序進入 DPI Aware 模式，解決座標偏移問題
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()

        # 隱藏背景的終端機視窗 (如果目前是透過 python.exe 啟動)
        whnd = ctypes.windll.kernel32.GetConsoleWindow()
        if whnd != 0:
            ctypes.windll.user32.ShowWindow(whnd, 0)  # 0 代表 SW_HIDE
            
        root = tk.Tk()
        app = NobunagaToolApp(root)
        root.mainloop()
    else:
        # 重新啟動程式並要求管理員權限
        print("權限不足，嘗試取得管理員權限...")
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
