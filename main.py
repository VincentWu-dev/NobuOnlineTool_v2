import tkinter as tk
import sys
import time
import ctypes
import threading
from tkinter import messagebox
import win32gui
import win32con
from pywinauto import Application
from nobunaga_automation import NobunagaAutomation, NobunagaStateCheck, NobunagaAction
import crafting_logic

class NobunagaToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("信長Online 輔助小工具")
        self.root.geometry("700x450")
        
        # 初始化核心邏輯類別
        self.auto = NobunagaAutomation()
        self.windows_data = []  # 存儲 [(hwnd, title, x, y), ...]
        self.state_check = NobunagaStateCheck()
        self.nobu_action = NobunagaAction()
        
        # 用於中斷執行緒的事件
        self.stop_event = threading.Event()

        # 冥宮掛機樓層數
        self.floor_display_str = tk.StringVar(value="已戰鬥: 0 次")
        self.debug_mode_var = tk.BooleanVar(value=False)

        self._setup_ui()
        self.refresh_windows()

    def _setup_ui(self):
        # 左側清單
        frame_left = tk.LabelFrame(self.root, text="遊戲視窗列表", padx=10, pady=10)
        frame_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.lb_windows = tk.Listbox(frame_left, font=("Microsoft JhengHei", 10), exportselection=False)
        self.lb_windows.pack(fill=tk.BOTH, expand=True)
        
        btn_refresh = tk.Button(frame_left, text="重新整理視窗", command=self.refresh_windows)
        btn_refresh.pack(fill=tk.X, pady=5)

        # 右側功能
        frame_right = tk.LabelFrame(self.root, text="功能清單", padx=10, pady=10)
        frame_right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        self.lb_features = tk.Listbox(frame_right, font=("Microsoft JhengHei", 10), exportselection=False)
        features = ["稼業連點", "冥宮掛機", "測試選項一", "測試選項二"]
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

        # 冥宮掛機進度顯示
        self.lbl_progress = tk.Label(frame_right, textvariable=self.floor_display_str, font=("Microsoft JhengHei", 12, "bold"), fg="blue")
        self.lbl_progress.pack(fill=tk.X, pady=10)

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

        if not win_idx or not feat_idx:
            messagebox.showwarning("提示", "請同時選擇一個視窗與一個功能")
            return

        target_win = self.windows_data[win_idx[0]]
        hwnd = target_win[0]
        feature_name = self.lb_features.get(feat_idx[0])

        # 重設停止訊號
        self.stop_event.clear()

        # 使用 Thread 執行功能，避免 UI 凍結
        task_thread = threading.Thread(
            target=self._execute_feature_logic, 
            args=(hwnd, target_win[1], feature_name), 
            daemon=True # 設定為守護執行緒，主程式關閉時會一併停止
        )
        task_thread.start()

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

    def _toggle_debug_mode(self):
        """切換自動化工具的偵錯模式"""
        self.auto.debug_mode = self.debug_mode_var.get()
        print(f"偵錯模式已設定為: {self.auto.debug_mode}")

def is_admin():
    """檢查是否具有管理員權限"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if __name__ == "__main__":
    if is_admin():
        root = tk.Tk()
        app = NobunagaToolApp(root)
        root.mainloop()
    else:
        # 重新啟動程式並要求管理員權限
        print("權限不足，嘗試取得管理員權限...")
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
