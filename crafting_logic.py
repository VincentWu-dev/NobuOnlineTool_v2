import time
import win32con

def start_crafting_loop(hwnd, automation, stop_event, template_path=None):
    """
    稼業連點邏輯：
    每 0.1 秒送出 Enter，並檢查特定圖像是否存在。
    """
    print("[稼業連點] 迴圈開始")
    
    # Enter 的虛擬鍵碼是 0x0D
    VK_ENTER = win32con.VK_RETURN

    while not stop_event.is_set():
        # 1. 送出 Enter
        automation.send_key(hwnd, VK_ENTER, hold_time=0.05)
        
        # 2. 檢查圖像 (如果有提供樣板路徑)
        if template_path:
            # 假設檢查畫面的特定區域，這裡先用全視窗搜尋範例
            result = automation.find_image(hwnd, template_path)
            if result:
                print(f"[稼業連點] 偵測到停止圖像 {result}，自動停止。")
                break
        
        # 3. 短暫停頓 0.1 秒
        # 注意：send_key 內部已有 hold_time，這裡補足間隔
        time.sleep(0.1)

    print("[稼業連點] 迴圈已結束")
