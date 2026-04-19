import time
import cv2
import numpy as np
import win32gui
import win32api
import win32con
from PIL import ImageGrab
from nobunaga_utils import NobunagaVKKey, get_resource_path

class NobunagaAutomation:
    """
    提供信長Online自動化核心功能：影像搜尋與按鍵模擬
    """

    def __init__(self):
        self.debug_mode = False

    def find_image(self, hwnd, template_path, start_x=0, start_y=0, search_width=None, search_height=None, grayscale=False):
        """
        在指定視窗內搜尋影像
        :param hwnd: 視窗句柄 (handle)
        :param template_path: 樣板圖檔路徑
        :param start_x, start_y: 視窗內的相對起始搜尋座標
        :param search_width, search_height: 搜尋範圍的寬高 (若不指定則搜尋到視窗邊界)
        :param grayscale: 是否使用灰階比對
        :param threshold: 門檻值，預設為 0.8
        :return: 若找到則傳回 (relative_x, relative_y)，否則傳回 None
        """
        try:
            # 確保路徑在打包環境下也能正確指向
            template_path = get_resource_path(template_path)

            # 取得視窗座標
            rect = win32gui.GetWindowRect(hwnd)
            # 擷取視窗畫面 (left, top, right, bottom)
            img = ImageGrab.grab(bbox=rect)
            screen_np = np.array(img)
            screen_bgr = cv2.cvtColor(screen_np, cv2.COLOR_RGB2BGR)

            # 讀取樣板 (使用 numpy 讀取以支援中文路徑)
            template = cv2.imdecode(np.fromfile(template_path, dtype=np.uint8), cv2.IMREAD_COLOR)
            if template is None:
                raise FileNotFoundError(f"無法載入圖檔: {template_path}")

            # 定義搜尋區域 (ROI)
            # 確保不會超出原始截圖範圍
            h, w = screen_bgr.shape[:2]
            end_x = min(start_x + search_width, w) if search_width else w
            end_y = min(start_y + search_height, h) if search_height else h
            roi = screen_bgr[min(start_y, h):end_y, min(start_x, w):end_x]

            if grayscale:
                roi = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
                template = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)

            # 檢查 ROI 是否大於樣板，若 ROI 太小會導致 matchTemplate 崩潰
            rh, rw = roi.shape[:2]
            th, tw = template.shape[:2]
            if rh < th or rw < tw:
                # 搜尋區域比圖檔樣板還小，不可能找到
                print(f"rh: {rh}, rw: {rw}, th: {th}, tw: {tw}")
                return None

            # 進行樣板比對
            res = cv2.matchTemplate(roi, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(res)

            match_found = False
            result_coords = None

            # Prepare debug canvas if debug_mode is True
            debug_canvas = None
            if self.debug_mode:
                debug_canvas = screen_bgr.copy()
                # 繪製搜尋範圍 (ROI) - 黃色
                cv2.rectangle(debug_canvas, (start_x, start_y), (end_x, end_y), (0, 255, 255), 2)
                print(f"DEBUG: 搜尋 {template_path} 於 ROI: ({start_x},{start_y}) 到 ({end_x},{end_y})")

            # 門檻值設定，通常 0.8 以上為高相關性
            threshold = 0.8
            if max_val >= threshold:
                match_found = True
                if self.debug_mode:
                    print(f"DEBUG: 找到圖像 {template_path}, 相似度: {max_val:.4f}")

                # 如果有指定搜尋範圍，則檢查圖像中心是否在範圍中心點 +/- 20 像素
                if search_width is not None and search_height is not None:
                    match_center_x = max_loc[0] + (tw // 2) # 找到的圖像中心點 (相對於 ROI)
                    match_center_y = max_loc[1] + (th // 2)
                    
                    # ROI 的中心點 (相對於 ROI 自身)
                    roi_center_x = (end_x - start_x) // 2 
                    roi_center_y = (end_y - start_y) // 2

                    if self.debug_mode:
                        print(f"DEBUG: 匹配中心({match_center_x},{match_center_y}), ROI中心({roi_center_x},{roi_center_y})")

                    if abs(match_center_x - roi_center_x) > 10 or abs(match_center_y - roi_center_y) > 10:
                        if self.debug_mode:
                            print("DEBUG: 偏移量過大，判定為不符合中心點條件。")
                        match_found = False # 如果中心點不符合，則視為未找到

                if match_found:
                    # 計算回傳座標（需加上原始 ROI 的偏移量）
                    res_x = max_loc[0] + start_x
                    res_y = max_loc[1] + start_y
                    result_coords = (res_x, res_y)

                    if self.debug_mode:
                        # 繪製找到的圖像位置 - 紅色
                        cv2.rectangle(debug_canvas, (res_x, res_y), (res_x + tw, res_y + th), (0, 0, 255), 2)
            else:
                if self.debug_mode:
                    print(f"DEBUG: 未找到圖像 {template_path}, 相似度: {max_val:.4f} (低於門檻 {threshold})")

            # Display debug canvas if debug_mode is True, regardless of match_found
            '''
            if self.debug_mode and debug_canvas is not None:
                cv2.imshow("Automation Debug - find_image", debug_canvas)
                cv2.waitKey(500) # Display for 0.5 seconds
                cv2.destroyAllWindows() # Close the window after display
            '''

            return result_coords
        except Exception as e:
            print(f"影像搜尋發生錯誤: {e}")
            return None
    
    def find_image_click(self, hwnd, template_path, start_x=0, start_y=0):
        """
        在視窗內搜尋指定圖像，並移動滑鼠左鍵點擊該圖像中心。
        :return: 若找到並成功發送點擊訊息傳回 True，否則返回 False
        """
        match_pos = self.find_image(hwnd, template_path, start_x, start_y)
        # template_path 處理已包含在 find_image 內，但下方讀取樣板時仍需處理
        template_full_path = get_resource_path(template_path)
        if match_pos:
            try:
                # 讀取樣板以獲取寬高，用來計算中心點
                template = cv2.imdecode(np.fromfile(template_full_path, dtype=np.uint8), cv2.IMREAD_COLOR)
                if template is None:
                    return False
                
                h, w = template.shape[:2]
                # 計算中心點相對座標（相對於 find_image 回傳的視窗座標）
                center_x = match_pos[0] + w // 2
                center_y = match_pos[1] + h // 2

                # 將相對視窗外框的座標轉換為螢幕絕對座標
                rect = win32gui.GetWindowRect(hwnd)
                screen_pos = (rect[0] + center_x, rect[1] + center_y)

                # 轉換為視窗客戶區 (Client Area) 座標，這是 PostMessage 要求的座標系
                client_x, client_y = win32gui.ScreenToClient(hwnd, screen_pos)

                if self.debug_mode:
                    c_left, c_top, c_right, c_bottom = win32gui.GetClientRect(hwnd)
                    print(f"DEBUG [find_image_click]:")
                    print(f"  - 樣板路徑: {template_path} ({w}x{h})")
                    print(f"  - 匹配左上角 (視窗相對): {match_pos}")
                    print(f"  - 點擊中心點 (視窗相對): ({center_x}, {center_y})")
                    print(f"  - 最終 Client 座標: ({client_x}, {client_y})")
                    print(f"  - 客戶區寬高: {c_right}x{c_bottom}")

                    # 視覺化除錯：擷取目前視窗並標示位置
                    img = ImageGrab.grab(bbox=rect)
                    debug_img = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
                    # 繪製匹配到的紅色外框
                    cv2.rectangle(debug_img, (match_pos[0], match_pos[1]), 
                                  (match_pos[0] + w, match_pos[1] + h), (0, 0, 255), 2)
                    # 繪製預計點擊的綠色實心圓點
                    cv2.circle(debug_img, (center_x, center_y), 7, (0, 255, 0), -1)
                    cv2.putText(debug_img, f"CLICK HERE ({client_x}, {client_y})", 
                                (center_x + 10, center_y), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 255, 0), 2)

                    cv2.imshow("Debug: find_image_click", debug_img)
                    cv2.waitKey(1000) # 暫停 1 秒讓開發者確認
                    cv2.destroyAllWindows()

                lparam = win32api.MAKELONG(client_x, client_y)

                # 激活視窗並模擬移動與點擊
                win32gui.SendMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_CLICKACTIVE, 0)
                win32api.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lparam)
                time.sleep(0.05)
                # 發送滑鼠按下與放開訊息 (後台模擬點擊)
                win32api.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
                time.sleep(0.1) # 增加點擊持續時間，確保遊戲引擎偵測到
                win32api.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
                
                return True
            except Exception as e:
                print(f"find_image_click 發生錯誤: {e}")
                return False
        return False

    def send_key(self, hwnd, vk_code, modifiers=None, hold_time=0.1):
        """
        對指定視窗發送鍵盤按鍵（後台模擬）
        :param hwnd: 視窗句柄
        :param vk_code: 虛擬鍵碼 (例如 0x20 為 Space)
        :param modifiers: 輔助按鍵列表 (例如 [win32con.VK_SHIFT, win32con.VK_CONTROL])
        :param hold_time: 按下與放開之間的延遲時間 (秒)
        """
        try:
            if modifiers is None:
                modifiers = []

            # 1. 按下所有輔助按鍵
            for mod in modifiers:
                win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, mod, 0)

            # 發送按下訊號 (KeyDown)
            win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, 0)
            
            # 持續時間
            if hold_time > 0:
                time.sleep(hold_time)
            
            # 發送放開訊號 (KeyUp)
            win32api.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, 0xC0000001)

            # 2. 逆序放開所有輔助按鍵
            for mod in reversed(modifiers):
                win32api.PostMessage(hwnd, win32con.WM_KEYUP, mod, 0xC0000001)

        except Exception as e:
            print(f"發送按鍵發生錯誤: {e}")

class NobunagaStateCheck:
    """
    提供遊戲狀態檢查功能，透過影像比對判斷當前狀態。
    """

    def is_combat_in(self, hwnd, automation):
        """檢查視窗畫面是否有 img/戰鬥中.png. 相符則成立"""
        # 呼叫傳入的 automation 實例進行影像搜尋
        result = automation.find_image(hwnd, 'img/戰鬥中.png')
        return result is not None

    def is_combat_end(self, hwnd, automation):
        """檢查視窗畫面是否有 img/戰鬥結束_剩下.png. 相符則成立"""
        result = automation.find_image(hwnd, 'img/戰鬥結束_剩下.png')
        return result is not None
    
    def is_dead(self, hwnd, automation):
        """檢查視窗畫面是否有 img/成佛圖示.png. 相符則成立"""
        result = automation.find_image(hwnd, 'img/成佛圖示.png')
        return result is not None

    def is_next_floor_dialog(self, hwnd, automation):
        """檢查視窗畫面是否有 img/是否移動下一層.png. 相符則成立"""
        result = automation.find_image(hwnd, 'img/是否移動下一層.png')
        return result is not None
    
class NobunagaAction:
    def menu_team_hero_select(self, hwnd, automation, hero_team_index=0):
        """
        英傑隊伍選擇自動化步驟：
        1. Z鍵開啟menu
        2. 搜尋 跟隨NPC.png 並且點擊
        3. 搜尋 英傑.png 並且點擊
        4. 按鍵 I -> Enter -> Enter
        5. 按鍵 K 次數為 hero_team_index - 1
        """
        VK_Z = NobunagaVKKey.VK_Z.value
        VK_I = NobunagaVKKey.VK_I.value
        VK_K = NobunagaVKKey.VK_K.value
        VK_ENTER = NobunagaVKKey.VK_ENTER.value
        VK_ESCAPE = NobunagaVKKey.VK_ESCAPE.value


        # 1. Z鍵開啟menu 搜尋 跟隨NPC.png並且滑鼠左點點擊
        while not automation.find_image_click(hwnd, 'img/跟隨NPC.png'):
            automation.send_key(hwnd, VK_ESCAPE, hold_time=0.2)
            time.sleep(0.5)
            automation.send_key(hwnd, VK_Z, hold_time=0.2)
            
        time.sleep(0.5)

        # 3. 搜尋 英傑.png並且滑鼠左點點擊
        if not automation.find_image_click(hwnd, 'img/英傑.png'):
            return False
        time.sleep(0.5)

        # 4. 按鍵I > Enter > Enter > Enter
        keys = [VK_I, VK_ENTER, VK_ENTER, VK_ENTER]
        automation.send_key(hwnd, VK_I, hold_time=0.2)
        time.sleep(0.3)

        while not automation.find_image(hwnd, 'img/英傑選單_一併登錄.png'):
            automation.send_key(hwnd, VK_ENTER, hold_time=0.2)
            time.sleep(0.3)
        
        time.sleep(0.5)

        # 5. 按鍵K. 次數為hero_team_index - 1
        # 如果 hero_team_index 為 1 或 0，則不執行循環
        for _ in range(max(0, hero_team_index - 1)):
            automation.send_key(hwnd, VK_K, hold_time=0.2)
            time.sleep(0.5)
        
        # 6. Enter > Enter
        keys = [VK_ENTER, VK_ENTER, VK_ENTER]
        for key in keys:
            automation.send_key(hwnd, key, hold_time=0.2)
            time.sleep(0.5)

        return True

    def move_head_north(self, hwnd, automation):
        """
        搜尋朝北圖像並且轉向：
        1. 檢查 img/第一人稱.png 是否為第一視角。否則先按下 V 鍵。
        2. 檢查是否有正北圖像 img/北.png。否則持續 A 鍵旋轉，直到找到。
        :return: 成功轉向朝北返回 True，否則返回 False
        """
        VK_V = NobunagaVKKey.VK_V.value
        VK_A = NobunagaVKKey.VK_A.value

        isNorthFound = False


        # 1. 檢查視角狀態
        # 如果沒偵測到第一人稱圖示，代表可能是第三人稱，按下 V 切換
        if not automation.find_image(hwnd, 'img/第一人稱.png'):
            print("[動作] 非第一人稱視角，按下 V 鍵切換...")
            automation.send_key(hwnd, VK_V, hold_time=0.2)
            time.sleep(1.0) # 等待視角切換動畫

        # 2. 旋轉尋找正北
        # 使用迴圈持續短按 A 鍵旋轉，直到在畫面上（通常是上方羅盤）偵測到 '北' 字
        max_attempts = 60 # 避免無限旋轉的保險門檻
        for i in range(max_attempts):
            #if automation.find_image(hwnd, 'img/北0.png', start_x=986, start_y=911, search_width=50, search_height=30):
            if automation.find_image(hwnd, 'img/北0.png', start_x=975, start_y=690, search_width=60, search_height=35):
                print(f"[動作] 已偵測到正北圖像 (嘗試次數: {i+1})")
                isNorthFound = True
                break
            
            # 尚未發現正北，短按 A 鍵向左微調
            automation.send_key(hwnd, VK_A, hold_time=0.05)
            time.sleep(0.1) # 留給遊戲畫面渲染與截圖的時間

        #print("[動作] 轉向失敗：已達到最大嘗試次數")
        if automation.find_image(hwnd, 'img/第一人稱.png'):
            print("[動作] 第一人稱視角，按下 V 鍵切換...")
            automation.send_key(hwnd, VK_V, hold_time=0.2)
            time.sleep(1.0) # 等待視角切換動畫

        return isNorthFound
