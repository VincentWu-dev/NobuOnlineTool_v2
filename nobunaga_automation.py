import time
import cv2
import numpy as np
import win32gui
import win32api
import win32con
from PIL import ImageGrab

class NobunagaAutomation:
    """
    提供信長Online自動化核心功能：影像搜尋與按鍵模擬
    """

    def find_image(self, hwnd, template_path, start_x=0, start_y=0, grayscale=False):
        """
        在指定視窗內搜尋影像
        :param hwnd: 視窗句柄 (handle)
        :param template_path: 樣板圖檔路徑
        :param start_x, start_y: 視窗內的相對起始搜尋座標
        :param grayscale: 是否使用灰階比對
        :return: 若找到則傳回 (relative_x, relative_y)，否則傳回 None
        """
        try:
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
            roi = screen_bgr[min(start_y, h):, min(start_x, w):]

            if grayscale:
                roi = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
                template = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)

            # 進行樣板比對
            res = cv2.matchTemplate(roi, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(res)

            # 門檻值設定，通常 0.8 以上為高相關性
            threshold = 0.8
            if max_val >= threshold:
                # 計算回傳座標（需加上原始 ROI 的偏移量）
                res_x = max_loc[0] + start_x
                res_y = max_loc[1] + start_y
                return (res_x, res_y)
            
            return None
        except Exception as e:
            print(f"影像搜尋發生錯誤: {e}")
            return None

    def send_key(self, hwnd, vk_code, hold_time=0.1):
        """
        對指定視窗發送鍵盤按鍵（後台模擬）
        :param hwnd: 視窗句柄
        :param vk_code: 虛擬鍵碼 (例如 0x20 為 Space)
        :param hold_time: 按下與放開之間的延遲時間 (秒)
        """
        try:
            # 發送按下訊號 (KeyDown)
            win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, 0)
            
            # 持續時間
            if hold_time > 0:
                time.sleep(hold_time)
            
            # 發送放開訊號 (KeyUp)
            # 依據 Windows API 規範，KeyUp 的 lparam 需設置第 31 位元為 1
            win32api.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, 0xC0000001)
        except Exception as e:
            print(f"發送按鍵發生錯誤: {e}")
