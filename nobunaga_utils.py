import os
import sys
import win32api
import win32con
from enum import Enum

def get_resource_path(relative_path):
    """ 取得資源的絕對路徑，支援開發環境與 auto-py-to-exe 打包後的環境 """
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 打包後的路徑
        return os.path.join(sys._MEIPASS, relative_path)
    # 開發環境下的路徑 (相對於目前執行目錄或指定的基準路徑)
    return os.path.join(os.path.abspath("."), relative_path)

class NobunagaVKKey(Enum):
    VK_ENTER = win32con.VK_RETURN
    VK_SHIFT = win32con.VK_SHIFT
    VK_S = 0x53
    VK_A = 0x41
    VK_D = 0x44
    VK_W = 0x57
    VK_ESCAPE = win32con.VK_ESCAPE
    VK_J = 0x4A
    VK_L = 0x4C
    VK_K = 0x4B
    VK_I = 0x49
    VK_V = 0x56
    VK_N = 0x4E  # 假設用於轉向北方或重置視角的按鍵
    VK_Z = 0x5A # 滑鼠右鍵選單
    VK_Y = 0x59 #選取對象



class DungeonState(Enum):
    """通用狀態機"""
    IDLE = 0            # 空閒/等待
    FINDING_TARGET = 1  # 前進至目標
    IN_BATTLE = 2       # 戰鬥中
    BATTLE_END = 3      # 戰鬥結束(點收物品)
    DEATH_CHECK = 4     # 判斷是否死亡
    RECALL_PARTY = 5    # 重新叫出隊伍
    NEXT_FLOOR = 6      # 下一層
    COMBAT_CLEANUP = 7  # 戰鬥結束清理


class NobuNagaImageList:
    """
    集中管理所有圖檔路徑，方便維護與調用。
    """
    # 戰鬥相關
    IN_BATTLE = get_resource_path('img/戰鬥中.png')
    BATTLE_RESULT = get_resource_path('img/戰鬥結束_剩下.png')
    EXPLORATION_UI = get_resource_path('img/exploration_ui.png')
    
    # 狀態相關
    DEATH_DIALOG = get_resource_path('img/death_dialog.png')
    CRAFTING_STOP = get_resource_path('img/材料不夠.png')

    @classmethod
    def get_dreamdungeon_config(cls):
        """
        回傳冥宮掛機邏輯所需的圖檔路徑字典
        """
        return {
            'in_battle': cls.IN_BATTLE,
            'battle_result': cls.BATTLE_RESULT,
            'exploration_ui': cls.EXPLORATION_UI,
            'death_dialog': cls.DEATH_DIALOG
        }
