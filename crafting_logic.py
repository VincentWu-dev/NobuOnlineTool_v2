import time
import win32con
from nobunaga_utils import NobunagaVKKey, DungeonState

def start_crafting_loop(hwnd, automation, stop_event, template_path=None):
    """
    稼業連點邏輯：
    每 0.1 秒送出 Enter，並檢查特定圖像是否存在。
    """
    print("[稼業連點] 迴圈開始")
    
    # Enter 的虛擬鍵碼是 0x0D
    #VK_ENTER = win32con.VK_RETURN
    VK_ENTER = NobunagaVKKey.VK_ENTER.value

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

def dream_dungeon_loop(hwnd, automation, state_check, nobu_action, stop_event,current_state=DungeonState.FINDING_TARGET,
                       hero_team_index = 3, update_floor_callback=None):
    """
    冥宮掛機邏輯 (狀態機實作)
    :param update_floor_callback: 用於更新主畫面樓層數的回調函數
    """
    print("[冥宮掛機] 任務開始")
    floor_count = 0
    #current_state = DungeonState.FINDING_TARGET
    
    # 取得鍵碼
    VK_W = NobunagaVKKey.VK_W.value
    VK_ENTER = NobunagaVKKey.VK_ENTER.value
    VK_V = NobunagaVKKey.VK_V.value
    VK_N = NobunagaVKKey.VK_N.value
    VK_J = NobunagaVKKey.VK_J.value
    VK_Y = NobunagaVKKey.VK_Y.value
    


    while not automation.find_image(hwnd, 'img/對象NPC.png'):
        print(f"對象不是NPC. Y鍵切換")
        automation.send_key(hwnd, VK_Y, hold_time=0.2)
        time.sleep(1.0)
        
    while not stop_event.is_set():
        # --- 狀態 1: 前進至目標 ---
        if current_state == DungeonState.FINDING_TARGET:
            
            # 按下Enter
            automation.send_key(hwnd, VK_ENTER)
            # 持續按下 W 往前移動
            automation.send_key(hwnd, VK_W, hold_time=0.3)
            
            # 判斷是否進入戰鬥 (搜尋戰鬥 UI 圖像)
            if state_check.is_combat_in(hwnd, automation):
                print("[冥宮掛機] 偵測到戰鬥開始，切換狀態：戰鬥中")
                current_state = DungeonState.IN_BATTLE
        
        # --- 狀態 2: 戰鬥中 ---
        elif current_state == DungeonState.IN_BATTLE:
            # 判斷戰鬥是否結束 (搜尋戰鬥結束標誌或結算畫面)
            if state_check.is_combat_end(hwnd, automation):
                print("[冥宮掛機] 戰鬥疑似結束，切換狀態：戰鬥結束點選")
                current_state = DungeonState.BATTLE_END
            elif state_check.is_dead(hwnd, automation):
                print("[冥宮掛機] 偵測到死亡，切換狀態：死亡檢查")
                current_state = DungeonState.DEATH_CHECK
            else:
                # 戰鬥中通常不需要頻繁點擊，每秒檢查一次即可
                time.sleep(1)

        # --- 狀態 3: 戰鬥結束 (處理物品與對話) ---
        elif current_state == DungeonState.BATTLE_END:
            # 嘗試送出 Enter 關閉結算視窗
            automation.send_key(hwnd, VK_ENTER, hold_time=0.1)
            
            # 判斷結算畫面是否消失 (搜尋一般探索時的 UI 圖像)
            if not state_check.is_combat_in(hwnd, automation):
                print("[冥宮掛機] 結算完成，移動進入下一層")
                current_state = DungeonState.NEXT_FLOOR
            time.sleep(0.5)

        # --- 狀態 4: 判斷是否死亡 ---
        elif current_state == DungeonState.DEATH_CHECK:
            # 搜尋成佛圖示
            if state_check.is_dead(hwnd, automation):
                print("[冥宮掛機] 偵測到死亡，關閉死亡對話...")
                # 送出 Enter 關閉死亡對話
                automation.send_key(hwnd, VK_ENTER, hold_time=0.1)
                time.sleep(0.5)
                if automation.find_image_click(hwnd, 'img/YN_確定.png'):
                    print("[冥宮掛機] 按下確定按鍵")
                time.sleep(3)
                #nobu_action.move_head_north(hwnd, automation)
                #current_state = DungeonState.RECALL_PARTY
            else:
                print("[冥宮掛機] 離開死亡對話，轉向正北，進入重新叫出隊友狀態...")
                # 執行轉向正北
                nobu_action.move_head_north(hwnd, automation)
                time.sleep(1)
                current_state = DungeonState.RECALL_PARTY

        # --- 狀態 5: 重新叫出隊伍 ---
        elif current_state == DungeonState.RECALL_PARTY:
            print("[冥宮掛機] 正在重新叫出隊伍/隊友...")
            #叫出隊伍三
            #automation.send_key(hwnd, VK_V, hold_time=0.2)
            nobu_action.menu_team_hero_select(hwnd, automation, hero_team_index)
            
            time.sleep(1) # 等待隊伍載入的時間
            
            # 完成後回到下一層狀態
            current_state = DungeonState.FINDING_TARGET
        
        # --- 狀態 6: 下一層 ---
        elif current_state == DungeonState.NEXT_FLOOR:
            print("[冥宮掛機] 檢查下一層對話. 移動到下一層...")

            if state_check.is_next_floor_dialog(hwnd, automation):
                automation.send_key(hwnd, VK_J, hold_time=0.2)
                automation.send_key(hwnd, VK_ENTER)
                floor_count += 1
                if update_floor_callback:
                    update_floor_callback(floor_count)
                current_state = DungeonState.FINDING_TARGET
                time.sleep(3)
            else:
                automation.send_key(hwnd, VK_ENTER)
                automation.send_key(hwnd, VK_W, hold_time=0.3)


        # 全域微小停頓避免 CPU 過高
        time.sleep(0.4)

    print("[冥宮掛機] 任務已停止")

def follow_combat_loop(hwnd, automation, state_check, nobu_action, stop_event, update_floor_callback=None, item_use=True):
    """
    跟隨戰鬥邏輯
    1. 檢查是否進入戰鬥。
    2. 戰鬥中檢查物品使用。
    3. 戰鬥結束檢查確定對話框。
    4. 戰鬥結束後清理畫面。
    """
    print("[跟隨戰鬥] 任務開始")
    battle_count = 0
    current_state = DungeonState.IDLE  # 使用 DungeonState 管理狀態
    
    VK_ENTER = NobunagaVKKey.VK_ENTER.value
    # 假設物品圖片路徑，可根據實際需求修改
    ITEM_IMAGE_PATH = 'img/選單_物品.png' 
    CONFIRM_IMAGE_PATH = 'img/對話_確定.png'

    while not stop_event.is_set():
        # --- 狀態 1: 檢查是否進入戰鬥 ---
        if current_state == DungeonState.IDLE:
            if state_check.is_combat_in(hwnd, automation):
                print("[跟隨戰鬥] 偵測到戰鬥開始，切換至戰鬥中狀態")
                current_state = DungeonState.IN_BATTLE

        # --- 狀態 2: 戰鬥中 (檢查物品) ---
        elif current_state == DungeonState.IN_BATTLE:
            if item_use:
                # 檢查是否有物品圖片，有則點擊
                if automation.find_image_click(hwnd, ITEM_IMAGE_PATH):
                    print(f"[跟隨戰鬥] 偵測到物品，執行點擊: {ITEM_IMAGE_PATH}")
            
            # 檢查是否戰鬥結束
            if state_check.is_combat_end(hwnd, automation):
                print("[跟隨戰鬥] 偵測到戰鬥結束標誌，進入結算狀態")
                current_state = DungeonState.BATTLE_END

        # --- 狀態 3: 檢查結束對話 ---
        elif current_state == DungeonState.BATTLE_END:
            # 檢查戰鬥畫面是否消失
            time.sleep(0.5)
            if not state_check.is_combat_in(hwnd, automation):
                print("[跟隨戰鬥] 戰鬥已完全結束，回到空閒狀態")
                battle_count += 1
                if update_floor_callback:
                    update_floor_callback(battle_count)
                current_state = DungeonState.IDLE
            else:
                #檢查是否有物品選擇
                if automation.find_image_click(hwnd, CONFIRM_IMAGE_PATH):
                    print("[跟隨戰鬥] 偵測並點擊物品選擇確定按鈕")
                else:
                    # 畫面未消失則送出 Enter
                    automation.send_key(hwnd, VK_ENTER, hold_time=0.1)

        time.sleep(0.5)  # 避免迴圈過快
    print("[跟隨戰鬥] 任務已停止")
