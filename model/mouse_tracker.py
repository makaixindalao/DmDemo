# mouse_tracker.py
import logging
from typing import Optional, Tuple
import win32com.client

class MouseTracker:
    """鼠标操作控制器"""
    
    def __init__(self, dm_instance: win32com.client.CDispatch):
        self.logger = logging.getLogger(__name__)
        self.dm = dm_instance

    def get_cursor_position(self) -> Optional[Tuple[int, int]]:
        """获取当前光标坐标"""
        try:
            result = self.dm.GetCursorPos()
            if result[0] == 1:
                return (result[1], result[2])
            self.logger.warning("获取坐标失败，错误码: %d", result[0])
            return None
        except AttributeError:
            self.logger.error("大漠插件实例未初始化")
            return None
        except Exception as e:
            self.logger.error("坐标获取异常: %s", str(e))
            return None

    def start_tracking(self, interval: float = 0.5):
        """启动持续跟踪"""
        import time
        try:
            while True:
                pos = self.get_cursor_position()
                if pos:
                    self.logger.info("当前光标位置: x=%d, y=%d", *pos)
                time.sleep(interval)
        except KeyboardInterrupt:
            self.logger.info("鼠标跟踪已停止")
