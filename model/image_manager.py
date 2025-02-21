# image_manager.py
import logging
import os
from typing import Optional
import win32com.client

class ImageManager:
    """大漠插件图色处理模块"""
    
    def __init__(self, dm_instance: win32com.client.CDispatch):
        """
        初始化图色模块
        Args:
            dm_instance: 已初始化的大漠插件实例
        """
        self.logger = logging.getLogger(__name__)
        self.dm = dm_instance
        self._default_path = os.getcwd()  # 默认保存路径

    def set_path(self, path: str) -> bool:
        """
        设置截图默认保存路径
        Args:
            path: 路径字符串
        Returns:
            bool: 是否设置成功
        """
        try:
            if not os.path.exists(path):
                os.makedirs(path, exist_ok=True)
                self.logger.info(f"创建目录: {path}")
            
            self._default_path = path
            return self.dm.SetPath(path) == 1
        except Exception as e:
            self.logger.error(f"设置路径失败: {str(e)}")
            return False

    def capture_jpg(self, x1: int, y1: int, x2: int, y2: int, 
                   filename: str, quality: int = 90) -> int:
        """
        区域截图保存为JPG
        Args:
            x1: 区域左上X坐标
            y1: 区域左上Y坐标
            x2: 区域右下X坐标
            y2: 区域右下Y坐标
            filename: 文件名（可包含路径）
            quality: 压缩质量 (1-100)
        Returns:
            int: 0失败，1成功
        """
        # 参数验证
        if not self._validate_coordinates(x1, y1, x2, y2):
            return 0
        
        if not self._validate_quality(quality):
            return 0

        # 路径处理
        full_path = self._process_filepath(filename)
        if not full_path:
            return 0

        try:
            result = self.dm.CaptureJpg(x1, y1, x2, y2, full_path, quality)
            if result == 1:
                self.logger.info(f"截图成功保存至: {full_path}")
            else:
                self.logger.error(f"截图失败，错误码: {result}")
            return result
        except Exception as e:
            self.logger.error(f"截图异常: {str(e)}")
            return 0

    def _validate_coordinates(self, x1: int, y1: int, x2: int, y2: int) -> bool:
        """验证坐标有效性"""
        if x1 >= x2 or y1 >= y2:
            self.logger.error(f"无效区域坐标: ({x1},{y1})-({x2},{y2})")
            return False
        return True

    def _validate_quality(self, quality: int) -> bool:
        """验证压缩质量"""
        if not 1 <= quality <= 100:
            self.logger.error(f"无效压缩质量: {quality}，应在1-100之间")
            return False
        return True

    def _process_filepath(self, filename: str) -> Optional[str]:
        """处理文件路径"""
        # 绝对路径直接使用
        if os.path.isabs(filename):
            dir_path = os.path.dirname(filename)
            if not os.path.exists(dir_path):
                try:
                    os.makedirs(dir_path, exist_ok=True)
                except Exception as e:
                    self.logger.error(f"创建目录失败: {dir_path} ({str(e)})")
                    return None
            return filename
        
        # 相对路径结合默认路径
        full_path = os.path.join(self._default_path, filename)
        dir_path = os.path.dirname(full_path)
        if not os.path.exists(dir_path):
            try:
                os.makedirs(dir_path, exist_ok=True)
            except Exception as e:
                self.logger.error(f"创建目录失败: {dir_path} ({str(e)})")
                return None
        return full_path

    # 可扩展添加其他图色功能，例如：
    # find_image, get_color, etc.
