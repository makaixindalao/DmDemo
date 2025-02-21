# dm_controller.py
import ctypes
import os
import logging
from ctypes import windll
from typing import Optional, Tuple
import win32com.client

class DmController:
    """大漠插件控制类（lib目录版本）"""
    
    def __init__(self, work_dir: str = None):
        self.logger = logging.getLogger(__name__)
        # 确保工作目录包含lib子目录
        self.work_dir = os.path.abspath(work_dir) if work_dir else os.getcwd()
        self.lib_dir = os.path.join(self.work_dir, "lib")  # 新增lib目录路径
        self.dm = None
        self._dll_loaded = False

        # 自动创建lib目录（如果需要）
        if not os.path.exists(self.lib_dir):
            self.logger.warning(f"lib目录不存在: {self.lib_dir}")

    def _validate_dll(self) -> bool:
        """验证DLL文件是否存在（lib目录版本）"""
        required_files = {
            "DmReg.dll": os.path.join(self.lib_dir, "DmReg.dll"),
            "dm.dll": os.path.join(self.lib_dir, "dm.dll")
        }
        
        missing = [name for name, path in required_files.items() if not os.path.exists(path)]
        if missing:
            self.logger.error(f"lib目录中缺失必要文件: {', '.join(missing)}")
            self.logger.info(f"预期路径: {self.lib_dir}")
            return False
        return True

    def load_dm_plugin(self) -> bool:
        """加载大漠插件（lib目录版本）"""
        if not self._validate_dll():
            return False

        try:
            reg_dll_path = os.path.join(self.lib_dir, "DmReg.dll")
            dm_dll_path = os.path.join(self.lib_dir, "dm.dll")
            
            # 显式指定宽字符路径（Windows API要求）
            reg_dll = windll.LoadLibrary(reg_dll_path)
            # 使用绝对路径并转换为Windows API需要的宽字符格式
            result = reg_dll.SetDllPathW(
                os.path.abspath(dm_dll_path).encode('utf-16le'), 
                0
            )
            self._dll_loaded = result == 1
            return self._dll_loaded
        except Exception as e:
            self.logger.error(f"DLL加载失败: {str(e)}")
            self.logger.debug("详细路径信息：", exc_info=True)
            return False

    def initialize(self, reg_code: str, ver_info: str) -> bool:
        """初始化插件实例"""
        if not self._dll_loaded and not self.load_dm_plugin():
            return False

        try:
            self.dm = win32com.client.Dispatch("dm.dmsoft")
            reg_result = self.dm.Reg(reg_code, ver_info)
            if reg_result != 1:
                self.logger.error(f"插件注册失败，错误码: {reg_result}")
            return reg_result == 1
        except Exception as e:
            self.logger.error(f"插件初始化异常: {str(e)}")
            return False

    @property
    def version(self) -> Optional[str]:
        """获取插件版本"""
        try:
            return self.dm.Ver() if self.dm else None
        except Exception as e:
            self.logger.error(f"获取版本失败: {str(e)}")
            return None
