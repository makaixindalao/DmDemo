# mouse_tracker.py
import logging
from typing import Optional, Tuple, Union
import win32com.client

class MouseTracker:
    """大漠插件鼠标键盘操作封装"""
    
    def __init__(self, dm_instance: win32com.client.CDispatch):
        self.logger = logging.getLogger(__name__)
        self.dm = dm_instance

    # region 基础验证
    def _check_instance(self) -> bool:
        """验证插件实例是否有效"""
        if not self.dm:
            self.logger.error("大漠插件实例未初始化")
            return False
        return True
    # endregion

    # region 鼠标操作
    def move_to(self, x: int, y: int) -> int:
        """移动鼠标到指定坐标
        Args:
            x: 目标X坐标
            y: 目标Y坐标
        Returns:
            0失败，1成功
        """
        if not self._check_instance():
            return 0
        try:
            return self.dm.MoveTo(x, y)
        except Exception as e:
            self.logger.error(f"移动鼠标到({x},{y})失败: {str(e)}")
            return 0

    def get_cursor_position(self) -> Tuple[int, int]:
        """获取鼠标当前位置"""
        if not self._check_instance():
            return 0, 0
        try:
            ret, x, y = self.dm.GetCursorPos()
            if ret == 1:
                return x, y
            return 0, 0
        except Exception as e:
            self.logger.error(f"获取鼠标位置失败: {str(e)}")
            return 0, 0

    def left_click(self) -> int:
        """左键单击"""
        return self._click_wrapper("LeftClick")

    def left_down(self) -> int:
        """左键按下"""
        return self._click_wrapper("LeftDown")

    def left_up(self) -> int:
        """左键弹起"""
        return self._click_wrapper("LeftUp")

    def right_click(self) -> int:
        """右键单击"""
        return self._click_wrapper("RightClick")

    def right_down(self) -> int:
        """右键按下"""
        return self._click_wrapper("RightDown")

    def right_up(self) -> int:
        """右键弹起"""
        return self._click_wrapper("RightUp")

    def middle_click(self) -> int:
        """中键单击"""
        return self._click_wrapper("MiddleClick")

    def wheel_down(self) -> int:
        """滚轮下滚"""
        return self._click_wrapper("WheelDown")

    def wheel_up(self) -> int:
        """滚轮上滚"""
        return self._click_wrapper("WheelUp")

    def _click_wrapper(self, method_name: str) -> int:
        """点击操作通用封装"""
        if not self._check_instance():
            return 0
        try:
            return getattr(self.dm, method_name)()
        except AttributeError:
            self.logger.error(f"不支持的方法: {method_name}")
            return 0
        except Exception as e:
            self.logger.error(f"执行{method_name}失败: {str(e)}")
            return 0
    # endregion

    # region 键盘操作
    def key_down(self, vk_code: Union[int, str]) -> int:
        """按下键盘按键
        Args:
            vk_code: 虚拟键码(int)或字符(str)
        Returns:
            0失败，1成功
        """
        return self._keyboard_wrapper("KeyDown", vk_code)

    def key_up(self, vk_code: Union[int, str]) -> int:
        """释放键盘按键"""
        return self._keyboard_wrapper("KeyUp", vk_code)

    def key_press(self, vk_code: Union[int, str]) -> int:
        """按下并释放键盘按键"""
        return self._keyboard_wrapper("KeyPress", vk_code)

    def key_press_str(self, key_str: str, delay: int = 100) -> int:
        """输入字符串
        Args:
            key_str: 要输入的字符串
            delay: 每个字符的间隔时间(毫秒)
        Returns:
            实际输入的字符数
        """
        if not self._check_instance():
            return 0
        try:
            return self.dm.KeyPressStr(key_str, delay)
        except Exception as e:
            self.logger.error(f"输入字符串'{key_str}'失败: {str(e)}")
            return 0

    def _keyboard_wrapper(self, method_name: str, key: Union[int, str]) -> int:
        """键盘操作通用封装"""
        if not self._check_instance():
            return 0
        
        try:
            # 自动转换字符为ASCII码
            if isinstance(key, str):
                if len(key) != 1:
                    raise ValueError("字符参数只能为单个字符")
                vk_code = ord(key.upper())
            else:
                vk_code = int(key)
            
            return getattr(self.dm, method_name)(vk_code)
        except ValueError as e:
            self.logger.error(f"无效的键码参数: {key} ({str(e)})")
            return 0
        except Exception as e:
            self.logger.error(f"执行{method_name}({key})失败: {str(e)}")
            return 0
    # endregion

    # region 延迟设置
    def set_keypad_delay(self, delay_type: str, delay: int) -> int:
        """设置键盘延迟
        Args:
            delay_type: 延迟类型 
                'normal'=普通模式, 'windows'=windows模拟方式
            delay: 延迟时间(毫秒)
        Returns:
            0失败，1成功
        """
        type_map = {'normal': 0, 'windows': 1}
        return self._set_delay_wrapper("SetKeypadDelay", delay_type, type_map, delay)

    def set_mouse_delay(self, delay_type: str, delay: int) -> int:
        """设置鼠标延迟
        Args:
            delay_type: 延迟类型
                'normal'=普通模式, 'windows'=windows模拟方式
            delay: 延迟时间(毫秒)
        """
        type_map = {'normal': 0, 'windows': 1}
        return self._set_delay_wrapper("SetMouseDelay", delay_type, type_map, delay)

    def set_mouse_speed(self, speed: int) -> int:
        """设置鼠标延迟
        函数原型:
        long SetMouseSpeed(speed)

        参数定义:
        speed 整形数:鼠标移动速度, 最小1，最大11.  居中为6. 推荐设置为6

        返回值:
        整形数:
        0:失败
        1:成功
        """
        return self.dm.SetMouseSpeed(speed)

    def _set_delay_wrapper(self, method_name: str, delay_type: str, 
                         type_map: dict, delay: int) -> int:
        """延迟设置通用封装"""
        if not self._check_instance():
            return 0
        
        try:
            type_code = type_map.get(delay_type.lower(), -1)
            if type_code == -1:
                raise ValueError(f"无效的延迟类型: {delay_type}")
            
            return getattr(self.dm, method_name)(type_code, delay)
        except ValueError as e:
            self.logger.error(str(e))
            return 0
        except Exception as e:
            self.logger.error(f"设置延迟失败: {str(e)}")
            return 0
    # endregion

    # region 高级功能
    def set_sim_mode(self, mode: int) -> int:
        """设置模拟模式
        Args:
            mode: 模拟方式 
                0 正常模式(默认模式)
                1 硬件模拟
                2 硬件模拟2(ps2)
                3 硬件模拟3
        Returns:
            0失败，1成功
        """
        if not self._check_instance():
            return 0
        
        return self.dm.SetSimMode(mode)

    def wait_key(self, vk_code: Union[int, str], timeout: int = 5000) -> int:
        """等待按键按下
        Args:
            vk_code: 等待的虚拟键码
            timeout: 超时时间(毫秒)
        Returns:
            0=超时，1=成功
        """
        if not self._check_instance():
            return 0
        
        try:
            if isinstance(vk_code, str):
                if len(vk_code) != 1:
                    raise ValueError("字符参数只能为单个字符")
                vk_code = ord(vk_code.upper())
            return self.dm.WaitKey(int(vk_code), int(timeout))
        except ValueError as e:
            self.logger.error(f"无效的键码参数: {vk_code} ({str(e)})")
            return 0
        except Exception as e:
            self.logger.error(f"等待按键失败: {str(e)}")
            return 0
    # endregion
