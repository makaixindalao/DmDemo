# main.py
import logging
import os
import configparser
from model.dm_controller import DmController
from model.mouse_tracker import MouseTracker
from model.image_manager import ImageManager


def setup_logging():
    """配置日志系统"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('app.log'),
            logging.StreamHandler()
        ]
    )

def load_config() -> configparser.ConfigParser:
    """加载配置文件"""
    config = configparser.ConfigParser()
    config_path = os.path.join(os.getcwd(), 'config.ini')
    
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"配置文件不存在: {config_path}")
    
    config.read(config_path, encoding='utf-8')
    return config

def main():
    setup_logging()
    logger = logging.getLogger(__name__)

    try:
        config = load_config()
        reg_code = config.get('DEFAULT', 'reg_code')
        ver_info = config.get('DEFAULT', 'ver_info')
    except Exception as e:
        logger.error("配置加载失败: %s", str(e))
        return

    # 初始化大漠插件
    dm_controller = DmController()
    if not dm_controller.load_dm_plugin():
        logger.error("大漠插件加载失败")
        return
    
    if not dm_controller.initialize(reg_code, ver_info):
        logger.error("大漠插件初始化失败")
        return

    logger.info("大漠插件版本: %s", dm_controller.version)

    try:
        # 鼠标
        mouse_tracker = MouseTracker(dm_controller.dm)
        # mouse_tracker.set_sim_mode(0)
        # mouse_tracker.set_mouse_speed(6)
        # print(mouse_tracker.get_cursor_position())
        # mouse_tracker.move_to(1000, 1000)
        # print(mouse_tracker.get_cursor_position())
        # mouse_tracker.start_tracking()
        # 图色
        # image_manager = ImageManager(dm_controller.dm)
        # image_manager.set_path("./screenshots")  # 设置截图保存目录
        print(dm_controller.dm.SetPath(str(os.path.join(os.getcwd(), "screenshots"))))
        print(dm_controller.dm.CaptureJpg(0, 0, 1024, 768, "desktop.jpg", 85))
        print(dm_controller.dm.MoveTo(2920, 600))
        # result = image_manager.capture_jpg(
        #     x1=0, y1=0, x2=1024, y2=768,
        #     filename="desktop.jpg",
        #     quality=85
        # )
        # if result == 1:
        #     logger.info("截图操作成功")
    except KeyboardInterrupt:
        logger.info("程序正常退出")

if __name__ == "__main__":
    main()
