__copyright__ = "Copyright 2022-2023 USTB AI3D LAB"
__description__ = "Wisdom Store Desktop Client"

import logging
import sys
import os
import time
from PyQt5 import QtCore, QtGui
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication
import pywintypes  # 在win32api之前加载，尝试解决win32api涨不到模块的问题
from wisdom_store.config import Config
from wisdom_store.src.views import HomeView
from wisdom_store.src.api_old import API
from wisdom_store.auth import Auth
from wisdom_store.wins.start_loading_page import CustomSplashScreen
from wisdom_store.wins.start_error_report import ErrorReportWin
from wisdom_store.wins.start_update_tip import UpdateTipWin
import multiprocessing
from wisdom_store.settings import _isProduction, _version
import traceback
from logging.handlers import RotatingFileHandler
from wisdom_store.src.utils.addExtendSupport import add_file_association, refresh_icon
from wisdom_store.src.sdk.project.project import Project
import shutil
from pathlib import Path


class WisdomStore:

    def __init__(self, file=None):
        self.config: Config = None
        self.auth: Auth = None
        self.exeFilePath = file
        self.errorReportWin = None
        self.updateTipWin = None
        self.view = None

    def run(self):
        multiprocessing.freeze_support()  # 解决打包后启动进程直接启动整个程序的问题
        # 适应windows缩放
        QtCore.QCoreApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
        # 设置支持小数放大比例（适应如125%的缩放比）
        QtGui.QGuiApplication.setAttribute(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)

        app = QApplication(sys.argv)

        # 启动界面
        splash = CustomSplashScreen()
        splash.show()

        app.processEvents()  # 处理主进程事件

        # 初始化配置和授权
        splash.updateInfo(15)
        self.config = Config()
        self.config.language = 'zh_CN'
        self.config.version = _version
        self.config.exeFilePath = os.path.abspath(self.exeFilePath)
        print('Path of exe: ', self.config.exeFilePath)
        splash.updateInfo(20)
        self.auth = Auth()

        # 上报之前的使用记录
        print(f'Mode: {"Production" if _isProduction else "Development"}')
        if _isProduction:
            from wisdom_store.reporter import recorder
            # 启动记录器
            recorder.on_create()  # 新增一条记录
            # 定时记录使用记录
            recorder.auto_exec(60, app)  # 每隔一段时间自动更新记录（线程）
            recorder.upload()  # 软件启动后，上报之前的记录

        # 创建界面
        splash.updateInfo(20)
        # 主界面
        self.view = HomeView(self.config, self.auth)
        # self.view.win.ui.pBtnNew.clicked.connect(self.error_test)
        # 崩溃提示界面
        self.errorReportWin = ErrorReportWin(self.config)  # 错误报告界面
        # 软件更新界面
        self.updateTipWin = UpdateTipWin(self.config)
        self.config.updateWin = self.updateTipWin
        splash.updateInfo(isFinished=True)
        time.sleep(0.5)
        self.view.show()
        splash.close()

        # 尝试打开项目
        if len(sys.argv) == 2:
            try:
                self.view.openProject(configPath=sys.argv[1])
            except Exception as e:
                self.view.win.alertError('提示', f'无法打开项目：{sys.argv[1]}', str(e))

        # 关联.wsp扩展名与exe程序
        if os.path.splitext(self.exeFilePath)[1] == '.exe':
            try:  # 需要管理员权限
                add_file_association(os.path.splitext(Project.ConfigName)[1], self.exeFilePath)  # 创建文件关联
                refresh_icon()  # 刷新图标
                logging.info('建立扩展名关联成功！')
            except Exception as e:
                logging.error(f'权限不足，无法创建扩展名关联：{str(e)}')
                logging.info("请使用管理员权限重新启动")

        # 复制Yolo字体文件到用户目录
        try:
            yolo_config_dir = self.config.get_yolo_user_config_dir()
            os.makedirs(yolo_config_dir, exist_ok=True)
            os.makedirs(yolo_config_dir, exist_ok=True)
            # 复制字体文件
            base_dir = Path(self.exeFilePath).parent
            for font_name in ['Arial.ttf', 'Arial.Unicode.ttf']:
                src_path = base_dir / font_name
                dist_path = yolo_config_dir / font_name
                if not dist_path.exists():
                    shutil.copy(src_path, dist_path)
        except Exception as e:
            logging.error(f'复制字体到目标路径失败：{str(e)}')

        # 重写捕获异常处理方法
        # sys.excepthook = self.handle_exception
        sys.exit(app.exec_())

    # def error_test(self):
    #     raise Exception('奔溃测试')

    def handle_exception(self, *args):
        '''
        sys.excepthook以底层的方式捕获异常try..exception无法捕获到的Pyqt5的异常（已经try...except处理的不会被捕获）
        :param args:
        :return:
        '''
        error_type = args[0]
        error_info = args[1]
        error_traceback = args[2].tb_frame
        info = [error_type, error_info, error_traceback]
        logging.info(f'软件奔溃信息：{info}')
        title = f'软件奔溃：{error_type}'
        content = f'错误信息：{error_info}；问题定位：{error_traceback}'
        self.errorReportWin.receiveInfo(title=title, content=content)
        self.errorReportWin.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.errorReportWin.show()
        self.view.hide()


if __name__ == '__main__':
    # logging.basicConfig(level=logging.DEBUG, format='%(processName)s %(asctime)s %(levelname)s %(message)s')
    logging.basicConfig(level=logging.DEBUG,
                        format='%(asctime)s - %(processName)s %(pathname)s [%(funcName)s/line:%(lineno)d] - %(levelname)s: %(message)s')
    FileLogHandler = RotatingFileHandler(Config.LOG_PATH, maxBytes=1024 * 1024 * 100, backupCount=10)
    logging.getLogger().addHandler(FileLogHandler)
    # app = WisdomStore()
    # app.run()

    # 确定应用程序是脚本文件还是被冻结的exe
    if getattr(sys, 'frozen', False):
        # 获取应用程序exe的路径（打包后）
        app = WisdomStore(sys.executable)
        app.run()
    elif __file__:
        # 获取脚本程序的路径（打包前）
        app = WisdomStore(__file__)
        app.run()
