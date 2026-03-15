"""
Apple Music Migrator - 音乐搜索自动化工具
重构版本：优化代码结构，增强鲁棒性，移除冗余
"""
import os
os.environ.setdefault("QT_LOGGING_RULES", "qt.qpa.window=false")
import sys
import pandas as pd
import pyautogui
import pyperclip
import time
import ctypes
from ctypes import wintypes
from threading import Thread, Event
from multiprocessing import Process, Event as MEvent, Value
import multiprocessing
from PIL import Image, ImageEnhance, ImageQt
from PySide6 import QtCore, QtGui, QtWidgets
import win32gui
import win32con
import win32api
import win32ui
import numpy as np

# 设置 PyAutoGUI 响应速度为最高（瞬时执行，减少步骤间延迟）
pyautogui.PAUSE = 0
pyautogui.FAILSAFE = True

class TaskTerminated(Exception):
    """自定义异常：用于瞬时中断自动化任务"""
    pass

class AppConfig:
    """应用程序配置 - 集中管理所有常量"""
    VERSION = "v1.0"
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    ASSETS_DIR = os.path.join(BASE_DIR, "assets")
    DATA_DIR = os.path.join(BASE_DIR, "data")
    FONTS_DIR = os.path.join(BASE_DIR, "fonts")
    EXCEL_PATH = os.path.join(DATA_DIR, "歌单.xlsx")
    
    # 图像资源路径
    IMG_ADD_TO_PLAYLIST = os.path.join(ASSETS_DIR, "add_to_playlist.png")
    IMG_TARGET_PLAYLIST = os.path.join(ASSETS_DIR, "target_playlist.png")
    IMG_SONG_TITLE = os.path.join(ASSETS_DIR, "song_title_landmark.png")
    IMG_MORE_OPTIONS = os.path.join(ASSETS_DIR, "more_options.png")
    IMG_BEST_RESULT = os.path.join(ASSETS_DIR, "best_result.png")
    IMG_SKIP_DUPLICATE = os.path.join(ASSETS_DIR, "skip_duplicate.png")
    IMG_DOWNLOAD = os.path.join(ASSETS_DIR, "download.png")
    IMG_ADD_ICON = os.path.join(ASSETS_DIR, "add_icon.png")
    
    # 图像识别配置
    IMAGE_CONFIDENCE = 0.8
    SCROLL_LIMIT = 25
    
    # UI 配置
    UI_WIDTH = 300
    UI_MIN_HEIGHT = 300
    PIXEL_FONTS = ["Zpix", "Press Start 2P", "Pixel Operator", "Fixedsys", "Terminal", "Consolas", "Courier New"]

    # 预览窗口配置
    PREVIEW_LINE_WIDTH = 1
    PREVIEW_CORNER_LENGTH = 20
    PREVIEW_FILL_ALPHA = 50  # 0-255
    SHOW_RECOGNITION_PREVIEW = True
    PREVIEW_MIN_VISIBLE_MS = 200
    MENU_POP_DELAY = 0.8

    FONT_ZPIX = os.path.join(FONTS_DIR, "最像素 Zpix.ttf")

    # 颜色主题 (Nord-like 低饱和度复古灰色系)
    COLOR_BG = "#2B2B2B"
    COLOR_PANEL = "#3C3F41"
    COLOR_BORDER = "#555555"
    COLOR_TEXT = "#A9B7C6"
    COLOR_MUTED = "#808080"
    COLOR_ACCENT = "#5E81AC"
    COLOR_ACCENT_ALT = "#D08770"
    COLOR_SUCCESS = "#8FA384"
    COLOR_WARNING = "#CDB77A"
    COLOR_DANGER = "#A56A72"
    COLOR_INFO = "#79A4B0"
    COLOR_BUTTON_BG = "#4E5254"
    COLOR_BUTTON_HOVER = "#5C6164"
    COLOR_BUTTON_ACTIVE = "#3C3F41"
    COLOR_TOOLTIP_BG = "#323232"


# Windows Hook 常量
WH_MOUSE_LL = 14
WM_RBUTTONDOWN = 0x0204
WM_RBUTTONUP = 0x0205
HC_ACTION = 0
LLMHF_INJECTED = 0x00000001
WM_APP_STOP = 0x8000 + 1001  # 自定义消息，用于瞬时通知主窗口停止


class MSLLHOOKSTRUCT(ctypes.Structure):
    """鼠标钩子结构体"""
    _fields_ = [
        ("pt", wintypes.POINT),
        ("mouseData", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", wintypes.ULONG)
    ]


class MouseHookManager:
    """
    简化的低级别鼠标钩子管理器
    功能：阻止物理鼠标输入，但允许自动化注入的输入
    """
    def __init__(self, stop_event, parent_hwnd=None, pause_event=None, clip_event=None, 
                 allow_click_event=None, lock_x=None, lock_y=None, blocking_event=None):
        self.stop_event = stop_event
        self.parent_hwnd = parent_hwnd
        self.paused = pause_event if pause_event else Event()
        self.clip_enabled = clip_event if clip_event else Event()
        self.allow_clicks = allow_click_event if allow_click_event else Event()
        self.blocking_enabled = blocking_event if blocking_event else Event()
        self.lock_x = lock_x
        self.lock_y = lock_y
        self.local_lock_pos = (0, 0)
        
        self._hook_id = None
        self._thread = None
        self._user32 = ctypes.windll.user32
        self._kernel32 = ctypes.windll.kernel32
        
        # 初始化默认值
        if not isinstance(self.clip_enabled, multiprocessing.synchronize.Event):
            self.clip_enabled.set()
        if not isinstance(self.blocking_enabled, multiprocessing.synchronize.Event):
            self.blocking_enabled.set()
        
        self._hook_proc = ctypes.WINFUNCTYPE(
            ctypes.c_long, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM
        )(self._low_level_mouse_proc)
        
        class _RECT(ctypes.Structure):
            _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long), 
                       ("right", ctypes.c_long), ("bottom", ctypes.c_long)]
        self._RECT = _RECT

    def _get_lock_pos(self):
        if self.lock_x is not None and self.lock_y is not None:
            return (self.lock_x.value, self.lock_y.value)
        return self.local_lock_pos

    def _apply_clip(self):
        if not self.clip_enabled.is_set() or self.paused.is_set():
            return
        x, y = self._get_lock_pos()
        rect = self._RECT(x, y, x + 1, y + 1)
        self._user32.ClipCursor(ctypes.byref(rect))

    def _remove_clip(self):
        self._user32.ClipCursor(None)

    def _low_level_mouse_proc(self, nCode, wParam, lParam):
        if nCode == HC_ACTION:
            try:
                # 如果已经触发停止，立即允许所有操作并解除限制
                if self.stop_event.is_set():
                    self._remove_clip()
                    return self._user32.CallNextHookEx(self._hook_id, nCode, wParam, lParam)

                struct = ctypes.cast(lParam, ctypes.POINTER(MSLLHOOKSTRUCT)).contents
                
                # 允许注入的事件（自动化）
                if struct.flags & LLMHF_INJECTED:
                    return self._user32.CallNextHookEx(self._hook_id, nCode, wParam, lParam)

                # 右键按下或 ESC：最高优先级瞬时终止
                if wParam == WM_RBUTTONDOWN:
                    self.stop_event.set()
                    self._remove_clip() 
                    if self.parent_hwnd:
                        self._user32.PostMessageW(self.parent_hwnd, WM_APP_STOP, 0, 0)
                    self._user32.SetCursorPos(0, 0) # 物理移动到左上角触发 PyAutoGUI FailSafe
                    return self._user32.CallNextHookEx(self._hook_id, nCode, wParam, lParam)

                # 阻止物理鼠标输入（仅在任务运行且未停止时）
                if not self.paused.is_set() and self.blocking_enabled.is_set():
                    if self.allow_clicks.is_set() and (wParam == 0x0201 or wParam == 0x0202):
                        return self._user32.CallNextHookEx(self._hook_id, nCode, wParam, lParam)
                    if wParam == WM_RBUTTONUP:
                        return self._user32.CallNextHookEx(self._hook_id, nCode, wParam, lParam)
                    return 1
            except:
                pass
        return self._user32.CallNextHookEx(self._hook_id, nCode, wParam, lParam)

    def run_blocking(self):
        pt = wintypes.POINT()
        self._user32.GetCursorPos(ctypes.byref(pt))
        if self.lock_x: self.lock_x.value = pt.x
        if self.lock_y: self.lock_y.value = pt.y
        self.local_lock_pos = (pt.x, pt.y)

        self._thread = Thread(target=self._hook_loop, daemon=True)
        self._thread.start()
        
        self._maint_thread = Thread(target=self._maintenance_loop, daemon=True)
        self._maint_thread.start()
        
        self._apply_clip()
        self._esc_poll_loop_blocking()

    def _maintenance_loop(self):
        while not self.stop_event.is_set():
            if not self.paused.is_set() and self.clip_enabled.is_set():
                try:
                    lx, ly = self._get_lock_pos()
                    # 再次检查，避免在触发停止后又瞬间锁定
                    if self.stop_event.is_set():
                        break
                    self._apply_clip()
                    pt = wintypes.POINT()
                    self._user32.GetCursorPos(ctypes.byref(pt))
                    if abs(pt.x - lx) > 2 or abs(pt.y - ly) > 2:
                        self._user32.SetCursorPos(lx, ly)
                except:
                    pass
            time.sleep(0.01)

    def _esc_poll_loop_blocking(self):
        while not self.stop_event.is_set():
            if self._user32.GetAsyncKeyState(0x1B) & 0x8000:
                self.stop_event.set()
                self._remove_clip() # 立即解除锁定
                if self.parent_hwnd:
                    self._user32.PostMessageW(self.parent_hwnd, WM_APP_STOP, 0, 0)
                break
            time.sleep(0.01) # 提高轮询频率
        self.disable()

    def _hook_loop(self):
        self._hook_id = self._user32.SetWindowsHookExW(
            WH_MOUSE_LL, self._hook_proc, self._kernel32.GetModuleHandleW(None), 0
        )
        msg = wintypes.MSG()
        while not self.stop_event.is_set():
            if self._user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
                self._user32.TranslateMessage(ctypes.byref(msg))
                self._user32.DispatchMessageW(ctypes.byref(msg))
            else:
                break
        if self._hook_id:
            self._user32.UnhookWindowsHookEx(self._hook_id)

    def disable(self):
        self._remove_clip()
        if self._thread and self._thread.is_alive():
            self._user32.PostThreadMessageW(self._thread.ident, 0x0000, 0, 0)


def run_mouse_guard_process(stop_event, parent_hwnd, pause_event, clip_event, allow_click_event, 
                           lock_x, lock_y, blocking_event):
    """鼠标保护进程入口函数"""
    manager = MouseHookManager(stop_event, parent_hwnd, pause_event, clip_event, allow_click_event, 
                               lock_x, lock_y, blocking_event)
    manager.run_blocking()



class ProcessMouseGuard:
    """进程级鼠标保护器"""
    def __init__(self, on_esc):
        self.on_esc = on_esc
        self.stop_event = MEvent()
        self.pause_event = MEvent()
        self.clip_event = MEvent()
        self.allow_click_event = MEvent()
        self.blocking_event = MEvent()
        self.lock_x = Value('i', 0)
        self.lock_y = Value('i', 0)
        self.process = None
        self.monitor_thread = None
        self.suppress_on_esc = False

    def enable(self, parent_hwnd=None):
        if self.process and self.process.is_alive():
            return
        self.suppress_on_esc = False
        self.stop_event.clear()
        self.pause_event.clear()
        self.clip_event.set()
        self.blocking_event.set()
        self.allow_click_event.clear()
        
        pt = wintypes.POINT()
        ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
        self.lock_x.value = pt.x
        self.lock_y.value = pt.y
        
        self.process = Process(
            target=run_mouse_guard_process,
            args=(self.stop_event, parent_hwnd, self.pause_event, self.clip_event, 
                  self.allow_click_event, self.lock_x, self.lock_y, self.blocking_event),
            daemon=True
        )
        self.process.start()
        self.monitor_thread = Thread(target=self._monitor_loop, daemon=True)
        self.monitor_thread.start()

    def _monitor_loop(self):
        while True:
            if self.stop_event.is_set():
                if not self.suppress_on_esc:
                    self.on_esc()
                break
            if not self.process or not self.process.is_alive():
                break
            time.sleep(0.05)

    def disable(self, silent=True):
        self.suppress_on_esc = silent
        self.stop_event.set()
        proc = self.process
        if proc:
            proc.join(timeout=1.0)
            if proc.is_alive():
                proc.terminate()
            self.process = None

    def pause(self):
        self.pause_event.set()
        self.clip_event.clear()
        self.blocking_event.clear()

    def resume(self):
        self.pause_event.clear()
        self.clip_event.set()
        self.blocking_event.set()
        
    def set_lock_position(self, x, y):
        self.lock_x.value = int(x)
        self.lock_y.value = int(y)

    def set_clip(self, active):
        self.clip_event.set() if active else self.clip_event.clear()
            
    def set_blocking(self, active):
        self.blocking_event.set() if active else self.blocking_event.clear()
            
    @property
    def allow_clicks(self):
        return self.allow_click_event.is_set()
        
    @allow_clicks.setter
    def allow_clicks(self, value):
        self.allow_click_event.set() if value else self.allow_click_event.clear()


class FramelessDialog(QtWidgets.QDialog):
    """无边框对话框基类"""
    def __init__(self, parent=None, title="Dialog"):
        super().__init__(parent)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.Dialog | QtCore.Qt.WindowStaysOnTopHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.drag_pos = None
        self.resize(600, 400)
        
        self.main_layout = QtWidgets.QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)
        
        self.container = QtWidgets.QWidget()
        self.container.setObjectName("DialogContainer")
        self.container.setStyleSheet(f"""
            QWidget#DialogContainer {{
                background-color: {AppConfig.COLOR_BG};
                border: 2px solid {AppConfig.COLOR_BORDER};
                border-radius: 0px;
            }}
        """)
        self.main_layout.addWidget(self.container)
        
        self.container_layout = QtWidgets.QVBoxLayout(self.container)
        self.container_layout.setContentsMargins(0, 0, 0, 0)
        self.container_layout.setSpacing(0)
        
        self.setup_title_bar(title)
        
        self.content_widget = QtWidgets.QWidget()
        self.content_layout = QtWidgets.QVBoxLayout(self.content_widget)
        self.content_layout.setContentsMargins(8, 8, 8, 8)
        self.content_layout.setSpacing(6)
        self.container_layout.addWidget(self.content_widget)

    def setup_title_bar(self, title):
        self.title_bar = QtWidgets.QWidget()
        self.title_bar.setFixedHeight(36)
        title_layout = QtWidgets.QHBoxLayout(self.title_bar)
        title_layout.setContentsMargins(8, 0, 8, 0)
        title_layout.setSpacing(8)
        
        title_label = QtWidgets.QLabel(title)
        title_label.setStyleSheet(f"color: {AppConfig.COLOR_TEXT}; font-weight: bold; font-family: '{AppConfig.PIXEL_FONTS[0]}';")
        title_layout.addWidget(title_label)
        title_layout.addStretch(1)
        
        btn_min = QtWidgets.QPushButton("_")
        btn_min.setFixedSize(24, 24)
        btn_min.setCursor(QtCore.Qt.PointingHandCursor)
        btn_min.clicked.connect(self.showMinimized)
        btn_min.setStyleSheet(f"""
            QPushButton {{
                background-color: {AppConfig.COLOR_PANEL}; 
                border: 1px solid {AppConfig.COLOR_BORDER}; 
                color: {AppConfig.COLOR_TEXT};
                font-family: '{AppConfig.PIXEL_FONTS[0]}';
                padding: 0px;
                padding-bottom: 4px;
                border-radius: 0px;
            }}
            QPushButton:hover {{ background-color: {AppConfig.COLOR_BUTTON_HOVER}; }}
        """)
        title_layout.addWidget(btn_min)

        btn_close = QtWidgets.QPushButton("X")
        btn_close.setFixedSize(24, 24)
        btn_close.setCursor(QtCore.Qt.PointingHandCursor)
        btn_close.clicked.connect(self.close)
        btn_close.setStyleSheet(f"""
            QPushButton {{
                background-color: {AppConfig.COLOR_DANGER}; 
                border: 1px solid {AppConfig.COLOR_BORDER}; 
                color: {AppConfig.COLOR_BG};
                font-family: '{AppConfig.PIXEL_FONTS[0]}';
                padding: 0px;
                border-radius: 0px;
            }}
            QPushButton:hover {{ background-color: #D08770; }}
        """)
        title_layout.addWidget(btn_close)
        self.container_layout.addWidget(self.title_bar)

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton and self.title_bar.underMouse():
            self.drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == QtCore.Qt.LeftButton and self.drag_pos:
            self.move(event.globalPosition().toPoint() - self.drag_pos)
            event.accept()

    def add_widget(self, widget):
        self.content_layout.addWidget(widget)

    def set_layout(self, layout):
        self.content_layout.addLayout(layout)


class PixelMessageBox(FramelessDialog):
    """像素风格消息框"""
    def __init__(self, parent, title, text, type="info"):
        super().__init__(parent, title)
        self.resize(400, 220)
        layout = QtWidgets.QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(20, 20, 20, 20)
        
        msg_label = QtWidgets.QLabel(text)
        msg_label.setWordWrap(True)
        msg_label.setAlignment(QtCore.Qt.AlignCenter)
        msg_label.setStyleSheet(f"font-size: 11pt; color: {AppConfig.COLOR_TEXT};")
        layout.addWidget(msg_label)
        
        btn_box = QtWidgets.QHBoxLayout()
        btn_box.addStretch(1)
        ok_btn = QtWidgets.QPushButton("确定")
        ok_btn.setFixedSize(100, 36)
        ok_btn.setCursor(QtCore.Qt.PointingHandCursor)
        ok_btn.clicked.connect(self.accept)
        
        color = {
            "error": AppConfig.COLOR_DANGER,
            "warning": AppConfig.COLOR_WARNING,
            "success": AppConfig.COLOR_SUCCESS
        }.get(type, AppConfig.COLOR_INFO)
            
        ok_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                color: {AppConfig.COLOR_BG};
                border: 2px solid {AppConfig.COLOR_BORDER};
                font-weight: bold;
                font-family: '{AppConfig.PIXEL_FONTS[0]}';
                border-radius: 0px;
            }}
            QPushButton:hover {{
                background-color: {AppConfig.COLOR_TEXT};
                border-color: {AppConfig.COLOR_TEXT};
                color: {AppConfig.COLOR_BG};
            }}
        """)
        
        btn_box.addWidget(ok_btn)
        btn_box.addStretch(1)
        layout.addLayout(btn_box)
        self.set_layout(layout)

    @staticmethod
    def show_info(parent, title, text):
        PixelMessageBox(parent, title, text, "info").exec()
        
    @staticmethod
    def show_warning(parent, title, text):
        PixelMessageBox(parent, title, text, "warning").exec()

    @staticmethod
    def show_error(parent, title, text):
        PixelMessageBox(parent, title, text, "error").exec()
        
    @staticmethod
    def show_success(parent, title, text):
        PixelMessageBox(parent, title, text, "success").exec()


class ScreenSnipper(QtWidgets.QDialog):
    """屏幕截图工具"""
    def __init__(self, owner, callback, on_close=None):
        super().__init__(None)
        self.owner = owner
        self.callback = callback
        self.on_close = on_close
        screen = self._resolve_screen(owner)
        self.screen_geometry = screen.geometry()
        self.monitor_rect = self._resolve_monitor_rect(owner, screen)
        full_img = self._grab_full_image(self.monitor_rect)
        if full_img is None:
            return
        self.scale_x = full_img.width / max(1, self.screen_geometry.width())
        self.scale_y = full_img.height / max(1, self.screen_geometry.height())
        self._init_images(full_img, self.screen_geometry.size())
        self.start_pos = None
        self.current_pos = None
        self._setup_ui()

    def _resolve_screen(self, owner):
        if owner and owner.windowHandle() and owner.windowHandle().screen():
            return owner.windowHandle().screen()
        return QtGui.QGuiApplication.primaryScreen()

    def _resolve_monitor_rect(self, owner, screen):
        monitors = win32api.EnumDisplayMonitors()
        if owner and owner.winId():
            try:
                hwnd = int(owner.winId())
                rect = win32gui.GetWindowRect(hwnd)
                cx, cy = (rect[0] + rect[2]) // 2, (rect[1] + rect[3]) // 2
                for _, _, mrect in monitors:
                    if mrect[0] <= cx < mrect[2] and mrect[1] <= cy < mrect[3]:
                        return mrect
            except:
                pass
        return monitors[0][2] if monitors else (screen.geometry().x(), screen.geometry().y(), 
                                                 screen.geometry().x() + screen.geometry().width(), 
                                                 screen.geometry().y() + screen.geometry().height())

    def _grab_full_image(self, monitor_rect):
        try:
            left, top, right, bottom = monitor_rect
            width, height = int(right - left), int(bottom - top)
            hwnd = win32gui.GetDesktopWindow()
            hwindc = win32gui.GetWindowDC(hwnd)
            srcdc = win32ui.CreateDCFromHandle(hwindc)
            memdc = srcdc.CreateCompatibleDC()
            bmp = win32ui.CreateBitmap()
            bmp.CreateCompatibleBitmap(srcdc, width, height)
            memdc.SelectObject(bmp)
            memdc.BitBlt((0, 0), (width, height), srcdc, (left, top), win32con.SRCCOPY)
            bmpinfo = bmp.GetInfo()
            bmpstr = bmp.GetBitmapBits(True)
            img = Image.frombuffer("RGB", (bmpinfo["bmWidth"], bmpinfo["bmHeight"]), 
                                  bmpstr, "raw", "BGRX", 0, 1)
            return img
        except Exception as e:
            PixelMessageBox.show_error(self.owner, "截图失败", f"无法捕获屏幕：{e}")
            if self.on_close: self.on_close()
            self.close()
            return None
        finally:
            try:
                if memdc: memdc.DeleteDC()
                if srcdc: srcdc.DeleteDC()
                if hwindc and hwnd: win32gui.ReleaseDC(hwnd, hwindc)
                if bmp: win32gui.DeleteObject(bmp.GetHandle())
            except:
                pass

    def _init_images(self, full_img, target_size):
        self.full_img = full_img
        try:
            enhancer = ImageEnhance.Brightness(self.full_img)
            self.dimmed_img = enhancer.enhance(0.5)
        except:
            self.dimmed_img = self.full_img
        dpr = min(self.scale_x, self.scale_y)
        self.full_pixmap = QtGui.QPixmap.fromImage(ImageQt.ImageQt(self.full_img))
        self.dimmed_pixmap = QtGui.QPixmap.fromImage(ImageQt.ImageQt(self.dimmed_img))
        self.full_pixmap.setDevicePixelRatio(dpr)
        self.dimmed_pixmap.setDevicePixelRatio(dpr)

    def _setup_ui(self):
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnTopHint)
        self.setGeometry(self.screen_geometry)
        self.setMouseTracking(True)

    def cancel(self):
        if self.on_close: self.on_close()
        self.close()

    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_Escape:
            self.cancel()

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.RightButton:
            self.cancel()
            return
        if event.button() == QtCore.Qt.LeftButton:
            self.start_pos = event.position().toPoint()
            self.current_pos = self.start_pos
            self.update()

    def mouseMoveEvent(self, event):
        if self.start_pos is None: return
        self.current_pos = event.position().toPoint()
        self.update()

    def mouseReleaseEvent(self, event):
        if self.start_pos is None or event.button() != QtCore.Qt.LeftButton: return
        end_pos = event.position().toPoint()
        rect = QtCore.QRect(self.start_pos, end_pos).normalized()
        self.start_pos = None
        self.current_pos = None
        self._finalize_selection(rect)

    def _finalize_selection(self, rect):
        if self.on_close: self.on_close()
        self.close()
        if rect.width() <= 5 or rect.height() <= 5: return
        try:
            x1 = max(0, int(rect.left() * self.scale_x))
            y1 = max(0, int(rect.top() * self.scale_y))
            x2 = min(self.full_img.width, int(rect.right() * self.scale_x))
            y2 = min(self.full_img.height, int(rect.bottom() * self.scale_y))
            img = self.full_img.crop((x1, y1, x2, y2))
            self.callback(img)
        except Exception as e:
            PixelMessageBox.show_error(self.parent(), "错误", f"裁剪图片失败：{e}")

    def paintEvent(self, event):
        painter = QtGui.QPainter(self)
        painter.drawPixmap(0, 0, self.dimmed_pixmap)
        if self.start_pos and self.current_pos:
            rect = QtCore.QRect(self.start_pos, self.current_pos).normalized()
            painter.setClipRect(rect)
            painter.drawPixmap(0, 0, self.full_pixmap)
            painter.setClipping(False)
            pen = QtGui.QPen(QtGui.QColor(AppConfig.COLOR_DANGER))
            pen.setWidth(3)
            painter.setPen(pen)
            painter.drawRect(rect)


class TargetPreviewWindow(QtWidgets.QWidget):
    """
    目标定位预览窗口
    功能：在屏幕上实时显示图像识别结果的框选线条
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(
            QtCore.Qt.FramelessWindowHint | 
            QtCore.Qt.WindowStaysOnTopHint | 
            QtCore.Qt.Tool |
            QtCore.Qt.WindowDoesNotAcceptFocus
        )
        try:
            self.setWindowFlag(QtCore.Qt.WindowTransparentForInput, True)
        except Exception:
            pass
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.setAttribute(QtCore.Qt.WA_TransparentForMouseEvents)
        self.setAttribute(QtCore.Qt.WA_ShowWithoutActivating, True)
        self.setFocusPolicy(QtCore.Qt.NoFocus)
        
        self.target_rect = None  # (x, y, width, height)
        self.target_label = ""
        self.line_color = QtGui.QColor(AppConfig.COLOR_DANGER)
        self.line_width = AppConfig.PREVIEW_LINE_WIDTH
        self.fill_color = QtGui.QColor(255, 0, 0, AppConfig.PREVIEW_FILL_ALPHA)  # 半透明红色填充
        self.corner_length = AppConfig.PREVIEW_CORNER_LENGTH
        self.min_visible_ms = AppConfig.PREVIEW_MIN_VISIBLE_MS
        self._last_update_ts = 0.0
        self._hide_timer = QtCore.QTimer(self)
        self._hide_timer.setSingleShot(True)
        self._hide_timer.timeout.connect(self._do_hide)
        
    def update_target(self, rect, label="", color=None):
        """
        更新目标框选位置
        :param rect: QRect 或 (x, y, w, h) 元组
        :param label: 可选的标签文本
        :param color: 可选的线条颜色
        """
        if isinstance(rect, tuple):
            rect = QtCore.QRect(rect[0], rect[1], rect[2], rect[3])
        
        self.target_label = label
        self._last_update_ts = time.monotonic()
        if self._hide_timer.isActive():
            self._hide_timer.stop()
        
        if color:
            self.line_color = color
        
        # 设置窗口几何位置以覆盖目标区域
        margin = 50
        preview_rect = rect.adjusted(-margin, -margin, margin, margin)
        self.setGeometry(preview_rect)
        local_rect = QtCore.QRect(rect)
        local_rect.translate(-preview_rect.left(), -preview_rect.top())
        self.target_rect = local_rect
        
        # 移动窗口使目标区域居中
        self.move(rect.left() - margin, rect.top() - margin)
        
        self.update()
        self.show()
      
    def clear_target(self):
        """清除框选显示"""
        if not self.target_rect:
            return
        elapsed_ms = (time.monotonic() - self._last_update_ts) * 1000.0
        remaining = self.min_visible_ms - elapsed_ms
        if remaining > 0:
            self._hide_timer.start(int(remaining))
            return
        self._do_hide()

    def _do_hide(self):
        self.target_rect = None
        self.target_label = ""
        self.hide()
    
    def paintEvent(self, event):
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        
        if self.target_rect:
            # 绘制半透明填充
            painter.setBrush(QtGui.QBrush(self.fill_color))
            painter.setPen(QtCore.Qt.NoPen)
            painter.drawRect(self.target_rect)
            
            # 绘制边框线条
            pen = QtGui.QPen(self.line_color)
            pen.setWidth(self.line_width)
            pen.setStyle(QtCore.Qt.SolidLine)
            painter.setPen(pen)
            painter.setBrush(QtCore.Qt.NoBrush)
            painter.drawRect(self.target_rect)
            
            # 绘制角标（增强可见性）
            corner_length = self.corner_length
            corner_pen = QtGui.QPen(self.line_color)
            corner_pen.setWidth(self.line_width + 1)
            painter.setPen(corner_pen)
            
            # 左上角
            tl = self.target_rect.topLeft()
            painter.drawLine(tl.x(), tl.y() + corner_length, tl.x(), tl.y())
            painter.drawLine(tl.x(), tl.y(), tl.x() + corner_length, tl.y())
            
            # 右上角
            tr = self.target_rect.topRight()
            painter.drawLine(tr.x() - corner_length, tr.y(), tr.x(), tr.y())
            painter.drawLine(tr.x(), tr.y(), tr.x(), tr.y() + corner_length)
            
            # 左下角
            bl = self.target_rect.bottomLeft()
            painter.drawLine(bl.x(), bl.y() - corner_length, bl.x(), bl.y())
            painter.drawLine(bl.x(), bl.y(), bl.x() + corner_length, bl.y())
            
            # 右下角
            br = self.target_rect.bottomRight()
            painter.drawLine(br.x() - corner_length, br.y(), br.x(), br.y())
            painter.drawLine(br.x(), br.y() - corner_length, br.x(), br.y())
            
            # 绘制标签文本
            if self.target_label:
                font = painter.font()
                font.setBold(True)
                font.setPointSize(12) # 字号稍大一点
                painter.setFont(font)
                
                # 计算文本位置
                fm = QtGui.QFontMetrics(font)
                text_width = fm.horizontalAdvance(self.target_label)
                text_height = fm.height()
                
                # 文本起始点 (x, y) - y 是基线位置
                text_x = self.target_rect.left()
                text_y = self.target_rect.top() - 8
                
                # 1. 首先绘制一个半透明深色背景条，确保在任何背景下文字都可见
                bg_rect = QtCore.QRect(text_x - 2, text_y - text_height + 2, text_width + 4, text_height)
                painter.setBrush(QtGui.QBrush(QtGui.QColor(0, 0, 0, 160))) # 深黑色半透明背景
                painter.setPen(QtCore.Qt.NoPen)
                painter.drawRect(bg_rect)
                
                # 2. 使用 QPainterPath 实现双层渲染（描边+填充）
                path = QtGui.QPainterPath()
                path.addText(text_x, text_y, font, self.target_label)
                
                # 绘制黑色外部轮廓（描边）
                stroke_pen = QtGui.QPen(QtGui.QColor(0, 0, 0))
                stroke_pen.setWidthF(2.5)
                stroke_pen.setJoinStyle(QtCore.Qt.RoundJoin) # 圆角连接更平滑
                painter.strokePath(path, stroke_pen)
                
                # 填充白色主体
                painter.fillPath(path, QtGui.QColor(255, 255, 255))



class MusicSearcherApp(QtWidgets.QWidget):
    """主应用程序类"""
    ui_callback = QtCore.Signal(object)
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle(AppConfig.VERSION)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnTopHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.setFixedWidth(AppConfig.UI_WIDTH)
        self.setMinimumHeight(AppConfig.UI_MIN_HEIGHT)
        self.normal_min_height = AppConfig.UI_MIN_HEIGHT
        self.normal_max_height = self.maximumHeight()
        self.normal_height = None
        self.drag_pos = None
        self.excel_path = AppConfig.EXCEL_PATH
        self.df = None
        self.pending_indices = []
        self.current_list_index = -1
        self.stop_event = Event()
        self.is_running = False
        self.skip_ui_reset = False
        self.run_token = 0
        self.after_timer = QtCore.QTimer(self)
        self.after_timer.setSingleShot(True)
        self.after_timer.timeout.connect(lambda: self.search_action(False))
        self.mouse_lock = ProcessMouseGuard(on_esc=self._on_esc_pressed)
        self.full_song_text = ""
        self.full_artist_text = ""
        self.song_scroll_pos = 0
        self.artist_scroll_pos = 0
        self.scroll_timer = QtCore.QTimer(self)
        self.scroll_timer.timeout.connect(self.scroll_labels)
        self.location_cache = {}
        self.countdown_flags = {}
        self.cycle_start_time = 0
        self.last_cycle_duration = 0
        self.last_more_options_offset = None  # 缓存三点图标相对于标题锚点的偏移 (dx, dy)
        self.preview_window = TargetPreviewWindow(self)
        self.ui_callback.connect(self._run_ui_callback)
        self._hwnd = int(self.winId())
        self.setup_styles()
        self.setup_ui()
        self.load_excel()
        self.adjustSize()

    def nativeEvent(self, eventType, message):
        if eventType == "windows_generic_MSG":
            msg = wintypes.MSG.from_address(int(message))
            if msg.message == WM_APP_STOP:
                self._on_esc_pressed()
                return True, 0
        return super().nativeEvent(eventType, message)

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton and hasattr(self, 'title_bar') and self.title_bar.underMouse():
            self.drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == QtCore.Qt.LeftButton and self.drag_pos:
            self.move(event.globalPosition().toPoint() - self.drag_pos)
            event.accept()

    def run_on_ui(self, func):
        self.ui_callback.emit(func)

    def _run_ui_callback(self, func):
        func()

    def apply_pixel_shadow(self, widget, color="#151218", offset=2):
        shadow = QtWidgets.QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(0)
        shadow.setColor(QtGui.QColor(color))
        shadow.setOffset(offset, offset)
        widget.setGraphicsEffect(shadow)

    def setup_styles(self):
        font_id = QtGui.QFontDatabase.addApplicationFont(AppConfig.FONT_ZPIX)
        font_family = QtGui.QFontDatabase.applicationFontFamilies(font_id)[0] if font_id != -1 else \
                     next((f for f in AppConfig.PIXEL_FONTS if f in QtGui.QFontDatabase().families()), "Consolas")
        self.font_base = QtGui.QFont(font_family, 10)
        self.font_title = QtGui.QFont(font_family, 12, QtGui.QFont.Bold)
        self.font_small = QtGui.QFont(font_family, 9)
        self.setFont(self.font_base)
        
        self.setStyleSheet(f"""
        QWidget {{ background-color: {AppConfig.COLOR_BG}; color: {AppConfig.COLOR_TEXT}; font-family: "{font_family}"; font-size: 10pt; }}
        QGroupBox {{ background-color: {AppConfig.COLOR_PANEL}; border: 2px solid {AppConfig.COLOR_BORDER}; margin-top: 10px; font-weight: bold; border-radius: 0px; }}
        QGroupBox::title {{ subcontrol-origin: margin; subcontrol-position: top left; padding: 0 6px; color: {AppConfig.COLOR_ACCENT_ALT}; background-color: {AppConfig.COLOR_PANEL}; }}
        QLabel#TitleLabel {{ font-size: 11pt; font-weight: bold; color: {AppConfig.COLOR_ACCENT_ALT}; }}
        QLabel#MutedLabel {{ color: {AppConfig.COLOR_MUTED}; }}
        QLabel#SuccessLabel {{ color: {AppConfig.COLOR_SUCCESS}; font-weight: bold; }}
        QLabel#StatusLabel {{ background-color: {AppConfig.COLOR_BG}; color: {AppConfig.COLOR_WARNING}; font-size: 9pt; padding: 4px; }}
        QPushButton {{ background-color: {AppConfig.COLOR_BUTTON_BG}; border: 2px solid {AppConfig.COLOR_BORDER}; padding: 6px 10px; color: {AppConfig.COLOR_TEXT}; outline: none; border-radius: 0px; }}
        QPushButton[innerBorder="true"] {{ outline: 1px solid {AppConfig.COLOR_BORDER}; outline-offset: -4px; }}
        QPushButton:hover {{ background-color: {AppConfig.COLOR_BUTTON_HOVER}; border-color: {AppConfig.COLOR_ACCENT}; }}
        QPushButton:pressed {{ background-color: {AppConfig.COLOR_BUTTON_ACTIVE}; border-color: {AppConfig.COLOR_ACCENT_ALT}; }}
        QPushButton:disabled {{ background-color: {AppConfig.COLOR_PANEL}; border-color: {AppConfig.COLOR_BORDER}; color: {AppConfig.COLOR_MUTED}; }}
        QPushButton#ActionButton {{ background-color: {AppConfig.COLOR_ACCENT}; color: {AppConfig.COLOR_BG}; font-weight: bold; border-color: {AppConfig.COLOR_ACCENT}; }}
        QPushButton#SuccessButton {{ background-color: {AppConfig.COLOR_SUCCESS}; color: {AppConfig.COLOR_BG}; border-color: {AppConfig.COLOR_SUCCESS}; font-weight: bold; }}
        QPushButton#SuccessButton:hover {{ background-color: {AppConfig.COLOR_SUCCESS}; border-color: {AppConfig.COLOR_ACCENT_ALT}; }}
        QPushButton#WarningButton {{ background-color: {AppConfig.COLOR_WARNING}; color: {AppConfig.COLOR_BG}; border-color: {AppConfig.COLOR_WARNING}; font-weight: bold; }}
        QPushButton#WarningButton:hover {{ background-color: {AppConfig.COLOR_WARNING}; border-color: {AppConfig.COLOR_ACCENT_ALT}; }}
        QPushButton#DangerButton {{ background-color: {AppConfig.COLOR_DANGER}; color: {AppConfig.COLOR_BG}; border-color: {AppConfig.COLOR_DANGER}; font-weight: bold; }}
        QPushButton#DangerButton:hover {{ background-color: {AppConfig.COLOR_DANGER}; border-color: {AppConfig.COLOR_ACCENT_ALT}; }}
        QPushButton#InfoButton {{ background-color: {AppConfig.COLOR_INFO}; color: {AppConfig.COLOR_BG}; border-color: {AppConfig.COLOR_INFO}; font-weight: bold; }}
        QPushButton#InfoButton:hover {{ background-color: {AppConfig.COLOR_INFO}; border-color: {AppConfig.COLOR_ACCENT_ALT}; }}
        QToolTip {{ background-color: {AppConfig.COLOR_TOOLTIP_BG}; color: {AppConfig.COLOR_TEXT}; border: 2px solid {AppConfig.COLOR_BORDER}; }}
        QTableWidget {{ background-color: {AppConfig.COLOR_PANEL}; border: 2px solid {AppConfig.COLOR_BORDER}; gridline-color: {AppConfig.COLOR_BORDER}; color: {AppConfig.COLOR_TEXT}; border-radius: 0px; }}
        QHeaderView::section {{ background-color: {AppConfig.COLOR_BUTTON_BG}; border: 2px solid {AppConfig.COLOR_BORDER}; padding: 4px; color: {AppConfig.COLOR_TEXT}; border-radius: 0px; }}
        """)

    def setup_ui(self):
        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        self.container = QtWidgets.QWidget()
        self.container.setObjectName("MainContainer")
        self.container.setStyleSheet(f"""
            QWidget#MainContainer {{ border: 2px solid {AppConfig.COLOR_BORDER}; background-color: {AppConfig.COLOR_BG}; border-radius: 0px; }}
        """)
        container_layout = QtWidgets.QVBoxLayout(self.container)
        container_layout.setContentsMargins(8, 8, 8, 8)
        container_layout.setSpacing(6)
        main_layout.addWidget(self.container)

        self.title_bar = QtWidgets.QWidget()
        self.title_bar.setFixedHeight(32)
        title_layout = QtWidgets.QHBoxLayout(self.title_bar)
        title_layout.setContentsMargins(4, 0, 4, 0)
        title_layout.setSpacing(8)
        
        app_title = QtWidgets.QLabel(f"MUSIC SEARCHER {AppConfig.VERSION}")
        app_title.setStyleSheet(f"color: {AppConfig.COLOR_ACCENT_ALT}; font-weight: bold;")
        self.apply_pixel_shadow(app_title, AppConfig.COLOR_BG, 2)
        title_layout.addWidget(app_title)
        title_layout.addStretch(1)
        
        btn_min = QtWidgets.QPushButton("_")
        btn_min.setFixedSize(24, 24)
        btn_min.setCursor(QtCore.Qt.PointingHandCursor)
        btn_min.clicked.connect(self.showMinimized)
        btn_min.setStyleSheet(f"""
            QPushButton {{ background-color: {AppConfig.COLOR_PANEL}; border: 1px solid {AppConfig.COLOR_BORDER}; color: {AppConfig.COLOR_TEXT}; font-family: '{AppConfig.PIXEL_FONTS[0]}'; padding: 0px; border-radius: 0px; }}
            QPushButton:hover {{ background-color: {AppConfig.COLOR_BUTTON_HOVER}; }}
        """)
        title_layout.addWidget(btn_min)

        btn_close = QtWidgets.QPushButton("X")
        btn_close.setFixedSize(24, 24)
        btn_close.setCursor(QtCore.Qt.PointingHandCursor)
        btn_close.clicked.connect(self.close)
        btn_close.setStyleSheet(f"""
            QPushButton {{ background-color: {AppConfig.COLOR_DANGER}; border: 1px solid {AppConfig.COLOR_BORDER}; color: {AppConfig.COLOR_BG}; font-family: '{AppConfig.PIXEL_FONTS[0]}'; padding: 0px; border-radius: 0px; }}
            QPushButton:hover {{ background-color: #D08770; }}
        """)
        title_layout.addWidget(btn_close)
        container_layout.addWidget(self.title_bar)

        self.info_group = QtWidgets.QGroupBox("歌曲信息")
        info_layout = QtWidgets.QVBoxLayout(self.info_group)
        info_layout.setSpacing(4)
        info_layout.setContentsMargins(8, 16, 8, 8)
        self.song_label = QtWidgets.QLabel("歌曲：-")
        self.song_label.setObjectName("TitleLabel")
        self.apply_pixel_shadow(self.song_label, AppConfig.COLOR_BG, 1)
        info_layout.addWidget(self.song_label)
        self.artist_label = QtWidgets.QLabel("歌手：-")
        self.artist_label.setObjectName("MutedLabel")
        info_layout.addWidget(self.artist_label)
        self.progress_label = QtWidgets.QLabel("进度：- / - (0.0%)")
        info_layout.addWidget(self.progress_label)
        container_layout.addWidget(self.info_group)

        self.search_group = QtWidgets.QGroupBox("搜索控制")
        search_layout = QtWidgets.QGridLayout(self.search_group)
        search_layout.setSpacing(4)
        search_layout.setContentsMargins(8, 16, 8, 8)
        self.btn_search_song = QtWidgets.QPushButton("搜索歌名")
        self.btn_search_song.setObjectName("ActionButton")
        self.btn_search_song.setFocusPolicy(QtCore.Qt.NoFocus)
        self.btn_search_song.setProperty("innerBorder", True)
        self.btn_search_song.clicked.connect(lambda: self.search_action(False))
        search_layout.addWidget(self.btn_search_song, 0, 0)
        self.btn_search_full = QtWidgets.QPushButton("搜索全名")
        self.btn_search_full.setObjectName("ActionButton")
        self.btn_search_full.setFocusPolicy(QtCore.Qt.NoFocus)
        self.btn_search_full.setProperty("innerBorder", True)
        self.btn_search_full.clicked.connect(lambda: self.search_action(True))
        search_layout.addWidget(self.btn_search_full, 0, 1)
        self.btn_prev = QtWidgets.QPushButton("上一首")
        self.btn_prev.setFocusPolicy(QtCore.Qt.NoFocus)
        self.btn_prev.setProperty("innerBorder", True)
        self.btn_prev.clicked.connect(self.prev_song)
        search_layout.addWidget(self.btn_prev, 1, 0)
        self.btn_preview = QtWidgets.QPushButton("文档列表")
        self.btn_preview.setFocusPolicy(QtCore.Qt.NoFocus)
        self.btn_preview.setProperty("innerBorder", True)
        self.btn_preview.clicked.connect(self.show_preview)
        search_layout.addWidget(self.btn_preview, 1, 1)
        container_layout.addWidget(self.search_group)
        self.search_group_size_policy = self.search_group.sizePolicy()
        self.search_group_max_height = self.search_group.maximumHeight()

        self.mark_group = QtWidgets.QGroupBox("状态标记")
        mark_layout = QtWidgets.QGridLayout(self.mark_group)
        mark_layout.setSpacing(4)
        mark_layout.setContentsMargins(8, 16, 8, 8)
        self.btn_yisl = QtWidgets.QPushButton("已收录")
        self.btn_yisl.setObjectName("SuccessButton")
        self.btn_yisl.setFocusPolicy(QtCore.Qt.NoFocus)
        self.btn_yisl.setProperty("innerBorder", True)
        self.btn_yisl.clicked.connect(lambda: self.mark_only("已收录"))
        mark_layout.addWidget(self.btn_yisl, 0, 0)
        self.btn_fyuq = QtWidgets.QPushButton("非原曲")
        self.btn_fyuq.setObjectName("WarningButton")
        self.btn_fyuq.setFocusPolicy(QtCore.Qt.NoFocus)
        self.btn_fyuq.setProperty("innerBorder", True)
        self.btn_fyuq.clicked.connect(lambda: self.mark_only("非原曲"))
        mark_layout.addWidget(self.btn_fyuq, 0, 1)
        self.btn_wsl = QtWidgets.QPushButton("未收录")
        self.btn_wsl.setObjectName("DangerButton")
        self.btn_wsl.setFocusPolicy(QtCore.Qt.NoFocus)
        self.btn_wsl.setProperty("innerBorder", True)
        self.btn_wsl.clicked.connect(lambda: self.mark_only("未收录"))
        mark_layout.addWidget(self.btn_wsl, 1, 0)
        self.btn_bsl = QtWidgets.QPushButton("不收录")
        self.btn_bsl.setObjectName("InfoButton")
        self.btn_bsl.setFocusPolicy(QtCore.Qt.NoFocus)
        self.btn_bsl.setProperty("innerBorder", True)
        self.btn_bsl.clicked.connect(lambda: self.mark_only("不收录"))
        mark_layout.addWidget(self.btn_bsl, 1, 1)
        self.btn_img_mgmt = QtWidgets.QPushButton("识别库")
        self.btn_img_mgmt.setFocusPolicy(QtCore.Qt.NoFocus)
        self.btn_img_mgmt.setProperty("innerBorder", True)
        self.btn_img_mgmt.clicked.connect(self.open_image_management)
        mark_layout.addWidget(self.btn_img_mgmt, 2, 0, 1, 2)
        container_layout.addWidget(self.mark_group)
        self.mark_group_size_policy = self.mark_group.sizePolicy()
        self.mark_group_max_height = self.mark_group.maximumHeight()

        self.status_label = QtWidgets.QLabel("状态：准备就绪")
        self.status_label.setObjectName("StatusLabel")
        self.apply_pixel_shadow(self.status_label, AppConfig.COLOR_BG, 1)
        container_layout.addWidget(self.status_label)
        self.song_label.setToolTip(self.full_song_text)
        self.artist_label.setToolTip(self.full_artist_text)

    def set_status(self, text, color=AppConfig.COLOR_MUTED):
        try:
            ts = time.strftime("%H:%M:%S")
            print(f"[{ts}] {text}", flush=True)
        except:
            pass
        def _update():
            self.status_label.setText(f"状态：{text}")
            self.status_label.setStyleSheet(f"color: {color};")
        self.run_on_ui(_update)

    def _apply_running_ui(self, is_running):
        def _update_ui():
            self.search_group.setVisible(not is_running)
            self.mark_group.setVisible(not is_running)
            if not is_running:
                self.search_group.setMaximumHeight(self.search_group_max_height)
                self.mark_group.setMaximumHeight(self.mark_group_max_height)
                self.search_group.setSizePolicy(self.search_group_size_policy)
                self.mark_group.setSizePolicy(self.mark_group_size_policy)
                self.search_group.show()
                self.mark_group.show()
                self.setMinimumHeight(self.normal_min_height)
                self.setMaximumHeight(self.normal_max_height)
            else:
                self.search_group.setMaximumHeight(0)
                self.mark_group.setMaximumHeight(0)
                self.search_group.setSizePolicy(QtWidgets.QSizePolicy.Ignored, QtWidgets.QSizePolicy.Ignored)
                self.mark_group.setSizePolicy(QtWidgets.QSizePolicy.Ignored, QtWidgets.QSizePolicy.Ignored)
                self.setMinimumHeight(0)
                self.setMaximumHeight(16777215)
            self.layout().activate()
            self.container.layout().activate()
            self.setFixedWidth(AppConfig.UI_WIDTH)
            self.container.setFixedWidth(AppConfig.UI_WIDTH)
            target_height = self.layout().sizeHint().height()
            if is_running:
                self.setFixedHeight(target_height)
            else:
                self.setFixedWidth(AppConfig.UI_WIDTH)
                if self.normal_height:
                    self.resize(AppConfig.UI_WIDTH, self.normal_height)
            self.update()
            self.repaint()
        self.run_on_ui(_update_ui)

    def enter_running_mode(self):
        self.run_token += 1
        token = self.run_token
        if not self.is_running:
            self.normal_height = self.height()
        self.is_running = True
        self._apply_running_ui(True)
        return token

    def exit_running_mode(self, token=None, force=False):
        if self.skip_ui_reset and not force:
            return
        if force:
            if self.skip_ui_reset:
                return
            self.run_token += 1
            self.is_running = False
            self.setMinimumHeight(self.normal_min_height)
            self.setMaximumHeight(self.normal_max_height)
            self._apply_running_ui(False)
            return
        if token is not None and token != self.run_token:
            return
        if not self.is_running:
            self.setMinimumHeight(self.normal_min_height)
            self.setMaximumHeight(self.normal_max_height)
            self._apply_running_ui(False)
            return
        self.is_running = False
        self.setMinimumHeight(self.normal_min_height)
        self.setMaximumHeight(self.normal_max_height)
        self._apply_running_ui(False)

    def set_running_mode(self, is_running):
        if is_running:
            self.enter_running_mode()
        else:
            self.exit_running_mode(force=True)

    def open_image_management(self):
        self.mgmt_dialog = FramelessDialog(self, "识别库管理")
        self.mgmt_dialog.resize(520, 600)
        self.mgmt_dialog.setWindowModality(QtCore.Qt.ApplicationModal)
        
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"""
            QScrollArea {{ border: none; background-color: {AppConfig.COLOR_BG}; }}
            QScrollBar:vertical {{ background: {AppConfig.COLOR_PANEL}; width: 12px; margin: 0px; border-left: 1px solid {AppConfig.COLOR_BORDER}; }}
            QScrollBar::handle:vertical {{ background: {AppConfig.COLOR_ACCENT}; min-height: 30px; margin: 0px; border-radius: 0px; }}
            QScrollBar::handle:vertical:hover {{ background: {AppConfig.COLOR_ACCENT_ALT}; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: none; }}
        """)
        
        container = QtWidgets.QWidget()
        container.setStyleSheet(f"background-color: {AppConfig.COLOR_BG};")
        layout = QtWidgets.QVBoxLayout(container)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        items = [
            (AppConfig.IMG_BEST_RESULT, "识别「最佳结果」", "用于确认搜索成功"),
            (AppConfig.IMG_SONG_TITLE, "识别「歌曲」", "用于定位歌曲列表锚点"),
            (AppConfig.IMG_DOWNLOAD, "识别「下载图标」", "行定位参考：云端下载图标"),
            (AppConfig.IMG_ADD_ICON, "识别「添加图标」", "行定位参考：加号图标"),
            (AppConfig.IMG_MORE_OPTIONS, "点击「 ··· 」", "歌曲右侧的操作菜单按钮"),
            (AppConfig.IMG_ADD_TO_PLAYLIST, "点击「添加到播放列表」", "右键菜单中的选项"),
            (AppConfig.IMG_TARGET_PLAYLIST, "点击「目标歌单」", "要添加到的目标歌单名称"),
            (AppConfig.IMG_SKIP_DUPLICATE, "识别「跳过」", "遇到重复歌曲时的跳过按钮")
        ]

        for filename, title, subtitle in items:
            card = QtWidgets.QFrame()
            card.setObjectName("MgmtCard")
            card.setStyleSheet(f"""
                #MgmtCard {{ background-color: {AppConfig.COLOR_PANEL}; border: 1px solid {AppConfig.COLOR_BORDER}; border-radius: 0px; }}
                #MgmtCard:hover {{ border-color: {AppConfig.COLOR_ACCENT}; background-color: {AppConfig.COLOR_BG}; }}
            """)
            card_layout = QtWidgets.QHBoxLayout(card)
            card_layout.setContentsMargins(10, 10, 10, 10)
            card_layout.setSpacing(12)

            preview_label = QtWidgets.QLabel()
            preview_label.setFixedSize(80, 50)
            preview_label.setAlignment(QtCore.Qt.AlignCenter)
            preview_label.setStyleSheet(f"background-color: {AppConfig.COLOR_BG}; border: 1px solid {AppConfig.COLOR_BORDER}; color: {AppConfig.COLOR_MUTED}; font-size: 8pt; border-radius: 0px;")
            card_layout.addWidget(preview_label)

            info_layout = QtWidgets.QVBoxLayout()
            info_layout.setSpacing(2)
            title_row = QtWidgets.QHBoxLayout()
            title_row.setSpacing(8)
            title_lbl = QtWidgets.QLabel(title)
            title_lbl.setStyleSheet(f"color: {AppConfig.COLOR_TEXT}; font-weight: bold; font-size: 10pt;")
            title_row.addWidget(title_lbl)
            title_row.addStretch(1)
            desc_lbl = QtWidgets.QLabel(subtitle)
            desc_lbl.setStyleSheet(f"color: {AppConfig.COLOR_MUTED}; font-size: 8pt;")
            info_layout.addLayout(title_row)
            info_layout.addWidget(desc_lbl)
            card_layout.addLayout(info_layout, stretch=1)

            action_layout = QtWidgets.QVBoxLayout()
            action_layout.setSpacing(4)
            cap_btn = QtWidgets.QPushButton("截取")
            cap_btn.setFixedSize(56, 26)
            cap_btn.setFocusPolicy(QtCore.Qt.NoFocus)
            cap_btn.setProperty("innerBorder", True)
            cap_btn.setStyleSheet(f"border-radius: 0px;")
            del_btn = QtWidgets.QPushButton("删除")
            del_btn.setFixedSize(56, 26)
            del_btn.setFocusPolicy(QtCore.Qt.NoFocus)
            del_btn.setProperty("innerBorder", True)
            del_btn.setStyleSheet(f"background-color: {AppConfig.COLOR_DANGER}; color: {AppConfig.COLOR_BG}; border-radius: 0px;")
            action_layout.addWidget(cap_btn)
            action_layout.addWidget(del_btn)
            card_layout.addLayout(action_layout)

            def make_update_func(f, pl, tl):
                def update():
                    exists = os.path.exists(f)
                    if exists:
                        pixmap = QtGui.QPixmap(f)
                        if not pixmap.isNull():
                            pl.setPixmap(pixmap.scaled(pl.size(), QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation))
                        else:
                            pl.setText("损坏")
                        tl.setStyleSheet(f"color: {AppConfig.COLOR_SUCCESS}; font-weight: bold; font-size: 10pt;")
                    else:
                        pl.clear()
                        pl.setText("无图")
                        tl.setStyleSheet(f"color: {AppConfig.COLOR_DANGER}; font-weight: bold; font-size: 10pt;")
                return update

            update_func = make_update_func(filename, preview_label, title_lbl)
            update_func()
            cap_btn.clicked.connect(lambda _, f=filename, uf=update_func, b=cap_btn: self.start_countdown_capture(self.mgmt_dialog, f, uf, b))
            del_btn.clicked.connect(lambda _, f=filename, uf=update_func: self.delete_image(f, uf))
            layout.addWidget(card)

        layout.addStretch(1)
        scroll.setWidget(container)
        main_layout = QtWidgets.QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.addWidget(scroll)
        self.mgmt_dialog.set_layout(main_layout)
        self.mgmt_dialog.show()

    def start_countdown_capture(self, parent_win, filename, update_callback, cap_btn):
        original_text = cap_btn.text()
        cap_btn.setEnabled(False)
        self.countdown_flags[filename] = True
        self._run_capture_countdown(5, cap_btn, original_text, parent_win, filename, update_callback)

    def _run_capture_countdown(self, seconds, cap_btn, original_text, parent_win, filename, update_callback):
        if not self.countdown_flags.get(filename, False):
            cap_btn.setText(original_text)
            cap_btn.setEnabled(True)
            return
        if seconds > 0:
            cap_btn.setText(f"{seconds}s")
            QtCore.QTimer.singleShot(1000, lambda: self._run_capture_countdown(seconds - 1, cap_btn, original_text, parent_win, filename, update_callback))
            return
        self.countdown_flags[filename] = False
        cap_btn.setText(original_text)
        cap_btn.setEnabled(True)
        self._launch_capture(parent_win, filename, update_callback)

    def _launch_capture(self, parent_win, filename, update_callback):
        self.hide()
        parent_win.hide()
        def on_close():
            self.show()
            parent_win.show()
        def do_capture():
            self.snipper = ScreenSnipper(self, callback=lambda img: self.save_screenshot_and_update(img, filename, update_callback), on_close=on_close)
            self.snipper.show()
        QtCore.QTimer.singleShot(300, do_capture)

    def save_screenshot_and_update(self, img, filename, update_callback):
        self.save_screenshot(img, filename)
        if callable(update_callback):
            update_callback()

    def delete_image(self, filename, update_callback):
        self.countdown_flags[filename] = False
        if os.path.exists(filename):
            try:
                os.remove(filename)
                if callable(update_callback):
                    update_callback()
            except Exception as e:
                PixelMessageBox.show_error(self, "错误", f"删除失败：{e}")

    def load_excel(self):
        if not os.path.exists(self.excel_path):
            PixelMessageBox.show_error(self, "错误", f"找不到文件：{self.excel_path}")
            return
        try:
            self.df = pd.read_excel(self.excel_path, header=None)
            if self.df.empty:
                self.set_status("Excel 文件为空", AppConfig.COLOR_DANGER)
                return
            if len(self.df.columns) < 2:
                self.df[1] = None
            self.pending_indices = self.df[self.df[1].isna()].index.tolist()
            if not self.pending_indices:
                self.current_list_index = len(self.pending_indices)
                self.set_status("所有歌曲已处理完毕", AppConfig.COLOR_SUCCESS)
            else:
                self.current_list_index = 0
                self.update_current_song_display()
                self.set_status("已加载，请点击搜索按钮", AppConfig.COLOR_SUCCESS)
        except Exception as e:
            PixelMessageBox.show_error(self, "加载失败", f"无法读取 Excel 文件，请确保文件格式正确且未被其他程序占用。\n\n错误详情：{e}")

    def update_current_song_display(self):
        self.song_scroll_pos = 0
        self.artist_scroll_pos = 0
        if self.scroll_timer.isActive():
            self.scroll_timer.stop()
        if 0 <= self.current_list_index < len(self.pending_indices):
            actual_index = self.pending_indices[self.current_list_index]
            val = self.df.iloc[actual_index, 0]
            full_name = str(val) if pd.notna(val) else ""
            if " - " in full_name:
                parts = full_name.split(" - ")
                song_name = parts[0].strip()
                artist_name = parts[-1].strip()
            else:
                song_name = full_name.strip()
                artist_name = "-"
            self.full_song_text = song_name if song_name else "-"
            self.full_artist_text = artist_name if artist_name else "-"
            self.song_label.setText(f"歌曲：{self.full_song_text}")
            self.artist_label.setText(f"歌手：{self.full_artist_text}")
            self.song_label.setToolTip(self.full_song_text)
            self.artist_label.setToolTip(self.full_artist_text)
            total_count = len(self.df)
            current_pos = actual_index + 1
            percentage = (current_pos / total_count * 100) if total_count > 0 else 0
            eta_suffix = ""
            remaining_songs = len(self.pending_indices) - self.current_list_index
            if self.last_cycle_duration > 0 and remaining_songs > 0:
                total_seconds = int(self.last_cycle_duration * remaining_songs)
                minutes = total_seconds // 60
                seconds = total_seconds % 60
                eta_suffix = f" (ETA {minutes:02d}:{seconds:02d})"
            self.progress_label.setText(f"进度：{current_pos} / {total_count} ({percentage:.1f}%){eta_suffix}")
            if len(self.full_song_text) > AppConfig.SCROLL_LIMIT or len(self.full_artist_text) > AppConfig.SCROLL_LIMIT:
                self.scroll_timer.start(250)
        else:
            self.song_label.setText("歌曲：-")
            self.artist_label.setText("歌手：-")
            self.progress_label.setText(f"进度：{len(self.df)} / {len(self.df)} (100.0%)")

    def scroll_labels(self):
        limit = AppConfig.SCROLL_LIMIT
        changed = False
        if len(self.full_song_text) > limit:
            self.song_scroll_pos = (self.song_scroll_pos + 1) % (len(self.full_song_text) + 6)
            combined = self.full_song_text + "      " + self.full_song_text
            display = combined[self.song_scroll_pos : self.song_scroll_pos + limit]
            self.song_label.setText(f"歌曲：{display}")
            changed = True
        else:
            self.song_label.setText(f"歌曲：{self.full_song_text}")
        if len(self.full_artist_text) > limit:
            self.artist_scroll_pos = (self.artist_scroll_pos + 1) % (len(self.full_artist_text) + 6)
            combined = self.full_artist_text + "      " + self.full_artist_text
            display = combined[self.artist_scroll_pos : self.artist_scroll_pos + limit]
            self.artist_label.setText(f"歌手：{display}")
            changed = True
        else:
            self.artist_label.setText(f"歌手：{self.full_artist_text}")
        if not changed and self.scroll_timer.isActive():
            self.scroll_timer.stop()

    def focus_apple_music(self):
        windows = []
        def callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if "Apple Music" in title or "iTunes" in title:
                    windows.append(hwnd)
        win32gui.EnumWindows(callback, windows)
        if windows:
            hwnd = windows[0]
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            try:
                win32gui.SetForegroundWindow(hwnd)
                return True
            except:
                pass
        return False

    def terminate_current_tasks(self):
        """强制终止当前任务"""
        self.stop_event.set()
        if self.after_timer.isActive():
            self.after_timer.stop()
        # 立即尝试从主线程释放鼠标，不等待进程关闭
        try:
            ctypes.windll.user32.ClipCursor(None)
        except:
            pass
        # 异步关闭保护进程，避免阻塞主线程导致几秒的延迟
        Thread(target=self.mouse_lock.disable, daemon=True).start()
        self.skip_ui_reset = False
        self.exit_running_mode(force=True)

    def _on_esc_pressed(self):
        """ESC/右键紧急终止"""
        self.stop_event.set()
        # 立即从主线程释放鼠标
        try:
            ctypes.windll.user32.ClipCursor(None)
        except:
            pass
        def _cleanup():
            self.set_status("任务已强制终止", AppConfig.COLOR_DANGER)
            self.terminate_current_tasks()
        self.run_on_ui(_cleanup)

    def check_assets_integrity(self):
        """检查识别库完整性"""
        required_assets = [
            (AppConfig.IMG_BEST_RESULT, "识别「最佳结果」"),
            (AppConfig.IMG_SONG_TITLE, "识别「歌曲」锚点"),
            (AppConfig.IMG_DOWNLOAD, "识别「下载图标」"),
            (AppConfig.IMG_ADD_ICON, "识别「添加图标」"),
            (AppConfig.IMG_MORE_OPTIONS, "点击「 ··· 」"),
            (AppConfig.IMG_ADD_TO_PLAYLIST, "点击「添加到播放列表」"),
            (AppConfig.IMG_TARGET_PLAYLIST, "点击「目标歌单」"),
            (AppConfig.IMG_SKIP_DUPLICATE, "识别「跳过」按钮")
        ]
        
        missing = []
        for path, name in required_assets:
            if not os.path.exists(path):
                missing.append(name)
        
        if missing:
            missing_str = "\n".join([f"· {m}" for m in missing])
            PixelMessageBox.show_error(
                self, 
                "识别库不完整", 
                f"以下识别模板尚未截取，请先进入「识别库」进行截取：\n\n{missing_str}"
            )
            return False
        return True

    def search_action(self, full_copy):
        self.terminate_current_tasks()
        
        # 强制完整性检查
        if not self.check_assets_integrity():
            return
            
        self.skip_ui_reset = False
        if self.current_list_index < 0 or self.current_list_index >= len(self.pending_indices):
            self.set_status("请先加载数据", AppConfig.COLOR_DANGER)
            return
        actual_index = self.pending_indices[self.current_list_index]
        full_name = str(self.df.iloc[actual_index, 0])
        if full_copy:
            parts = full_name.split(" - ")
            search_text = f"{parts[0].strip()} {parts[-1].strip()}" if len(parts) >= 2 else full_name.strip()
        else:
            search_text = full_name.split(" - ")[0].strip()
        self.stop_event.clear()
        self.cycle_start_time = time.time()
        token = self.enter_running_mode()
        self.mouse_lock.enable(self._hwnd)
        self.search_success = False

        def _thread_target():
            try:
                self._execute_search(search_text)
            finally:
                try:
                    self.mouse_lock.disable()
                except:
                    pass
                should_keep_ui = (
                    self.search_success 
                    and not self.stop_event.is_set() 
                    and (self.current_list_index + 1 < len(self.pending_indices))
                )
                if should_keep_ui:
                    self.skip_ui_reset = True
                else:
                    self.run_on_ui(lambda: self.exit_running_mode(token))

        Thread(target=_thread_target, daemon=True).start()

    def check_stop(self):
        """检查停止信号并抛出异常以实现瞬时终止"""
        if self.stop_event.is_set():
            raise TaskTerminated("User requested termination")

    def _execute_search(self, text):
        try:
            self.check_stop()
            pyperclip.copy(text)
            if not self.focus_apple_music():
                self.set_status("未找到 Apple Music 窗口", AppConfig.COLOR_DANGER)
                return
            
            hwnd_list = []
            def callback(h, l):
                if win32gui.IsWindowVisible(h) and ("Apple Music" in win32gui.GetWindowText(h) or "iTunes" in win32gui.GetWindowText(h)):
                    l.append(h)
            win32gui.EnumWindows(callback, hwnd_list)
            active_hwnd = hwnd_list[0] if hwnd_list else None
            
            self.set_status(f"正在搜索：{text}", AppConfig.COLOR_INFO)
            pyautogui.hotkey("ctrl", "f")
            self._monitored_sleep_throw(0.4)
            
            pyautogui.hotkey("ctrl", "a")
            pyautogui.press("backspace")
            self._monitored_sleep_throw(0.1)
            
            pyautogui.hotkey("ctrl", "v")
            self._monitored_sleep_throw(0.1)
            pyautogui.press("enter")
            self._monitored_sleep_throw(0.8)

            self.check_stop()
            if not os.path.exists(AppConfig.IMG_BEST_RESULT):
                self.set_status("识别库中缺失最佳结果图", AppConfig.COLOR_DANGER)
                return

            self.set_status("正在定位最佳结果...", AppConfig.COLOR_INFO)
            best_pos = self.find_and_operate_image_stabilized(AppConfig.IMG_BEST_RESULT, action="move", max_wait=6.0)
            if not best_pos:
                self.set_status("无法识别最佳结果", AppConfig.COLOR_DANGER)
                return

            self._monitored_sleep_throw(0.2)
            self.set_status("执行滚轮翻页...", AppConfig.COLOR_INFO)
            pagedown_success = self.pagedown_with_verify(active_hwnd)
            if not pagedown_success: return

            self.check_stop()
            self.set_status("正在定位歌曲位置...", AppConfig.COLOR_INFO)
            song_pos = self.find_and_operate_image(AppConfig.IMG_SONG_TITLE, action="none")
            if not song_pos:
                self.set_status("无法识别标题锚点", AppConfig.COLOR_DANGER)
                return

            # --- 动态垂直扫描 ---
            self.check_stop()
            more_pos = None
            if self.last_more_options_offset:
                dx, dy = self.last_more_options_offset
                quick_region = (int(song_pos.x() + dx - 60), int(song_pos.y() + dy - 60), 120, 120)
                more_pos = self.find_and_operate_image(AppConfig.IMG_MORE_OPTIONS, action="click", region=quick_region, max_wait=0.5)

            if not more_pos:
                self.check_stop()
                hover_x, hover_y = int(song_pos.x() + 200), int(song_pos.y() + 80)
                self.mouse_lock.set_lock_position(hover_x, hover_y)
                pyautogui.moveTo(hover_x, hover_y)
                self._monitored_sleep_throw(0.15)
                
                more_pos = self.find_nearest_more_options(song_pos)
                if more_pos:
                    self.check_stop()
                    self.find_and_operate_image(AppConfig.IMG_MORE_OPTIONS, action="click", region=(more_pos.x()-50, more_pos.y()-50, 100, 100), max_wait=0.3)
                    self.last_more_options_offset = (more_pos.x() - song_pos.x(), more_pos.y() - song_pos.y())

            if not more_pos:
                self.set_status("未能锁定操作按钮", AppConfig.COLOR_DANGER)
                return
            
            self._monitored_sleep_throw(AppConfig.MENU_POP_DELAY)
            self.check_stop()
            add_pos = self.find_and_operate_image_stabilized(AppConfig.IMG_ADD_TO_PLAYLIST, action="click", max_wait=2.5, stabilize_count=1)
            if not add_pos: return

            self._monitored_sleep_throw(0.2)
            self.check_stop()
            location = self.find_and_operate_image_stabilized(AppConfig.IMG_TARGET_PLAYLIST, action="click", max_wait=2.5)
            if not location: return

            self._monitored_sleep_throw(0.2)
            if os.path.exists(AppConfig.IMG_SKIP_DUPLICATE):
                self.find_and_operate_image_stabilized(AppConfig.IMG_SKIP_DUPLICATE, action="click", max_wait=0.4, stabilize_count=1)
            
            self.search_success = True
            self.run_on_ui(lambda: self.mark_and_next_internal("已收录"))
            
        except TaskTerminated:
            # 异常被抛出，静默退出或在 finally 中处理
            pass
        except Exception as e:
            self.set_status(f"执行异常: {e}", AppConfig.COLOR_DANGER)

    def _monitored_sleep_throw(self, duration):
        """带异常抛出的监控睡眠"""
        end_time = time.time() + duration
        while time.time() < end_time:
            self.check_stop()
            time.sleep(0.02)

    def _monitored_sleep(self, duration):
        end_time = time.time() + duration
        while time.time() < end_time:
            if self.stop_event.is_set():
                return True
            time.sleep(0.05)
        return False

    def _grab_full_image(self, monitor_rect=None):
        """
        获取屏幕截图
        使用 pyautogui 确保 DPI 兼容性
        返回：(PIL 图像，屏幕左偏移，屏幕上偏移)
        """
        try:
            if monitor_rect:
                # 截取指定区域
                left, top, right, bottom = monitor_rect
                img = pyautogui.screenshot(region=(left, top, right - left, bottom - top))
                return img, left, top
            else:
                # 截取整个屏幕（主显示器）
                img = pyautogui.screenshot()
                return img, 0, 0
        except Exception as e:
            print(f"Screenshot failed: {e}")
            return None, 0, 0

    def _smart_locate_image(self, image_path, confidence=0.8, region=None, use_cache=True, show_preview=None):
        """
        优化的图像定位方法
        完全废弃 pyautogui 的图片加载，避免编码问题
        """
        if show_preview is None:
            show_preview = AppConfig.SHOW_RECOGNITION_PREVIEW
        
        # 截图 - 获取整个屏幕
        haystack_img, screen_left, screen_top = self._grab_full_image()
        if haystack_img is None:
            return None
        
        # 搜索区域处理
        search_img = haystack_img
        offset_x, offset_y = screen_left, screen_top
        if region:
            rx, ry, rw, rh = region
            ix, iy = max(0, rx - screen_left), max(0, ry - screen_top)
            iw, ih = min(rw, haystack_img.width - ix), min(rh, haystack_img.height - iy)
            if iw > 10 and ih > 10:
                search_img = haystack_img.crop((ix, iy, ix + iw, iy + ih))
                offset_x, offset_y = screen_left + ix, screen_top + iy

        # 获取模板图像尺寸（用于预览）
        import cv2
        tpl_gray = self._cv2_read_gray(image_path, cv2)
        if tpl_gray is None:
            return None
        tpl_h, tpl_w = tpl_gray.shape[:2]

        # 仅使用我们的 OpenCV 逻辑，避免 pyautogui 触发 imread 告警
        location = self._cv2_locate(image_path, search_img, confidence)
        if location:
            center_x = offset_x + location[0]
            center_y = offset_y + location[1]
            
            if show_preview:
                box = (int(center_x - tpl_w/2), int(center_y - tpl_h/2), tpl_w, tpl_h)
                self._update_preview_box(box, image_path)
            
            return QtCore.QPoint(int(center_x), int(center_y))
        
        if show_preview:
            self._clear_preview_box()
        return None

    def _cache_region(self, image_path, left, top, width, height, padding=10):
        left = max(0, int(left - padding))
        top = max(0, int(top - padding))
        width = max(1, int(width + padding * 2))
        height = max(1, int(height + padding * 2))
        try:
            screen_w, screen_h = pyautogui.size()
            if left >= screen_w or top >= screen_h:
                return
            width = max(1, min(width, screen_w - left))
            height = max(1, min(height, screen_h - top))
        except Exception:
            pass
        self.location_cache[image_path] = (left, top, width, height)

    def _scale_rect_for_qt(self, rect):
        try:
            screen = QtGui.QGuiApplication.primaryScreen()
            if not screen:
                return rect
            logical_w = screen.geometry().width()
            logical_h = screen.geometry().height()
            if logical_w <= 0 or logical_h <= 0:
                return rect
            phys_w, phys_h = pyautogui.size()
            scale_x = phys_w / logical_w
            scale_y = phys_h / logical_h
            if abs(scale_x - 1.0) < 0.01 and abs(scale_y - 1.0) < 0.01:
                return rect
            x, y, w, h = rect
            return (int(x / scale_x), int(y / scale_y), int(w / scale_x), int(h / scale_y))
        except Exception:
            return rect

    def _clear_preview_box(self):
        self.run_on_ui(lambda: self.preview_window.clear_target())

    def _update_preview_box(self, box, label_text=""):
        """
        更新预览窗口的框选显示
        :param box: pyautogui.locate 返回的 Box 或 (x, y, w, h) 元组
        :param label_text: 显示的标签文本
        """
        if hasattr(box, 'left'):
            rect = (box.left, box.top, box.width, box.height)
        else:
            rect = box
        
        rect = self._scale_rect_for_qt(rect)
        label = os.path.basename(label_text) if label_text else ""
        self.run_on_ui(lambda r=rect, l=label: self.preview_window.update_target(r, l))

    def _cv2_read_gray(self, image_path, cv2_mod):
        # 增加实例级缓存以提高性能并避免 IO 编码问题
        if not hasattr(self, '_cv2_temp_cache'):
            self._cv2_temp_cache = {}
            
        if image_path in self._cv2_temp_cache:
            return self._cv2_temp_cache[image_path]
            
        if not os.path.exists(image_path):
            return None
            
        try:
            # 使用更稳健的方式读取带中文路径的图片
            with open(image_path, 'rb') as f:
                bytes_data = f.read()
            nparr = np.frombuffer(bytes_data, np.uint8)
            img = cv2_mod.imdecode(nparr, cv2_mod.IMREAD_GRAYSCALE)
            if img is not None:
                self._cv2_temp_cache[image_path] = img
                return img
        except Exception as e:
            try:
                ts = time.strftime("%H:%M:%S")
                print(f"[{ts}] Error loading template {image_path}: {e}")
            except:
                pass
            
        # 回退到原始方式（可能触发警告但作为最后手段）
        try:
            img = cv2_mod.imread(image_path, cv2_mod.IMREAD_GRAYSCALE)
            if img is not None:
                self._cv2_temp_cache[image_path] = img
                return img
        except:
            pass
            
        return None

    def _cv2_locate_all(self, image_path, haystack_img, confidence=0.8):
        """
        使用 OpenCV 找到所有匹配项
        返回：列表，包含中心点坐标 [(cx, cy), ...]
        """
        try:
            import cv2
            hay_np = np.array(haystack_img)
            if hay_np.ndim == 3:
                hay_gray = cv2.cvtColor(hay_np, cv2.COLOR_RGB2GRAY)
            else:
                hay_gray = hay_np
            
            tpl_gray = self._cv2_read_gray(image_path, cv2)
            if tpl_gray is None: return []
            
            th, tw = tpl_gray.shape[:2]
            res = cv2.matchTemplate(hay_gray, tpl_gray, cv2.TM_CCOEFF_NORMED)
            
            # 找到所有高于置信度的位置
            locs = np.where(res >= float(confidence))
            points = []
            
            # 简单的非极大值抑制（去重）
            used = []
            for pt in zip(*locs[::-1]):
                is_duplicate = False
                for up in used:
                    if abs(pt[0] - up[0]) < tw/2 and abs(pt[1] - up[1]) < th/2:
                        is_duplicate = True
                        break
                if not is_duplicate:
                    points.append((pt[0] + tw/2, pt[1] + th/2))
                    used.append(pt)
            return points
        except Exception as e:
            print(f"CV2 locate_all error: {e}")
            return []

    def find_nearest_more_options(self, ref_pos, search_region=None):
        """
        全重构：定位距离参考点最近的三点图标
        """
        self.set_status("正在全屏搜索最近的 [ ··· ] 按钮...", AppConfig.COLOR_INFO)
        
        # 截取大区域
        haystack, screen_left, screen_top = self._grab_full_image()
        if not haystack: return None
        
        # 如果提供了区域，则裁剪
        offset_x, offset_y = screen_left, screen_top
        search_img = haystack
        if search_region:
            rx, ry, rw, rh = search_region
            search_img = haystack.crop((rx - screen_left, ry - screen_top, rx + rw - screen_left, ry + rh - screen_top))
            offset_x, offset_y = rx, ry

        # 找到所有三点图标
        candidates = self._cv2_locate_all(AppConfig.IMG_MORE_OPTIONS, search_img, confidence=0.75)
        if not candidates:
            return None
        
        # 计算全局坐标并寻找最近的一个
        best_pos = None
        min_dist = float('inf')
        ref_x, ref_y = ref_pos.x(), ref_pos.y()
        
        for cx, cy in candidates:
            gx, gy = offset_x + cx, offset_y + cy
            # 计算距离（优先考虑垂直距离，因为它们应该在同一行或紧随其后）
            # 距离公式：dy * 2 + dx (因为跨行远比跨列远影响更大)
            dist = abs(gy - ref_y) * 2 + abs(gx - (ref_x + 500)) # 假设三点在右侧 500px 左右
            
            if dist < min_dist:
                min_dist = dist
                best_pos = QtCore.QPoint(int(gx), int(gy))
        
        if best_pos:
            self._update_preview_box((best_pos.x()-20, best_pos.y()-20, 40, 40), "最近 [···]")
            return best_pos
        return None

    def _cv2_locate(self, image_path, haystack_img, confidence=0.8):
        """
        使用 OpenCV 进行快速模板匹配
        返回：(center_x, center_y) 或 None
        """
        try:
            import cv2
            
            # 转换为 numpy 数组
            hay_np = np.array(haystack_img)
            if hay_np.ndim == 3:
                hay_gray = cv2.cvtColor(hay_np, cv2.COLOR_RGB2GRAY)
            else:
                hay_gray = hay_np
            
            # 读取模板
            tpl_gray = self._cv2_read_gray(image_path, cv2)
            if tpl_gray is None:
                return None
            
            th, tw = tpl_gray.shape[:2]
            hh, hw = hay_gray.shape[:2]
            
            # 模板不能大于搜索区域
            if th > hh or tw > hw:
                return None
            
            # 模板匹配
            res = cv2.matchTemplate(hay_gray, tpl_gray, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(res)
            
            if max_val >= float(confidence):
                # 返回中心点坐标
                cx = max_loc[0] + tw / 2.0
                cy = max_loc[1] + th / 2.0
                return (cx, cy)
        except Exception as e:
            try:
                ts = time.strftime("%H:%M:%S")
                print(f"[{ts}] CV2 locate error: {e}", flush=True)
            except:
                pass
        return None

    def find_and_operate_image_stabilized(self, image_path, action="click", confidence=None, max_wait=3.0, region=None, stabilize_count=2, show_preview=None):
        """
        稳定的图像查找和操作（带稳定性验证）
        优化：减少等待时间，添加视觉反馈
        :param show_preview: 是否显示预览窗口
        """
        if show_preview is None:
            show_preview = AppConfig.SHOW_RECOGNITION_PREVIEW

        if confidence is None:
            confidence = AppConfig.IMAGE_CONFIDENCE
        
        # 获取图片尺寸用于视觉反馈
        img_w, img_h = 40, 40
        if os.path.exists(image_path):
            try:
                with Image.open(image_path) as img:
                    img_w, img_h = img.size
            except:
                pass
        
        start_time = time.time()
        last_pos = None
        consecutive_matches = 0
        
        while time.time() - start_time < max_wait:
            if self.stop_event.is_set():
                if show_preview:
                    self._clear_preview_box()
                return None
            
            # 快速查找（不使用缓存验证，直接搜索）
            location = self._smart_locate_image(image_path, confidence=confidence, region=region, use_cache=False, show_preview=show_preview)
            
            if location:
                curr_pos = (int(location.x()), int(location.y()))
                
                # 调试输出
                try:
                    ts = time.strftime("%H:%M:%S")
                    print(f"[{ts}] 找到图像 {os.path.basename(image_path)}: 位置 ({curr_pos[0]}, {curr_pos[1]}), 图片尺寸 {img_w}x{img_h}", flush=True)
                except:
                    pass
                
                # 检查位置稳定性
                if last_pos and abs(curr_pos[0] - last_pos[0]) <= 2 and abs(curr_pos[1] - last_pos[1]) <= 2:
                    consecutive_matches += 1
                else:
                    consecutive_matches = 1
                
                last_pos = curr_pos
                
                # 达到稳定次数后执行操作
                if consecutive_matches >= stabilize_count:
                    try:
                        ts = time.strftime("%H:%M:%S")
                        img_name = os.path.basename(image_path)
                        print(f"[{ts}] locate stabilized: {img_name} action={action} x={curr_pos[0]} y={curr_pos[1]}", flush=True)
                    except:
                        pass
                    
                    # 执行操作
                    self._perform_action(location, action)
                    
                    # 操作完成后清除预览
                    if show_preview:
                        self._clear_preview_box()
                    
                    return location
            else:
                consecutive_matches = 0
                last_pos = None
            
            # 缩短等待时间到 50ms
            if self._monitored_sleep(0.05):
                if show_preview:
                    self._clear_preview_box()
                return None
        
        if show_preview:
            self._clear_preview_box()
        
        return None
 
    def find_and_operate_image(self, image_path, action="click", confidence=None, retries=5, region=None, max_wait=6.0, show_preview=None):
        """
        图像查找和操作（快速模式）
        优化：使用 OpenCV 优先，添加红色框选反馈
        :param show_preview: 是否显示预览窗口
        """
        if show_preview is None:
            show_preview = AppConfig.SHOW_RECOGNITION_PREVIEW

        if confidence is None:
            confidence = AppConfig.IMAGE_CONFIDENCE
        
        # 获取图片尺寸用于视觉反馈
        img_w, img_h = 40, 40
        if os.path.exists(image_path):
            try:
                with Image.open(image_path) as img:
                    img_w, img_h = img.size
            except:
                pass
        
        start_time = time.time()
        
        for i in range(retries):
            if self.stop_event.is_set():
                if show_preview:
                    self._clear_preview_box()
                return None
            
            if max_wait is not None and time.time() - start_time >= max_wait:
                if show_preview:
                    self._clear_preview_box()
                return None
            
            try:
                # 使用优化的查找方法
                location = self._smart_locate_image(image_path, confidence=confidence, region=region, use_cache=(i > 0), show_preview=show_preview)
                
                if location:
                    try:
                        ts = time.strftime("%H:%M:%S")
                        img_name = os.path.basename(image_path)
                        print(f"[{ts}] locate ok: {img_name} action={action} x={int(location.x())} y={int(location.y())} region={region}", flush=True)
                    except:
                        pass
                    
                    # 执行操作
                    self._perform_action(location, action)
                    
                    # 操作完成后清除预览
                    if show_preview:
                        self._clear_preview_box()
                    
                    return location
                    
            except Exception:
                try:
                    ts = time.strftime("%H:%M:%S")
                    img_name = os.path.basename(image_path)
                    print(f"[{ts}] locate/action exception: {img_name} action={action}", flush=True)
                except:
                    pass
                if show_preview:
                    self._clear_preview_box()
                return None
            
            # 缩短等待时间到 300ms
            if self._monitored_sleep(0.3):
                if show_preview:
                    self._clear_preview_box()
                return None
        
        if show_preview:
            self._clear_preview_box()
        
        return None

    def _perform_action(self, location, action):
        if isinstance(location, tuple) or isinstance(location, list):
            target_x, target_y = int(location[0]), int(location[1])
        else:
            target_x, target_y = int(location.x()), int(location.y())
        self.mouse_lock.set_lock_position(target_x, target_y)
        self.mouse_lock.set_clip(True)
        time.sleep(0.05)
        self.mouse_lock.set_blocking(False)
        time.sleep(0.02)
        try:
            if action == "click":
                try:
                    ts = time.strftime("%H:%M:%S")
                    print(f"[{ts}] do robust-click: x={target_x} y={target_y}", flush=True)
                except:
                    pass
                pyautogui.mouseDown()
                time.sleep(0.05)
                pyautogui.mouseUp()
            elif action == "move":
                try:
                    ts = time.strftime("%H:%M:%S")
                    print(f"[{ts}] do move: x={target_x} y={target_y}", flush=True)
                except:
                    pass
            elif action == "none":
                pass
        except Exception as e:
            print(f"Action failed: {e}")
        finally:
            self.mouse_lock.set_blocking(True)
            self.mouse_lock.set_lock_position(target_x, target_y)
            self.mouse_lock.set_clip(True)

    def pagedown_with_verify(self, hwnd=None, max_retries=3):
        """
        翻页并验证
        优化：添加红色框选显示搜索区域
        """
        song_title_path = AppConfig.IMG_SONG_TITLE
        if not os.path.exists(song_title_path):
            self.set_status(f"识别库缺失 {song_title_path}, 无法验证翻页", AppConfig.COLOR_DANGER)
            return False
        
        self.set_status("正在执行滚轮翻页...", AppConfig.COLOR_INFO)
        pyautogui.scroll(-1000)
        if self._monitored_sleep(0.8):
            return False
        
        for attempt in range(max_retries):
            if self.stop_event.is_set():
                return False
            
            self.set_status(f"正在探测标题锚点 ({attempt + 1}/{max_retries})...", AppConfig.COLOR_INFO)
            
            # 查找标题锚点
            location = self.find_and_operate_image(song_title_path, action="none", confidence=0.7)
            
            if location:
                self.set_status("翻页成功：检测到标题锚点", AppConfig.COLOR_SUCCESS)
                return True
            
            if attempt < max_retries - 1:
                self.set_status(f"未检测到锚点，第 {attempt + 1} 次重试翻页...", AppConfig.COLOR_WARNING)
                pyautogui.scroll(-1000)
                if self._monitored_sleep(1.0):
                    return False
        
        self.set_status("翻页失败：多次翻页后仍未检测到标题锚点", AppConfig.COLOR_DANGER)
        return False

    def save_screenshot(self, img, filename):
        try:
            img.save(filename)
            self.set_status(f"{filename} 已保存，请重新开始搜索", AppConfig.COLOR_SUCCESS)
        except Exception as e:
            self.set_status(f"保存失败 - {str(e)}", AppConfig.COLOR_DANGER)
            PixelMessageBox.show_error(self, "错误", f"保存截图失败：{str(e)}")

    def mark_only(self, status_text):
        """手动标记状态并切换下一首"""
        if self.current_list_index < 0 or self.current_list_index >= len(self.pending_indices):
            return
        if self.cycle_start_time > 0:
            self.last_cycle_duration = time.time() - self.cycle_start_time
        actual_index = self.pending_indices[self.current_list_index]
        try:
            self.df.iloc[actual_index, 1] = status_text
            self.df.to_excel(self.excel_path, index=False, header=False)
            self.set_status(f"已手动标记 '{status_text}'", AppConfig.COLOR_SUCCESS)
            if self.current_list_index < len(self.pending_indices) - 1:
                self.current_list_index += 1
                self.update_current_song_display()
            else:
                self.set_status("已经是最后一首了", AppConfig.COLOR_SUCCESS)
        except Exception as e:
            PixelMessageBox.show_error(self, "保存失败", f"请关闭 Excel 后重试。\n错误：{e}")

    def mark_and_next_internal(self, status_text):
        """自动标记并流转下一首"""
        if self.current_list_index < 0 or self.current_list_index >= len(self.pending_indices):
            return
        actual_index = self.pending_indices[self.current_list_index]
        try:
            self.df.iloc[actual_index, 1] = status_text
            self.df.to_excel(self.excel_path, index=False, header=False)
            self.set_status(f"已自动标记 '{status_text}'", AppConfig.COLOR_SUCCESS)
            if self.current_list_index < len(self.pending_indices) - 1:
                self.current_list_index += 1
                self.update_current_song_display()
                self.skip_ui_reset = True
                if self.cycle_start_time > 0:
                    self.last_cycle_duration = time.time() - self.cycle_start_time
                self.after_timer.start(1500)
            else:
                self.current_list_index = len(self.pending_indices)
                self.update_current_song_display()
                self.set_status("全部处理完毕", AppConfig.COLOR_SUCCESS)
                PixelMessageBox.show_success(self, "完成", "所有歌曲已处理完成！")
                self.skip_ui_reset = False
                self.exit_running_mode(force=True)
        except Exception as e:
            self.set_status(f"保存异常：{e}", AppConfig.COLOR_DANGER)
            self.exit_running_mode(force=True)

    def prev_song(self):
        self.terminate_current_tasks()
        if self.current_list_index < 0 or not self.pending_indices:
            self.set_status("无进度可回退", AppConfig.COLOR_WARNING)
            return
        actual_index = self.pending_indices[self.current_list_index]
        if actual_index > 0:
            target_index = actual_index - 1
            try:
                self.df.iloc[target_index, 1] = None
                self.df.to_excel(self.excel_path, index=False, header=False)
                self.pending_indices = self.df[self.df.iloc[:, 1].isna()].index.tolist()
                if target_index in self.pending_indices:
                    self.current_list_index = self.pending_indices.index(target_index)
                else:
                    self.current_list_index = 0
                self.update_current_song_display()
                self.set_status(f"已回退至第 {target_index + 1} 行", AppConfig.COLOR_INFO)
            except Exception as e:
                PixelMessageBox.show_error(self, "回退失败", f"无法保存 Excel: {e}")
        else:
            self.set_status("已经是 Excel 的第一首了", AppConfig.COLOR_WARNING)

    def show_preview(self):
        self.terminate_current_tasks()
        if self.df is None:
            PixelMessageBox.show_warning(self, "警告", "未加载任何数据")
            return
        self.preview_dialog = FramelessDialog(None, "文档列表预览")
        self.preview_dialog.resize(640, 600)
        self.preview_dialog.setWindowModality(QtCore.Qt.ApplicationModal)
        content_layout = QtWidgets.QVBoxLayout()
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(6)
        table = QtWidgets.QTableWidget()
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["行号", "歌曲名称", "标记状态"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        table.horizontalHeader().setStretchLastSection(False)
        table.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)
        table.verticalHeader().setDefaultSectionSize(32)
        table.setShowGrid(False)
        table.setAlternatingRowColors(True)
        table.setStyleSheet(f"""
            QTableWidget {{ background-color: {AppConfig.COLOR_BG}; color: {AppConfig.COLOR_TEXT}; border: none; gridline-color: {AppConfig.COLOR_BORDER}; selection-background-color: {AppConfig.COLOR_ACCENT}; selection-color: {AppConfig.COLOR_BG}; font-family: '{AppConfig.PIXEL_FONTS[0]}'; }}
            QHeaderView::section {{ background-color: {AppConfig.COLOR_PANEL}; color: {AppConfig.COLOR_TEXT}; border: none; border-bottom: 2px solid {AppConfig.COLOR_BORDER}; padding: 4px; font-family: '{AppConfig.PIXEL_FONTS[0]}'; font-weight: bold; }}
            QScrollBar:vertical {{ background: {AppConfig.COLOR_PANEL}; width: 12px; margin: 0px; border-left: 1px solid {AppConfig.COLOR_BORDER}; }}
            QScrollBar::handle:vertical {{ background: {AppConfig.COLOR_ACCENT}; min-height: 30px; margin: 0px; border-radius: 0px; }}
            QScrollBar::handle:vertical:hover {{ background: {AppConfig.COLOR_ACCENT_ALT}; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: none; }}
        """)
        content_layout.addWidget(table)
        self.preview_dialog.set_layout(content_layout)
        current_actual_index = -1
        if 0 <= self.current_list_index < len(self.pending_indices):
            current_actual_index = self.pending_indices[self.current_list_index]
        table.setRowCount(len(self.df))
        current_item = None
        for idx, row in self.df.iterrows():
            song_name = str(row[0])
            status = str(row[1]) if pd.notna(row[1]) else ""
            row_items = [
                QtWidgets.QTableWidgetItem(str(idx + 1)),
                QtWidgets.QTableWidgetItem(song_name),
                QtWidgets.QTableWidgetItem(status)
            ]
            for col, item in enumerate(row_items):
                item.setTextAlignment(QtCore.Qt.AlignCenter if col != 1 else QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                table.setItem(idx, col, item)
            if idx == current_actual_index:
                for item in row_items:
                    item.setBackground(QtGui.QColor(AppConfig.COLOR_ACCENT_ALT))
                    item.setForeground(QtGui.QColor(AppConfig.COLOR_BORDER))
                current_item = row_items[0]
            elif pd.notna(row[1]):
                for item in row_items:
                    item.setForeground(QtGui.QColor(AppConfig.COLOR_MUTED))
        def on_double_click(row, col):
            song_item = table.item(row, 1)
            if song_item:
                QtWidgets.QApplication.clipboard().setText(song_item.text())
                self.set_status(f"已复制 {song_item.text()}", AppConfig.COLOR_SUCCESS)
        table.cellDoubleClicked.connect(on_double_click)
        if current_item:
            QtCore.QTimer.singleShot(100, lambda: table.scrollToItem(current_item, QtWidgets.QAbstractItemView.PositionAtCenter))
        self.preview_dialog.show()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    app = QtWidgets.QApplication(sys.argv)
    window = MusicSearcherApp()
    window.show()
    sys.exit(app.exec())
