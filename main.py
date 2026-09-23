# -*- coding: utf-8 -*-
"""
任务栏行情
在 Windows 11 任务栏上显示一条薄薄的横向行情状态栏（默认嵌入任务栏空白区）。
数据源：东方财富（quote.eastmoney.com）。
每个标的竖排显示：涨跌幅（上）、名称（下）。

交互：
  - 悬停某标的 -> 高亮
  - 单击标的   -> 上方弹出分时图
  - 移出标的   -> 分时图关闭
  - 三连击     -> 退出程序

配置文件 config.json 每次刷新自动重读，无需重启。
"""

import ctypes
from ctypes import wintypes
import datetime
import gzip
import http.client
import json
import os
import queue
import re
import socket
import sys
import threading
import time
import winreg

import tkinter as tk
from tkinter import font as tkfont

# 窗口模式（pythonw / --windowed）下没有控制台，sys.stdout 为 None，
# print 会抛 AttributeError，这里统一重定向到空设备。
if getattr(sys, "frozen", False) and sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
    sys.stderr = sys.stdout


def _app_dir():
    """程序目录。打包成 exe 后指向 exe 所在目录，配置文件与其同级便于修改。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = _app_dir()
CONFIG_PATH = os.path.join(APP_DIR, "config.json")


def _ensure_config():
    """打包运行时：exe 同级若无 config.json，则释放内置默认配置一份。"""
    if not getattr(sys, "frozen", False) or os.path.exists(CONFIG_PATH):
        return
    src = os.path.join(getattr(sys, "_MEIPASS", APP_DIR), "config.json")
    try:
        if os.path.exists(src):
            with open(src, "r", encoding="utf-8-sig") as f:
                data = f.read()
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                f.write(data)
    except Exception as e:
        print("[配置] 初始化 config.json 失败:", e)


_ensure_config()

# 东方财富行情接口（多主机容错：实时主机优先，被限流时自动回退到延时主机）
EASTMONEY_QUOTE_HOSTS = ["push2.eastmoney.com", "push2delay.eastmoney.com"]
EASTMONEY_QUOTE_PATH = "/api/qt/stock/get"
EASTMONEY_TREND_HOSTS = ["push2his.eastmoney.com", "push2delay.eastmoney.com"]
EASTMONEY_TREND_PATH = "/api/qt/stock/trends2/get"
# 东方财富标准鉴权参数 ut；fltt=2 让服务端直接返回已按小数位折算好的浮点价格，
# 避免依赖 data.decimal（延时主机不返回该字段，缺省时会导致价格被放大百倍）。
EASTMONEY_UT = "fa5fd1943c7b386f172d6893dbfba10b"
# 实时字段：f43现价 f44最高 f45最低 f46开盘 f47成交量 f48成交额 f57代码 f58名称 f60昨收
EASTMONEY_QUOTE_FIELDS = "f43,f44,f45,f46,f47,f48,f57,f58,f60"
# 分时字段：f51时间 f52开盘 f53现价 f54最高 f55最低 f56成交量 f57成交额 f58均价
EASTMONEY_TREND_FIELDS1 = "f1,f2,f3,f4,f5,f6,f7,f8,f9,f10,f11,f12,f13"
EASTMONEY_TREND_FIELDS2 = "f51,f52,f53,f54,f55,f56,f57,f58"
# 手机版个股页：行情数据由服务端渲染在 HTML 的 quotedata 里，不经过 push 接口。
# 实测 push2 实时接口会随机断开连接（RemoteDisconnected），该页连通稳定且
# A 股价格为实时值，因此作为主行情源；push 接口降为备份源
# （wap 缺最高/最低/开盘价，当前界面未使用这些字段）。
WAP_QUOTE_HOST = "wap.eastmoney.com"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

DEFAULT_CONFIG = {
    "symbols": [
        {"code": "301308.SZ"},
        {"code": "300209.SZ"},
    ],
    "refresh_seconds": 5,
    # 显示模式：desktop = 桌面小组件（可拖动）；tray = 系统托盘图标 + 点击浮窗。
    # 两种模式可在程序内通过右键菜单切换。
    "mode": "desktop",
    "position": {"x": "right", "y": 40, "margin": 8, "gap_above_taskbar": 2,
                 "embed_taskbar": False},
    "statusbar": {"pad_x": 10, "pad_y": 5, "gap": 16, "show_avg": True,
                  "corner_radius": 8, "border_width": 1},
    "font": {"family": "Arial", "size": 9, "bold": False},
    "colors": {
        "background": "#1e1e1e",
        "name": "#aaaaaa",
        "price": "#eeeeee",
        "up": "#ff4d4f",
        "down": "#00c853",
        "flat": "#bbbbbb",
        "hover_bg": "#333333",
        "border": None,
        "auto_match_taskbar": True,
    },
    "timeline": {"width": 280, "height": 130, "bg": "#1e1e1e", "border": "#444444"},
    "layout": {"price_decimals": 2},
}


# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------

def _setup_console():
    """让控制台以 UTF-8 输出，避免中文乱码。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def _set_dpi_awareness():
    """启用系统 DPI 感知，使 Tk 与 ctypes 均使用物理像素，定位一致且文字清晰。"""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def _deep_merge(base, override):
    if not isinstance(base, dict) or not isinstance(override, dict):
        return override
    out = dict(base)
    for k, v in override.items():
        if k in out and isinstance(out[k], dict) and isinstance(v, dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def load_config():
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8-sig") as f:
                user = json.load(f)
            cfg = _deep_merge(cfg, user)
        except Exception as e:
            print("[配置] config.json 解析失败，使用默认配置:", e)
    return cfg


def _read_disk_config():
    """读取磁盘上的原始 config.json（未与默认值合并），失败返回 {}。"""
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8-sig") as f:
            return json.load(f)
    except Exception:
        return {}


def _write_disk_config(disk):
    """把字典写回 config.json（仅覆盖 mode/position 等用户可切换项，不污染其它）。"""
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(disk, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("[配置] 保存失败:", e)


def _system_dpi():
    """返回系统 DPI（每英寸像素数）。"""
    try:
        dpi = ctypes.windll.user32.GetDpiForSystem()
        if dpi:
            return int(dpi)
    except Exception:
        pass
    try:
        dc = ctypes.windll.user32.GetDC(0)
        dpi = ctypes.windll.gdi32.GetDeviceCaps(dc, 88)  # LOGPIXELSX
        ctypes.windll.user32.ReleaseDC(0, dc)
        if dpi:
            return int(dpi)
    except Exception:
        pass
    return 96


def _screen_metrics():
    """返回主屏幕物理尺寸 (宽, 高)。"""
    try:
        user32 = ctypes.windll.user32
        sw = user32.GetSystemMetrics(0)  # SM_CXSCREEN
        sh = user32.GetSystemMetrics(1)  # SM_CYSCREEN
        if sw and sh:
            return int(sw), int(sh)
    except Exception:
        pass
    return 1920, 1080


def _work_area():
    """返回主屏幕工作区（除去任务栏）物理坐标 (left, top, right, bottom)。"""
    class RECT(ctypes.Structure):
        _fields_ = [
            ("left", ctypes.c_long),
            ("top", ctypes.c_long),
            ("right", ctypes.c_long),
            ("bottom", ctypes.c_long),
        ]

    try:
        rect = RECT()
        ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0)
        return int(rect.left), int(rect.top), int(rect.right), int(rect.bottom)
    except Exception:
        sw, sh = _screen_metrics()
        return 0, 0, sw, sh


def _find_taskbar():
    """返回任务栏 Shell_TrayWnd 的窗口句柄，找不到返回 None。"""
    try:
        return ctypes.windll.user32.FindWindowW("Shell_TrayWnd", None)
    except Exception:
        return None


# --- AppBar 嵌入式任务栏（与腾讯电脑管家/Win11 天气相同的官方方案） ---------
# SHAppBarMessage 消息
_ABM_NEW = 0x00000000
_ABM_REMOVE = 0x00000001
_ABM_QUERYPOS = 0x00000002
_ABM_SETPOS = 0x00000003
_ABM_SETSTATE = 0x0000000A
# 边缘
_ABE_TOP = 1
_ABE_BOTTOM = 3
# AppBarState
_ABS_AUTOHIDE = 0x01
_ABS_ALWAYSONTOP = 0x02


class _APPBARDATA(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("hWnd", wintypes.HWND),
        ("uCallbackMessage", wintypes.UINT),
        ("uEdge", wintypes.UINT),
        ("rc", wintypes.RECT),
        ("lParam", wintypes.LPARAM),
    ]


class _TaskbarEmbedder:
    """用 Shell AppBar API 把一个窗口注册为"桌面工具栏"。

    注册成功后，Explorer 会把窗口作为任务栏的子区域安排位置（与腾讯电脑
    管家、Win11 天气小卡相同机制）。窗口不会被普通的桌面 z-order 遮挡，
    全屏最大化时也不会盖住。
    """

    def __init__(self):
        self._registered = False
        self._hwnd = None

    def attach(self, hwnd):
        if self._registered:
            return
        self._hwnd = hwnd
        try:
            shell32 = ctypes.windll.shell32
            abd = _APPBARDATA()
            abd.cbSize = ctypes.sizeof(_APPBARDATA)
            abd.hWnd = hwnd
            abd.uEdge = _ABE_BOTTOM  # 视觉锚定到底部
            # 必须先注册，再设置位置（ABM_SETPOS 时系统会回调 ABN_POSCHANGED）
            shell32.SHAppBarMessage(_ABM_NEW, ctypes.byref(abd))
            self._registered = True
            # 取消自动隐藏，确保任务栏不会被认定为自动隐藏而收缩
            abd.lParam = _ABS_ALWAYSONTOP
            shell32.SHAppBarMessage(_ABM_SETSTATE, ctypes.byref(abd))
        except Exception as e:
            print("[AppBar] 注册失败:", e)
            self._registered = False

    def detach(self):
        if not self._registered or not self._hwnd:
            return
        try:
            abd = _APPBARDATA()
            abd.cbSize = ctypes.sizeof(_APPBARDATA)
            abd.hWnd = self._hwnd
            ctypes.windll.shell32.SHAppBarMessage(_ABM_REMOVE, ctypes.byref(abd))
        except Exception:
            pass
        self._registered = False

    def _make_wndproc(self):
        # 留作以后处理 ABN_POSCHANGED 的钩子，目前 AppBar 模式不强制要求。
        return None


# --- 系统托盘图标（纯 GDI 绘制，不依赖 PIL） ---------------------------------

# 托盘窗口过程回调类型（模块级定义，供嵌套结构体字段引用）
_TRAY_WNDPROC = ctypes.WINFUNCTYPE(
    ctypes.c_longlong, wintypes.HWND, wintypes.UINT,
    wintypes.WPARAM, wintypes.LPARAM)


class _TrayIcon:
    """把程序挂到系统托盘：显示涨跌 K 线图标，悬停显示行情摘要。

    使用独立 Win32 隐藏窗口 + 独立线程消息泵承载托盘图标，避免与 Tk
    主循环冲突；托盘鼠标事件通过 on_event 回调抛给主线程处理。
    """

    WM_APP = 0x8000
    WM_TRAYCB = WM_APP + 1
    NIM_ADD = 0x00000000
    NIM_MODIFY = 0x00000001
    NIM_DELETE = 0x00000002
    NIF_MESSAGE = 0x00000001
    NIF_ICON = 0x00000002
    NIF_TIP = 0x00000004
    WM_LBUTTONUP = 0x0202
    WM_LBUTTONDBLCLK = 0x0203
    WM_RBUTTONUP = 0x0205

    class _NOTIFYICONDATAW(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD),
            ("hWnd", wintypes.HWND),
            ("uID", wintypes.UINT),
            ("uFlags", wintypes.UINT),
            ("uCallbackMessage", wintypes.UINT),
            ("hIcon", wintypes.HICON),
            ("szTip", wintypes.WCHAR * 128),
            ("dwState", wintypes.DWORD),
            ("dwStateMask", wintypes.DWORD),
            ("szInfo", wintypes.WCHAR * 256),
            ("uTimeoutOrVersion", wintypes.UINT),
            ("szInfoTitle", wintypes.WCHAR * 64),
            ("dwInfoFlags", wintypes.DWORD),
            ("guidItem", ctypes.c_byte * 16),
            ("hBalloonIcon", wintypes.HICON),
        ]

    class _WNDCLASSW(ctypes.Structure):
        _fields_ = [
            ("style", wintypes.UINT),
            ("lpfnWndProc", _TRAY_WNDPROC),
            ("cbClsExtra", ctypes.c_int),
            ("cbWndExtra", ctypes.c_int),
            ("hInstance", wintypes.HINSTANCE),
            ("hIcon", wintypes.HICON),
            ("hCursor", wintypes.HANDLE),
            ("hbrBackground", wintypes.HBRUSH),
            ("lpszMenuName", wintypes.LPCWSTR),
            ("lpszClassName", wintypes.LPCWSTR),
        ]

    class _MSG(ctypes.Structure):
        _fields_ = [
            ("hwnd", wintypes.HWND),
            ("message", wintypes.UINT),
            ("wParam", wintypes.WPARAM),
            ("lParam", wintypes.LPARAM),
            ("time", wintypes.DWORD),
            ("pt", wintypes.POINT),
        ]

    def __init__(self, on_event):
        self.on_event = on_event
        self._hwnd = None
        self._hicon = None
        self._hbm = None
        self._hmask = None
        self._alive = False
        self._user32 = ctypes.windll.user32
        self._shell32 = ctypes.windll.shell32
        self._kernel32 = ctypes.windll.kernel32
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def _run(self):
        try:
            self._kernel32.GetModuleHandleW.restype = wintypes.HINSTANCE
            self._user32.CreateWindowExW.restype = wintypes.HWND
            self._user32.CreateWindowExW.argtypes = [
                wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR,
                wintypes.DWORD, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                ctypes.c_int, wintypes.HWND, wintypes.HMENU,
                wintypes.HINSTANCE, wintypes.LPVOID,
            ]
            self._user32.DefWindowProcW.restype = ctypes.c_longlong
            self._user32.DefWindowProcW.argtypes = [
                wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
            self._hicon, self._hbm = self._make_icon()
            self._create_window()
            self._add_icon("行情")
            self._alive = True
            msg = self._MSG()
            PM_REMOVE = 0x0001
            while self._alive:
                if self._user32.PeekMessageW(
                        ctypes.byref(msg), None, 0, 0, PM_REMOVE):
                    self._user32.TranslateMessage(ctypes.byref(msg))
                    self._user32.DispatchMessageW(ctypes.byref(msg))
                else:
                    time.sleep(0.02)
        except Exception as e:
            print("[托盘] 初始化失败:", e)

    def _create_window(self):
        cls_name = "BarQuoteTrayWnd"
        wc = self._WNDCLASSW()
        wc.lpfnWndProc = _TRAY_WNDPROC(self._wnd_proc)
        self._wndproc_ref = wc.lpfnWndProc  # 防 GC
        wc.hInstance = self._kernel32.GetModuleHandleW(None)
        wc.lpszClassName = cls_name
        self._user32.RegisterClassW(ctypes.byref(wc))
        self._hwnd = self._user32.CreateWindowExW(
            0, cls_name, "", 0, 0, 0, 0, 0, None, None, wc.hInstance, None)
        if not self._hwnd:
            raise RuntimeError("创建托盘隐藏窗口失败")

    def _wnd_proc(self, hwnd, msg, wparam, lparam):
        if msg == self.WM_TRAYCB:
            if lparam in (self.WM_LBUTTONUP, self.WM_LBUTTONDBLCLK,
                          self.WM_RBUTTONUP):
                try:
                    self.on_event(int(lparam))
                except Exception:
                    pass
            return 0
        return self._user32.DefWindowProcW(hwnd, msg, wparam, lparam)

    def _nid(self):
        nid = self._NOTIFYICONDATAW()
        nid.cbSize = ctypes.sizeof(self._NOTIFYICONDATAW)
        nid.hWnd = self._hwnd
        nid.uID = 1
        return nid

    def _add_icon(self, tip):
        nid = self._nid()
        nid.uFlags = self.NIF_MESSAGE | self.NIF_ICON | self.NIF_TIP
        nid.uCallbackMessage = self.WM_TRAYCB
        nid.hIcon = self._hicon
        nid.szTip = tip
        self._shell32.Shell_NotifyIconW(self.NIM_ADD, ctypes.byref(nid))

    def set_tooltip(self, text):
        """更新悬停提示（可从任意线程调用，内部有句柄判空）。"""
        if not self._hwnd or not self._alive:
            return
        text = (text or "")[:127]
        try:
            nid = self._nid()
            nid.uFlags = self.NIF_TIP
            nid.szTip = text
            self._shell32.Shell_NotifyIconW(self.NIM_MODIFY, ctypes.byref(nid))
        except Exception:
            pass

    def stop(self):
        self._alive = False
        if self._hwnd:
            try:
                nid = self._nid()
                self._shell32.Shell_NotifyIconW(self.NIM_DELETE, ctypes.byref(nid))
            except Exception:
                pass
            try:
                self._user32.DestroyWindow(self._hwnd)
            except Exception:
                pass
            self._hwnd = None
        if self._hicon:
            try:
                self._user32.DestroyIcon(self._hicon)
            except Exception:
                pass
            self._hicon = None
        if self._hbm:
            try:
                ctypes.windll.gdi32.DeleteObject(self._hbm)
            except Exception:
                pass
            self._hbm = None
        if self._hmask:
            try:
                ctypes.windll.gdi32.DeleteObject(self._hmask)
            except Exception:
                pass
            self._hmask = None

    # ---- 图标绘制（32x32 ARGB，圆底 + 红绿双 K 线） -----------------------
    def _make_icon(self):
        size = 32
        gdi32 = ctypes.windll.gdi32
        user32 = self._user32
        gdi32.CreateDIBSection.restype = wintypes.HBITMAP
        user32.CreateIconIndirect.restype = wintypes.HICON

        class BIH(ctypes.Structure):
            _fields_ = [
                ("biSize", wintypes.DWORD),
                ("biWidth", ctypes.c_long),
                ("biHeight", ctypes.c_long),
                ("biPlanes", wintypes.WORD),
                ("biBitCount", wintypes.WORD),
                ("biCompression", wintypes.DWORD),
                ("biSizeImage", wintypes.DWORD),
                ("biXPelsPerMeter", ctypes.c_long),
                ("biYPelsPerMeter", ctypes.c_long),
                ("biClrUsed", wintypes.DWORD),
                ("biClrImportant", wintypes.DWORD),
            ]

        class BI(ctypes.Structure):
            _fields_ = [("bmiHeader", BIH), ("bmiColors", wintypes.DWORD * 3)]

        bmi = BI()
        bmi.bmiHeader.biSize = ctypes.sizeof(BIH)
        bmi.bmiHeader.biWidth = size
        bmi.bmiHeader.biHeight = size  # bottom-up（图标位图要求自下而上）
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = 0  # BI_RGB

        pixels = ctypes.c_void_p()
        hbm = gdi32.CreateDIBSection(
            None, ctypes.byref(bmi), 0, ctypes.byref(pixels), None, 0)
        if not hbm:
            raise RuntimeError("创建图标位图失败")
        buf = (ctypes.c_ubyte * (size * size * 4)).from_address(pixels.value)

        def setp(x, y, r, g, b, a=255):
            if 0 <= x < size and 0 <= y < size:
                i = ((size - 1 - y) * size + x) * 4  # bottom-up 翻转
                buf[i] = b
                buf[i + 1] = g
                buf[i + 2] = r
                buf[i + 3] = a

        cx = cy = size / 2.0
        rad = size / 2.0 - 1.0
        for y in range(size):
            for x in range(size):
                dx = x + 0.5 - cx
                dy = y + 0.5 - cy
                if dx * dx + dy * dy <= rad * rad:
                    setp(x, y, 36, 36, 36, 255)  # 深色圆底
                else:
                    setp(x, y, 0, 0, 0, 0)

        # 红 K（左，上涨）：实体 y=12..20，影线 y=8..24
        for y in range(8, 25):
            setp(11, y, 255, 77, 79, 255)
        for y in range(12, 21):
            for x in (10, 11, 12):
                setp(x, y, 255, 77, 79, 255)
        # 绿 K（右，下跌）：实体 y=8..16，影线 y=6..26
        for y in range(6, 27):
            setp(20, y, 0, 200, 83, 255)
        for y in range(8, 17):
            for x in (19, 20, 21):
                setp(x, y, 0, 200, 83, 255)

        class ICONINFO(ctypes.Structure):
            _fields_ = [
                ("fIcon", wintypes.BOOL),
                ("xHotspot", wintypes.DWORD),
                ("yHotspot", wintypes.DWORD),
                ("hbmMask", wintypes.HBITMAP),
                ("hbmColor", wintypes.HBITMAP),
            ]

        # 单色 mask（全 0 = 不透明，透明信息由 32 位 alpha 通道提供）
        mask_row = (size + 15) // 16 * 2  # 每行按 16 位对齐
        mask_buf = (ctypes.c_ubyte * (mask_row * size))()
        gdi32.CreateBitmap.restype = wintypes.HBITMAP
        hmask = gdi32.CreateBitmap(size, size, 1, 1, ctypes.byref(mask_buf))

        info = ICONINFO()
        info.fIcon = True
        info.hbmMask = hmask
        info.hbmColor = hbm
        hicon = user32.CreateIconIndirect(ctypes.byref(info))
        if not hicon:
            gdi32.DeleteObject(hbm)
            if hmask:
                gdi32.DeleteObject(hmask)
            raise RuntimeError("创建图标失败")
        self._hmask = hmask
        return hicon, hbm


def _taskbar_rect():
    """返回任务栏物理坐标 (left, top, right, bottom)，找不到返回 None。"""
    tb = _find_taskbar()
    if not tb:
        return None
    r = wintypes.RECT()
    try:
        ctypes.windll.user32.GetWindowRect(tb, ctypes.byref(r))
    except Exception:
        return None
    return (int(r.left), int(r.top), int(r.right), int(r.bottom))


def _tray_notify_left():
    """返回系统托盘区相对任务栏左上角的左边界 x（像素），找不到返回 None。"""
    tb = _find_taskbar()
    if not tb:
        return None
    try:
        user32 = ctypes.windll.user32
        notify = user32.FindWindowExW(tb, None, "TrayNotifyWnd", None)
        if not notify:
            return None
        nr = wintypes.RECT()
        tr = wintypes.RECT()
        user32.GetWindowRect(notify, ctypes.byref(nr))
        user32.GetWindowRect(tb, ctypes.byref(tr))
        return int(nr.left - tr.left)
    except Exception:
        return None


def _taskbar_light_theme():
    """返回任务栏是否为浅色主题（True=浅色，False=深色，None=未知）。"""
    try:
        with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as k:
            val, _ = winreg.QueryValueEx(k, "SystemUsesLightTheme")
            return bool(val)
    except Exception:
        return None


def _to_native_code(code):
    """把用户输入的代码转换成东方财富 secid（市场前缀 + 代码）。

    东方财富 secid 规则：0. 开头=深市，1. 开头=沪市，0./1. 也覆盖指数。
    港股为 116. 前缀（如 116.00700），美股无统一前缀、暂不支持。
    """
    code = (code or "").strip()
    if not code:
        return None
    # 已是东方财富 secid（形如 0.301308 / 1.600000 / 116.00700）
    if re.match(r"^\d+\.\d+$", code):
        return code
    m = re.match(r"^([0-9A-Za-z]+)\.([A-Za-z]{2})$", code)
    if not m:
        # 兼容无后缀纯数字/字母（默认当作 A 股，按代码首位判断市场）
        sym, market = code.upper(), ""
    else:
        sym, market = m.group(1).upper(), m.group(2).upper()

    if market in ("SH", "SZ", "BJ"):
        if market == "SH":
            return "1." + sym
        if market == "SZ":
            return "0." + sym
        return "0." + sym  # 北交所暂归深市前缀
    if market == "HK":
        digits = re.sub(r"\D", "", sym)
        if len(digits) < 5:
            digits = digits.zfill(5)
        elif len(digits) > 5:
            digits = digits[-5:]
        return "116." + digits
    if market == "US":
        return sym  # 美股暂不支持，原样返回以便提示
    # 无后缀：按 A 股代码首位判断市场
    sym = re.sub(r"\D", "", sym)
    if not sym:
        return None
    if sym.startswith(("5", "6", "9")):
        return "1." + sym
    return "0." + sym


# ---------------------------------------------------------------------------
# 行情查询窗口
# ---------------------------------------------------------------------------
# 周一至周五 9:15-16:00：持续查询行情接口并打印日志。
# 其余时间（含周末）：不查询、不打印，仅在任务启动时查询一次。
# 注：仅按周一至周五 + 固定时段判断，未覆盖法定节假日。

_QUERY_WINDOW_START = 9 * 60 + 15   # 09:15
_QUERY_WINDOW_END = 16 * 60         # 16:00


def _in_query_window(now=None):
    """是否处于行情查询窗口内（周一至周五 9:15-16:00）。"""
    now = now or datetime.datetime.now()
    if now.weekday() >= 5:  # 周六、周日
        return False
    hm = now.hour * 60 + now.minute
    return _QUERY_WINDOW_START <= hm < _QUERY_WINDOW_END


def _http_get_bytes(host, path, timeout=8):
    """用 http.client 请求东方财富并返回原始字节（自动解 gzip）。"""
    headers = {
        "User-Agent": USER_AGENT,
        "Referer": "https://quote.eastmoney.com/",
        "Accept": "*/*",
        "Accept-Encoding": "gzip, deflate",
    }
    conn = http.client.HTTPSConnection(host, timeout=timeout)
    try:
        conn.request("GET", path, headers=headers)
        resp = conn.getresponse()
        raw = resp.read()
        if raw[:2] == b"\x1f\x8b":  # gzip 魔数
            try:
                raw = gzip.decompress(raw)
            except Exception:
                pass
        return raw
    finally:
        try:
            conn.close()
        except Exception:
            pass


def _http_get_json(hosts, path, timeout=8, attempts_per_host=2):
    """请求东方财富接口并返回解析后的 JSON（多主机容错 + 重试）。

    东方财富 push 接口对高频/无 cookie 访问会做软限流，间歇性断开连接
    （RemoteDisconnected）。因此：
      - 每个主机重试 attempts_per_host 次，失败后短暂退避再试；
      - 多主机依次切换（实时主机被限流时自动回退到延时主机）。
    """
    if isinstance(hosts, str):
        hosts = [hosts]
    last_err = None
    for host in hosts:
        for attempt in range(attempts_per_host):
            try:
                raw = _http_get_bytes(host, path, timeout=timeout)
                return json.loads(raw.decode("utf-8", "replace"))
            except Exception as e:
                last_err = e
                if attempt < attempts_per_host - 1:
                    time.sleep(0.3 * (attempt + 1))
    raise last_err or RuntimeError("请求失败")


def _to_float(v):
    if v is None or v == "" or v == "-":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------------------
# 数据抓取
# ---------------------------------------------------------------------------

def _fetch_quote_push(native):
    """通过 push 接口抓取实时行情（含最高/最低/昨收）。

    请求带 fltt=2，价格字段直接返回已折算好的浮点数，无需再按 decimal 缩放。
    """
    path = "%s?secid=%s&fields=%s&ut=%s&fltt=2&invt=2" % (
        EASTMONEY_QUOTE_PATH, native, EASTMONEY_QUOTE_FIELDS, EASTMONEY_UT)
    obj = _http_get_json(EASTMONEY_QUOTE_HOSTS, path)
    data = (obj or {}).get("data")
    if not data:
        raise ValueError("无法解析返回数据")

    price = _to_float(data.get("f43"))    # 现价
    pre = _to_float(data.get("f60"))      # 昨收
    high = _to_float(data.get("f44"))     # 最高
    low = _to_float(data.get("f45"))      # 最低
    name = (data.get("f58") or "").strip()

    pct = None
    if price is not None and pre:
        pct = (price - pre) / pre * 100.0

    return {
        "native": native,
        "name": name,
        "price": price,
        "pre": pre,
        "high": high,
        "low": low,
        "avg": None,   # 均价由分时数据提供，实时接口单独取不可靠
        "pct": pct,
    }


def _fetch_quote_wap(native):
    """通过手机版个股页抓取行情（服务端渲染，连通稳定；无最高/最低/开盘）。

    wap.eastmoney.com 的个股页把行情内嵌在 HTML 的 `var quotedata = {...}` 里，
    不经过 push 接口。实测 push2 实时接口会随机断开连接（成功率约 3/8），
    而该页连通率 100% 且 A 股价格为实时值，故作为主行情源。
    """
    path = "/quote/stock/%s.html?appfenxiang=1" % native
    raw = None
    last_err = None
    for attempt in range(2):
        try:
            raw = _http_get_bytes(WAP_QUOTE_HOST, path)
            break
        except Exception as e:
            last_err = e
            if attempt == 0:
                time.sleep(0.3)
    if raw is None:
        raise last_err or RuntimeError("wap 请求失败")
    m = re.search(r"var quotedata = (\{.*?\});", raw.decode("utf-8", "replace"))
    if not m:
        raise ValueError("网页中未找到 quotedata")
    d = json.loads(m.group(1))

    try:
        dec = int(d.get("decimal59") or 2)
    except (TypeError, ValueError):
        dec = 2
    scale = 10 ** dec

    price = _to_float(d.get("price"))
    if price is not None:
        price = price / scale
    zde = _to_float(d.get("zde"))          # 涨跌额
    if zde is not None:
        zde = zde / scale
    pre = (price - zde) if price is not None and zde is not None else None
    zdf = _to_float(d.get("zdf"))          # 涨跌幅（已 *100 的整数，如 -603 表示 -6.03%）
    pct = (zdf / 100.0) if zdf is not None else None
    name = (d.get("name") or "").strip()

    return {
        "native": native,
        "name": name,
        "price": price,
        "pre": pre,
        "high": None,
        "low": None,
        "avg": None,
        "pct": pct,
    }


def fetch_quote(native):
    """抓取实时行情：优先手机版网页（连通稳定），失败时回退 push 接口。

    push 接口数据更全（含最高/最低），但当前会随机断连，故仅作备份。
    """
    try:
        return _fetch_quote_wap(native)
    except Exception as e:
        try:
            return _fetch_quote_push(native)
        except Exception as e2:
            raise RuntimeError("wap 与 push 均获取失败: %s / %s" % (e, e2))


def fetch_timeline(native):
    """抓取当日分时数据，返回 (昨收, [(时间, 现价, 均价), ...])。"""
    path = "%s?secid=%s&fields1=%s&fields2=%s&ut=%s&fltt=2&ndays=1&iscr=0&iscca=0" % (
        EASTMONEY_TREND_PATH, native,
        EASTMONEY_TREND_FIELDS1, EASTMONEY_TREND_FIELDS2, EASTMONEY_UT)
    obj = _http_get_json(EASTMONEY_TREND_HOSTS, path)
    data = (obj or {}).get("data")
    if not data:
        raise ValueError("无法解析分时数据")
    pre = _to_float(data.get("preClose"))
    rows = []
    for seg in (data.get("trends") or []):
        parts = str(seg).split(",")
        if len(parts) >= 8:
            # 时间, 开盘, 现价, 最高, 最低, 成交量, 成交额, 均价
            price = _to_float(parts[2])
            avg = _to_float(parts[7])
            if price is not None:
                rows.append((parts[0], price, avg))
    return pre, rows


# ---------------------------------------------------------------------------
# 单个标的控件
# ---------------------------------------------------------------------------

def _short_name(name, fallback=""):
    """显示名称：取行情接口返回名称的前两个字，取不到时回退到代码。"""
    name = (name or "").strip()
    if not name:
        return (fallback or "").strip()
    return name[:2]


def _is_convertible_bond(native):
    """是否可转债（东方财富 secid 形式）：沪市 11xxxx（1. 前缀）、深市 12xxxx（0. 前缀）。"""
    market, _, sym = (native or "").partition(".")
    if not sym:
        return False
    if market == "1":
        return sym.startswith("11")
    if market == "0":
        return sym.startswith("12")
    return False


class ItemWidget:
    def __init__(self, parent, native, code, font, bg):
        self.native = native
        self.code = code
        self.frame = tk.Frame(parent, bg=bg, highlightthickness=0)
        self.lbl_pct = tk.Label(self.frame, text="--", font=font, fg=bg,
                                bg=bg, bd=0, anchor="center", justify="center")
        self.lbl_name = tk.Label(self.frame, text="--", font=font, fg=bg,
                                 bg=bg, bd=0, anchor="center", justify="center")
        self.lbl_pct.pack(fill="x")
        self.lbl_name.pack(fill="x")
        self.widgets = (self.frame, self.lbl_pct, self.lbl_name)


# ---------------------------------------------------------------------------
# 主程序
# ---------------------------------------------------------------------------

class QuoteBar:
    def __init__(self):
        self.stop_event = threading.Event()
        self.ui_queue = queue.Queue()
        self.quotes = {}
        self.items = []
        self.symbol_key = None
        self.current_cfg = load_config()
        self.hovered = None
        self.click_times = []
        self._closing = False

        self.popup_win = None
        self.popup_canvas = None
        self.popup_req = 0
        self.embedded = False
        self.embedded_hwnd = None
        self.appbar = _TaskbarEmbedder()

        # 显示模式与托盘/拖动状态
        self.mode = None                 # 当前生效模式：desktop / tray
        self.tray = None                 # _TrayIcon 实例（tray 模式）
        self.float_visible = False       # tray 模式下浮窗是否可见
        self._drag_press = None          # 拖动起点 (rx, ry, wx, wy)
        self._dragged = False

        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.configure(bg=self.current_cfg["colors"]["background"])
        self.root.title("任务栏行情")

        self.dpi = _system_dpi()
        # Tk 字号（点）固定按 96dpi 渲染，这里按系统 DPI 换算点数保证视觉一致
        self.font = tkfont.Font(family="Arial",
                                size=-max(1, round(9 * self.dpi / 96.0)),
                                weight="normal")

        self.container = tk.Frame(self.root, bg=self.current_cfg["colors"]["background"])
        self.container.pack(fill="both", expand=True)
        self.inner = tk.Frame(self.container, bg=self.current_cfg["colors"]["background"])
        self.inner.pack(fill="both", expand=True)

        self.root.bind("<Motion>", self._on_motion)
        self.root.bind("<Leave>", lambda e: self._set_hover(None))
        # 拖动 + 点击：按下记录、移动拖动、释放判定是否点击
        self.root.bind("<ButtonPress-1>", self._on_press)
        self.root.bind("<B1-Motion>", self._on_drag)
        self.root.bind("<ButtonRelease-1>", self._on_release)
        self.root.bind("<Button-3>", self._on_right_click)

    # ---- 主循环 ----------------------------------------------------------

    def run(self):
        self._apply_config_style(self.current_cfg)
        self._build_items(self.current_cfg)
        self._paint_all()
        # 初始化显示模式（桌面小组件 / 系统托盘），内部完成定位或隐藏
        self._apply_mode(self.current_cfg)
        self.root.after(0, self._apply_window_style)
        self.root.after(1000, self._keep_on_top)
        self.root.after(200, self._poll_queue)

        threading.Thread(target=self._worker, daemon=True).start()
        self.root.mainloop()

    def _worker(self):
        last_quotes = {}  # 数据缓存：查询窗口外沿用最后一次结果
        first_round = True  # 任务启动时无条件查询一次
        while not self.stop_event.is_set():
            try:
                cfg = load_config()
            except Exception as e:
                print("[配置] 读取失败:", e)
                self.stop_event.wait(5)
                continue

            in_window = _in_query_window()
            # 窗口内持续查询；窗口外仅任务启动时查询一次
            do_fetch = in_window or first_round

            if do_fetch:
                codes = []
                for s in cfg.get("symbols", []):
                    codes.append(_to_native_code(s.get("code", "")))

                quotes = {}
                for c in codes:
                    if not c:
                        continue
                    try:
                        quotes[c] = fetch_quote(c)
                    except Exception as e:
                        print("[行情] %s 获取失败: %s" % (c, e))
                        prev = last_quotes.get(c)
                        if prev is not None:
                            quotes[c] = prev

                last_quotes = dict(quotes)

                # 控制台日志
                log = []
                for s, c in zip(cfg.get("symbols", []), codes):
                    q = quotes.get(c) if c else None
                    label = _short_name(q.get("name") if q else "",
                                        s.get("code", "")) or "?"
                    if q and q.get("price") is not None:
                        log.append("%s %.2f %+.2f%%" % (
                            label, q["price"], q.get("pct") or 0.0))
                    else:
                        log.append("%s 无数据" % label)
                tag = "" if in_window else "  [窗口外·启动时查询]"
                print("[%s] %s%s" % (
                    time.strftime("%H:%M:%S"), "  ".join(log), tag))
            else:
                # 窗口外：不查询、不打印日志，沿用缓存数据
                quotes = dict(last_quotes)

            self.ui_queue.put(("update", cfg, quotes))
            try:
                interval = int(cfg.get("refresh_seconds", 5) or 5)
            except (TypeError, ValueError):
                interval = 5
            if not in_window:
                interval = max(interval, 60)  # 窗口外低频等待，直到进入窗口
            self.stop_event.wait(max(1, interval))
            first_round = False

    def _poll_queue(self):
        try:
            while True:
                msg = self.ui_queue.get_nowait()
                kind = msg[0]
                if kind == "update":
                    self._apply_update(msg[1], msg[2])
                elif kind == "timeline":
                    self._draw_timeline(*msg[1:])
                elif kind == "tray":
                    self._handle_tray_event(msg[1])
        except queue.Empty:
            pass
        if not self._closing:
            self.root.after(200, self._poll_queue)

    # ---- 配置应用与构建 --------------------------------------------------

    def _apply_config_style(self, cfg):
        col = cfg["colors"]
        if bool(col.get("auto_match_taskbar", True)):
            light = _taskbar_light_theme()
            if light is not None:
                self._apply_theme_colors(col, light)

        f = cfg.get("font", {})
        try:
            size = int(f.get("size", 9))
        except (TypeError, ValueError):
            size = 9
        family = f.get("family") or "Arial"
        bold = bool(f.get("bold", False))
        # 字号按“点”理解；Tk 的点固定按 96dpi 渲染，故按系统 DPI 换算点数
        size_pt = max(1, round(size * self.dpi / 96.0))
        self.font.configure(family=family, size=-size_pt,
                            weight="bold" if bold else "normal")
        border = col.get("border") or col["background"]
        self.container.configure(bg=col["background"])
        self.inner.configure(bg=col["background"])
        self.root.configure(bg=border)

    @staticmethod
    def _apply_theme_colors(col, light):
        """按任务栏深浅色覆盖背景/文字色（涨跌色保持不变）。"""
        if light:
            col["background"] = "#f3f3f3"
            col["name"] = "#5f5f5f"
            col["price"] = "#111111"
            col["hover_bg"] = "#e4e4e4"
            col["flat"] = "#8a8a8a"
            col["border"] = "#c8c8c8"
        else:
            col["background"] = "#242424"
            col["name"] = "#9a9a9a"
            col["price"] = "#eeeeee"
            col["hover_bg"] = "#3a3a3a"
            col["flat"] = "#bbbbbb"
            col["border"] = "#5a5a5a"

    def _build_items(self, cfg):
        for it in self.items:
            it.frame.destroy()
        self.items = []
        col = cfg["colors"]
        bg = col["background"]
        syms = cfg.get("symbols", [])
        for s in syms:
            code = (s.get("code") or "").strip()
            native = _to_native_code(code)
            it = ItemWidget(self.inner, native, code, self.font, bg)
            self.items.append(it)

    def _symbol_key(self, cfg):
        return tuple(
            (s.get("code") or "").strip()
            for s in cfg.get("symbols", [])
        )

    def _apply_update(self, cfg, quotes):
        self.current_cfg = cfg
        self.quotes.update(quotes)
        self._apply_config_style(cfg)
        key = self._symbol_key(cfg)
        if key != self.symbol_key:
            self.symbol_key = key
            self._build_items(cfg)
        self._paint_all()
        self._apply_mode(cfg)
        self._update_tray_tooltip()

    # ---- 显示模式（桌面小组件 / 系统托盘）切换 ---------------------------

    def _apply_mode(self, cfg):
        """按 cfg["mode"] 应用显示模式；仅在模式变化时做重切换，其余刷新只保位。"""
        mode = (cfg.get("mode") or "desktop").strip().lower()
        if mode not in ("desktop", "tray"):
            mode = "desktop"
        if mode != self.mode:
            if self.mode == "tray":
                self._leave_tray()
            self.mode = mode
            if mode == "tray":
                self._enter_tray()
            else:
                self._enter_desktop()
        # 刷新时保持几何位置（拖动过程中跳过，避免把窗口拽回原位）
        if self._drag_press is not None:
            return
        if self.mode == "desktop":
            self._apply_geometry(cfg)
        elif self.mode == "tray" and self.float_visible:
            self._apply_geometry(cfg, tray_float=True)

    def _enter_tray(self):
        self._close_popup()
        self.root.withdraw()
        self.float_visible = False
        if self.tray is None:
            self.tray = _TrayIcon(self._on_tray_event)

    def _leave_tray(self):
        if self.tray is not None:
            try:
                self.tray.stop()
            except Exception:
                pass
            self.tray = None
        self.float_visible = False

    def _enter_desktop(self):
        self.root.deiconify()
        self.root.attributes("-topmost", True)
        self.root.lift()

    def _on_tray_event(self, code):
        # 托盘线程回调 → 转交主线程队列
        self.ui_queue.put(("tray", code))

    def _handle_tray_event(self, code):
        if code in (_TrayIcon.WM_LBUTTONUP, _TrayIcon.WM_LBUTTONDBLCLK):
            self._toggle_float()
        elif code == _TrayIcon.WM_RBUTTONUP:
            self._show_menu(None)

    def _toggle_float(self):
        if self.float_visible:
            self.root.withdraw()
            self.float_visible = False
        else:
            self.root.deiconify()
            self.root.attributes("-topmost", True)
            self.root.lift()
            self._apply_geometry(self.current_cfg, tray_float=True)
            self.float_visible = True

    def _update_tray_tooltip(self):
        if self.tray is None:
            return
        parts = []
        for it in self.items:
            q = self.quotes.get(it.native) if it.native else None
            if not q or q.get("price") is None:
                continue
            name = _short_name(q.get("name"), it.code)
            pct = q.get("pct")
            pct_s = ("%+.2f%%" % pct) if pct is not None else "--"
            parts.append("%s %.2f %s" % (name, q["price"], pct_s))
        self.tray.set_tooltip("  ".join(parts) if parts else "行情加载中…")

    # ---- 拖动与点击 -------------------------------------------------------

    def _on_press(self, event):
        if self.mode != "desktop":
            return
        self._drag_press = (event.x_root, event.y_root,
                            self.root.winfo_x(), self.root.winfo_y())
        self._dragged = False

    def _on_drag(self, event):
        if self._drag_press is None:
            return
        rx0, ry0, wx0, wy0 = self._drag_press
        dx = event.x_root - rx0
        dy = event.y_root - ry0
        if not self._dragged and (abs(dx) + abs(dy)) > 5:
            self._dragged = True
        if self._dragged:
            self.root.geometry("+%d+%d" % (wx0 + dx, wy0 + dy))

    def _on_release(self, event):
        if self._drag_press is None:
            return
        was_drag = self._dragged
        self._drag_press = None
        self._dragged = False
        if was_drag:
            self._save_position()
            return
        # 未拖动 = 单击：走原有点击逻辑（显示分时图 / 三连击退出）
        self._handle_click(event)

    def _save_position(self):
        x = self.root.winfo_x()
        y = self.root.winfo_y()
        pos = self.current_cfg.setdefault("position", {})
        pos["x"] = x
        pos["y"] = y
        disk = _read_disk_config()
        disk.setdefault("position", {})["x"] = x
        disk.setdefault("position", {})["y"] = y
        _write_disk_config(disk)

    def _on_right_click(self, event):
        self._show_menu(event)

    def _show_menu(self, event):
        menu = tk.Menu(self.root, tearoff=0)
        if self.mode == "desktop":
            menu.add_command(label="切换到系统托盘图标", command=self._switch_to_tray)
        else:
            menu.add_command(label="切换到桌面小组件", command=self._switch_to_desktop)
        menu.add_separator()
        menu.add_command(label="退出程序", command=self.shutdown)
        if event is not None:
            x, y = event.x_root, event.y_root
        else:
            pt = wintypes.POINT()
            ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
            x, y = int(pt.x), int(pt.y)
        try:
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def _switch_to_tray(self):
        self._switch_mode("tray")

    def _switch_to_desktop(self):
        self._switch_mode("desktop")

    def _switch_mode(self, mode):
        disk = _read_disk_config()
        disk["mode"] = mode
        _write_disk_config(disk)
        self.current_cfg["mode"] = mode
        self._apply_mode(self.current_cfg)

    # ---- 绘制 ------------------------------------------------------------

    def _paint_item(self, i):
        it = self.items[i]
        cfg = self.current_cfg
        col = cfg["colors"]
        q = self.quotes.get(it.native) if it.native else None
        pct = q.get("pct") if q else None

        if _is_convertible_bond(it.native):
            # 可转债：只显示价格，不显示涨跌幅
            price = q.get("price") if q else None
            try:
                dec = int(cfg.get("layout", {}).get("price_decimals", 2) or 0)
            except (TypeError, ValueError):
                dec = 2
            pct_txt = ("%%.%df" % dec % price) if price is not None else "--"
            pct_color = col["price"]
        elif pct is None:
            pct_txt, pct_color = "--", col["flat"]
        elif pct > 0:
            pct_txt, pct_color = "%+.2f%%" % pct, col["up"]
        elif pct < 0:
            pct_txt, pct_color = "%+.2f%%" % pct, col["down"]
        else:
            pct_txt, pct_color = "%+.2f%%" % pct, col["flat"]

        name = _short_name(q.get("name") if q else "", it.code) or "--"

        it.lbl_pct.config(text=pct_txt, fg=pct_color)
        it.lbl_name.config(text=name, fg=col["name"])

        bg = col["hover_bg"] if self.hovered == i else col["background"]
        for w in it.widgets:
            w.configure(bg=bg)

    def _paint_all(self):
        for i in range(len(self.items)):
            self._paint_item(i)

    # ---- 几何 / 定位 -----------------------------------------------------

    def _apply_geometry(self, cfg, tray_float=False):
        sb = cfg.get("statusbar", {})
        pos = cfg.get("position", {})
        try:
            pad_x = int(sb.get("pad_x", 10))
            pad_y = int(sb.get("pad_y", 5))
            gap = int(sb.get("gap", 16))
            bw = int(sb.get("border_width", 1))
        except (TypeError, ValueError):
            pad_x, pad_y, gap, bw = 10, 5, 16, 1
        bw = max(0, bw)

        self.container.pack_configure(padx=bw, pady=bw)
        self.inner.pack_configure(padx=pad_x, pady=pad_y)
        for idx, it in enumerate(self.items):
            px = (0, gap) if idx < len(self.items) - 1 else (0, 0)
            it.frame.pack_configure(side="left", padx=px)

        self.root.update_idletasks()
        widths = [it.frame.winfo_reqwidth() for it in self.items]
        heights = [it.frame.winfo_reqheight() for it in self.items]
        n = len(widths)
        content_w = sum(widths) + gap * max(0, n - 1) if n else 0
        content_h = max(heights) if heights else 0
        w = content_w + pad_x * 2 + bw * 2
        h = content_h + pad_y * 2 + bw * 2

        sw, sh = _screen_metrics()
        try:
            margin = int(pos.get("margin", 8))
        except (TypeError, ValueError):
            margin = 8
        try:
            gap_above = int(pos.get("gap_above_taskbar", 2))
        except (TypeError, ValueError):
            gap_above = 2

        # 水平位置（托盘浮窗固定贴右）
        x_mode = "right" if tray_float else pos.get("x", "left")
        if isinstance(x_mode, (int, float)):
            xv = int(x_mode)
            x = (sw + xv - w) if xv < 0 else xv
        elif x_mode == "center":
            x = (sw - w) // 2
        elif x_mode == "right":
            x = sw - w - margin
        else:
            x = margin

        # 垂直位置（托盘浮窗固定贴任务栏上方）
        if tray_float:
            y = self._auto_y(h, gap_above)
        else:
            y_cfg = pos.get("y")
            if y_cfg is None:
                y = self._auto_y(h, gap_above)
            else:
                try:
                    yv = int(y_cfg)
                except (TypeError, ValueError):
                    yv = None
                if yv is None:
                    y = self._auto_y(h, gap_above)
                elif yv < 0:
                    y = sh + yv - h
                else:
                    y = yv

        embed = bool(pos.get("embed_taskbar", True)) and not tray_float
        if embed and _find_taskbar():
            self._apply_embedded_geometry(w, h)
        else:
            self._detach_embed()
            x = max(0, min(x, sw - w))
            y = max(0, min(y, sh - h))
            # 注意：-topmost 必须在 geometry() 之前调用。
            # 否则 -topmost 底层的 SetWindowPos 会把尚未被事件循环处理的
            # geometry 位置强制覆盖回 (0,0)。
            self.root.attributes("-topmost", True)
            self.root.geometry("%dx%d+%d+%d" % (w, h, x, y))
        self._apply_round_corners()

    def _apply_embedded_geometry(self, w, h):
        """把窗口注册为 AppBar，让 Explorer 自动塞进任务栏内部（QQ管家/天气同款）。"""
        try:
            hwnd = self.root.winfo_id()
            if not self.embedded:
                # 取消 WS_EX_TOOLWINDOW 在 AppBar 下可能导致不显示
                GWL_EXSTYLE = -20
                user32 = ctypes.windll.user32
                style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
                WS_EX_TOOLWINDOW = 0x00000080
                user32.SetWindowLongW(hwnd, GWL_EXSTYLE,
                                      style & ~WS_EX_TOOLWINDOW)
                self.appbar.attach(hwnd)
                self.embedded = True
            self._position_embedded(w, h)
        except Exception as e:
            print("[嵌入] AppBar 失败，回退到任务栏上方:", e)
            self._detach_embed()

    def _position_embedded(self, w, h):
        """把窗口位置告诉 AppBar，再 SetWindowPos 实际落到任务栏内。"""
        tb = _find_taskbar()
        if not tb:
            return
        tr = wintypes.RECT()
        ctypes.windll.user32.GetWindowRect(tb, ctypes.byref(tr))
        tb_w = int(tr.right - tr.left)
        tb_h = int(tr.bottom - tr.top)
        try:
            margin = int(self.current_cfg.get("position", {}).get("margin", 8) or 0)
        except (TypeError, ValueError):
            margin = 8

        # 水平方向：贴系统托盘左侧。TrayNotifyWnd 在 Win11 默认合并，
        # 拿不到时退回到任务栏右端 - 宽度 - margin。
        nl = _tray_notify_left()
        x = (tb_w - (nl if nl is not None else tb_w) - w - margin)
        if nl is not None:
            x = nl - w - margin
        else:
            x = tb_w - w - margin
        x = max(0, min(x, tb_w - w))
        y = (tb_h - h) // 2

        # 关键：先调用 ABM_SETPOS 告知系统我们要占用这块区域
        try:
            abd = _APPBARDATA()
            abd.cbSize = ctypes.sizeof(_APPBARDATA)
            abd.hWnd = self.root.winfo_id()
            abd.uEdge = _ABE_BOTTOM
            abd.rc.left = x
            abd.rc.top = y
            abd.rc.right = x + w
            abd.rc.bottom = y + h
            shell32 = ctypes.windll.shell32
            shell32.SHAppBarMessage(_ABM_QUERYPOS, ctypes.byref(abd))
            shell32.SHAppBarMessage(_ABM_SETPOS, ctypes.byref(abd))
        except Exception:
            pass

        SWP_NOACTIVATE = 0x0010
        SWP_SHOWWINDOW = 0x0040
        ctypes.windll.user32.SetWindowPos(
            self.root.winfo_id(), 0, x, y, w, h,
            SWP_NOACTIVATE | SWP_SHOWWINDOW)

    def _detach_embed(self):
        if self.embedded:
            try:
                self.appbar.detach()
            except Exception:
                pass
            self.embedded = False
            self.embedded_hwnd = None
            try:
                self.root.attributes("-topmost", True)
            except Exception:
                pass

    def _apply_round_corners(self):
        """给状态栏窗口设置圆角（SetWindowRgn 裁剪，视觉更贴合任务栏）。"""
        try:
            hwnd = self.root.winfo_id()
            w = self.root.winfo_width()
            h = self.root.winfo_height()
            if w <= 2 or h <= 2:
                return
            try:
                r = int(self.current_cfg.get("statusbar", {}).get("corner_radius", 8) or 0)
            except (TypeError, ValueError):
                r = 8
            r = max(0, min(r, h // 2))
            rgn = ctypes.windll.gdi32.CreateRoundRectRgn(0, 0, w + 1, h + 1, r * 2, r * 2)
            if rgn:
                ctypes.windll.user32.SetWindowRgn(hwnd, rgn, True)
        except Exception as e:
            print("[圆角] 设置失败:", e)

    def _auto_y(self, h, gap_above):
        l, t, r, b = _work_area()
        sw, sh = _screen_metrics()
        if b < sh:  # 任务栏在底部
            return b - gap_above - h
        if t > 0:  # 任务栏在顶部
            return t + gap_above
        return sh - h  # 未检测到任务栏，贴屏幕底部

    # ---- 窗口样式 / 置顶 ------------------------------------------------

    def _apply_window_style(self):
        """设置 WS_EX_NOACTIVATE / WS_EX_TOOLWINDOW，避免抢焦点、不出现在任务栏。"""
        try:
            hwnd = self.root.winfo_id()
            GWL_EXSTYLE = -20
            WS_EX_NOACTIVATE = 0x08000000
            WS_EX_TOOLWINDOW = 0x00000080
            user32 = ctypes.windll.user32
            style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            user32.SetWindowLongW(hwnd, GWL_EXSTYLE,
                                 style | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW)
        except Exception as e:
            print("[窗口] 样式设置失败:", e)

    def _keep_on_top(self):
        if self._closing:
            return
        try:
            if self.embedded:
                w = self.root.winfo_width()
                h = self.root.winfo_height()
                self._apply_embedded_geometry(w, h)
            else:
                self.root.attributes("-topmost", True)
        except Exception:
            pass
        self.root.after(300, self._keep_on_top)

    # ---- 鼠标交互 --------------------------------------------------------

    def _hit_test(self, rx, ry):
        for i, it in enumerate(self.items):
            x0 = it.frame.winfo_rootx()
            y0 = it.frame.winfo_rooty()
            x1 = x0 + it.frame.winfo_width()
            y1 = y0 + it.frame.winfo_height()
            if x0 <= rx < x1 and y0 <= ry < y1:
                return i
        return None

    def _on_motion(self, event):
        self._set_hover(self._hit_test(event.x_root, event.y_root))

    def _set_hover(self, idx):
        if idx == self.hovered:
            return
        old = self.hovered
        self.hovered = idx
        if old is not None and 0 <= old < len(self.items):
            self._paint_item(old)
        if idx is not None and 0 <= idx < len(self.items):
            self._paint_item(idx)
        # 移出标的 -> 关闭分时图
        if idx is None or (old is not None and old != idx):
            self._close_popup()

    def _handle_click(self, event):
        idx = self._hit_test(event.x_root, event.y_root)
        if idx is None:
            return
        now = time.time()
        self.click_times.append(now)
        self.click_times = self.click_times[-3:]
        if len(self.click_times) >= 3 and (self.click_times[-1] - self.click_times[0]) <= 0.6:
            self.shutdown()
            return
        self._show_popup(idx)

    def shutdown(self):
        if self._closing:
            return
        self._closing = True
        print("[退出] 检测到三连击/菜单退出，程序退出。")
        self.stop_event.set()
        if self.tray is not None:
            try:
                self.tray.stop()
            except Exception:
                pass
            self.tray = None
        self.root.after(0, self.root.destroy)

    # ---- 分时图弹窗 ------------------------------------------------------

    def _show_popup(self, idx):
        it = self.items[idx]
        cfg = self.current_cfg
        tl = cfg.get("timeline", {})
        col = cfg["colors"]
        try:
            w = int(tl.get("width", 280))
            h = int(tl.get("height", 130))
        except (TypeError, ValueError):
            w, h = 280, 130
        bg = tl.get("bg", "#1e1e1e")
        border = tl.get("border", "#444444")

        self._close_popup()
        self.popup_req += 1
        req = self.popup_req

        top = tk.Toplevel(self.root)
        top.overrideredirect(True)
        top.attributes("-topmost", True)
        top.configure(bg=bg)
        cv = tk.Canvas(top, width=w, height=h, bg=bg, highlightthickness=0)
        cv.pack()
        cv.create_rectangle(0, 0, w - 1, h - 1, outline=border)
        cv.create_text(w / 2, h / 2, text="加载中…", fill=col["flat"], font=self.font)

        self.popup_win = top
        self.popup_canvas = cv

        # 定位：标的正上方，水平居中，越界时翻转到下方
        x0 = it.frame.winfo_rootx()
        y0 = it.frame.winfo_rooty()
        iw = it.frame.winfo_width()
        ih = it.frame.winfo_height()
        sw, sh = _screen_metrics()
        px = x0 + iw // 2 - w // 2
        py = y0 - h - 4
        if py < 0:
            py = y0 + ih + 4
        px = max(0, min(px, sw - w))
        py = max(0, min(py, sh - h))
        top.geometry("+%d+%d" % (px, py))

        native = it.native
        threading.Thread(target=self._load_timeline, args=(native, req),
                         daemon=True).start()

    def _load_timeline(self, native, req):
        try:
            pre, rows = fetch_timeline(native)
        except Exception as e:
            print("[分时] %s 获取失败: %s" % (native, e))
            pre, rows = None, []
        self.ui_queue.put(("timeline", req, native, pre, rows))

    def _draw_timeline(self, req, native, pre, rows):
        if req != self.popup_req or self.popup_win is None:
            return
        cv = self.popup_canvas
        cv.delete("all")
        cfg = self.current_cfg
        col = cfg["colors"]
        tl = cfg.get("timeline", {})
        border = tl.get("border", "#444444")
        w = cv.winfo_width()
        h = cv.winfo_height()
        cv.create_rectangle(0, 0, w - 1, h - 1, outline=border)

        dec = int(cfg.get("layout", {}).get("price_decimals", 2) or 0)
        q = self.quotes.get(native)
        name = _short_name(q.get("name") if q else "", native)
        price = q.get("price") if q else None
        pct = q.get("pct") if q else None
        if price is None and rows:
            price = rows[-1][1]
        if pre is None:
            pre = q.get("pre") if q else None
        if pct is None and price is not None and pre:
            pct = (price - pre) / pre * 100.0

        ps = ("%%.%df" % dec % price) if price is not None else "--"
        pc = ("%+.2f%%" % pct) if pct is not None else "--"

        if _is_convertible_bond(native):
            # 可转债：标题只显示名称与价格，不带涨跌幅
            color = col["price"]
            title = "%s  %s" % (name, ps)
        else:
            color = col["flat"]
            if pct is not None:
                color = col["up"] if pct > 0 else (col["down"] if pct < 0 else col["flat"])
            title = "%s  %s  %s" % (name, ps, pc)
        cv.create_text(6, 11, text=title, anchor="w", fill=color, font=self.font)

        if not rows:
            cv.create_text(w / 2, h / 2, text="无数据", fill=col["flat"], font=self.font)
            return

        top_area = 24
        bot = h - 4
        left = 4
        right = w - 4
        prices = [r[1] for r in rows]
        avgs = [r[2] for r in rows]
        vals = [v for v in prices if v is not None]
        if pre:
            vals.append(pre)
        vals += [v for v in avgs if v is not None]
        if not vals:
            cv.create_text(w / 2, h / 2, text="无数据", fill=col["flat"], font=self.font)
            return
        vmin, vmax = min(vals), max(vals)
        if vmax - vmin < 1e-9:
            vmax = vmin + 1.0
        n = len(rows)

        def X(i):
            return left + (right - left) * i / (n - 1) if n > 1 else (left + right) / 2

        def Y(v):
            return top_area + (bot - top_area) * (1.0 - (v - vmin) / (vmax - vmin))

        if pre:
            yp = Y(pre)
            cv.create_line(left, yp, right, yp, fill=col["flat"], dash=(3, 3), width=1)

        apts = [(X(i), Y(a)) for i, a in enumerate(avgs) if a is not None]
        if len(apts) > 1:
            cv.create_line(*[c for p in apts for c in p], fill="#d4a017", width=1)

        ppts = [(X(i), Y(p)) for i, p in enumerate(prices) if p is not None]
        if len(ppts) > 1:
            cv.create_line(*[c for p in ppts for c in p], fill=color, width=2)

    def _close_popup(self):
        if self.popup_win is not None:
            try:
                self.popup_win.destroy()
            except Exception:
                pass
            self.popup_win = None
            self.popup_canvas = None


def main():
    _setup_console()
    _set_dpi_awareness()
    socket.setdefaulttimeout(8)
    print("=" * 46)
    print(" 任务栏行情已启动（数据源：东方财富）")
    print(" 三连击任意标的退出程序；请勿关闭本窗口")
    print("=" * 46)
    app = QuoteBar()
    app.run()


if __name__ == "__main__":
    main()
