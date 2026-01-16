# server.py
# -*- coding: utf-8 -*-
import asyncio
import json
import os
import re
import socket
import subprocess
import sys
import threading
import time
import tempfile
from collections import deque
from dataclasses import dataclass
from typing import List, Tuple, Optional
import shlex
import queue

import pyautogui
import pystray
import qrcode
import websockets
from PIL import Image
from flask import Flask, send_file, jsonify
from pystray import MenuItem as item
from werkzeug.serving import make_server
from websockets.exceptions import ConnectionClosed, ConnectionClosedError, ConnectionClosedOK

import platform

IS_WINDOWS = platform.system() == "Windows"
IS_MACOS = platform.system() == "Darwin"

# Windows Toast：winotify
if IS_WINDOWS:
    try:
        from winotify import Notification
        WINOTIFY_AVAILABLE = True
    except Exception:
        WINOTIFY_AVAILABLE = False
else:
    WINOTIFY_AVAILABLE = False

if IS_WINDOWS:
    import ctypes
    from ctypes import wintypes


# ===================== 默认端口（自动选择可用）=====================
DEFAULT_HTTP_PORT = 8080
DEFAULT_WS_PORT = 8765
MAX_PORT_TRY = 50

# ===================== 行为配置 =====================
FORCE_CLICK_BEFORE_TYPE = True
FOCUS_SETTLE_DELAY = 0.06

CLEAR_BACKSPACE_MAX = 200
TEST_INJECT_TEXT = "[SendInput Test] 123 ABC 中文 测试"

SERVER_DEDUP_WINDOW_SEC = 1.2
HISTORY_MAX_LEN = 300

# WebSocket 心跳（让断线更快被识别）
WS_PING_INTERVAL = 20
WS_PING_TIMEOUT = 10

# ===================== 全局状态 =====================
HTTP_PORT: Optional[int] = None
WS_PORT: Optional[int] = None
QR_URL: Optional[str] = None
QR_PAYLOAD_URL: Optional[str] = None

tray_icon = None

CLIENT_COUNT = 0
CLIENT_LOCK = threading.Lock()

# ===================== 服务生命周期（启动/停止）=====================
SERVICE_LOCK = threading.Lock()
SERVICE_RUNNING = False

HTTP_SERVER = None
HTTP_THREAD = None

WS_LOOP = None
WS_THREAD = None
WS_SERVER = None

DOCK_ICON_HIDDEN = False

# ✅ 用户手动选择的 IP（None = 自动）
USER_IP: Optional[str] = None
CONFIG_DATA: dict = {}
COMMANDS: List[dict] = []


# ===================== PyInstaller 路径工具 =====================
def is_frozen() -> bool:
    return getattr(sys, "frozen", False) is True


def get_exe_dir() -> str:
    """打包后：exe 同级目录；源码：server.py 同级目录"""
    if is_frozen():
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_resource_dir() -> str:
    """
    资源目录：
    - onefile 打包：sys._MEIPASS（解压到临时目录，index.html 在这里）
    - 其他情况：server.py 同级目录
    """
    if is_frozen() and hasattr(sys, "_MEIPASS"):
        return getattr(sys, "_MEIPASS")
    return os.path.dirname(os.path.abspath(__file__))


def resource_path(name: str) -> str:
    return os.path.join(get_resource_dir(), name)


# ===================== 配置持久化（优先写 exe 同级 config.json，写失败 fallback 到用户目录）=====================
CONFIG_PATH_PRIMARY = os.path.join(get_exe_dir(), "config.json")
CONFIG_PATH_FALLBACK = os.path.join(os.path.expanduser("~"), "LanVI_config.json")
CONFIG_PATH_IN_USE = CONFIG_PATH_PRIMARY  # 运行时可能切到 fallback


def _try_write_json(path: str, data: dict) -> bool:
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def _try_read_json(path: str) -> Optional[dict]:
    try:
        if not os.path.exists(path):
            return None
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _normalize_commands(raw) -> List[dict]:
    if not isinstance(raw, list):
        return []
    return [c for c in raw if isinstance(c, dict)]


def load_config():
    """
    启动时读取 config：
    - 优先 exe 同级 config.json
    - 否则读取用户目录 LanVI_config.json
    - 两边都没有：创建（优先主路径，失败则 fallback）
    """
    global USER_IP, CONFIG_PATH_IN_USE, CONFIG_DATA, COMMANDS

    # 先读主路径
    data = _try_read_json(CONFIG_PATH_PRIMARY)
    if isinstance(data, dict):
        CONFIG_DATA = data
        COMMANDS = _normalize_commands(data.get("commands"))
        ip = (data.get("user_ip") or "").strip()
        USER_IP = ip if ip else None
        CONFIG_PATH_IN_USE = CONFIG_PATH_PRIMARY
        return

    # 再读 fallback
    data = _try_read_json(CONFIG_PATH_FALLBACK)
    if isinstance(data, dict):
        CONFIG_DATA = data
        COMMANDS = _normalize_commands(data.get("commands"))
        ip = (data.get("user_ip") or "").strip()
        USER_IP = ip if ip else None
        CONFIG_PATH_IN_USE = CONFIG_PATH_FALLBACK
        return

    # 都没有：创建默认（自动）
    USER_IP = None
    CONFIG_DATA = {"user_ip": None, "commands": []}
    COMMANDS = []
    save_config()


def save_config():
    """
    保存当前 USER_IP：
    - 优先写 exe 同级 config.json（你期望的位置）
    - 若无权限/失败：写到用户目录，并切换 CONFIG_PATH_IN_USE
    """
    global CONFIG_PATH_IN_USE, CONFIG_DATA, COMMANDS
    data = dict(CONFIG_DATA) if isinstance(CONFIG_DATA, dict) else {}
    data["user_ip"] = USER_IP
    data["commands"] = COMMANDS

    # 优先写主路径（exe 同级）
    if _try_write_json(CONFIG_PATH_PRIMARY, data):
        CONFIG_PATH_IN_USE = CONFIG_PATH_PRIMARY
        return

    # 主路径失败则写 fallback（保证一定能保存）
    if _try_write_json(CONFIG_PATH_FALLBACK, data):
        CONFIG_PATH_IN_USE = CONFIG_PATH_FALLBACK
        return


# ===================== 通知封装 =====================
def notify(title: str, msg: str, duration=3):
    """托盘气泡 + 系统原生通知，永不抛异常影响主程序"""
    global tray_icon

    # 托盘气泡（稳定兜底）
    try:
        if tray_icon:
            tray_icon.notify(msg, title)
    except Exception:
        pass

    # Windows Toast（winotify）
    if IS_WINDOWS and WINOTIFY_AVAILABLE:
        def _toast():
            try:
                toast = Notification(
                    app_id="LAN Voice Input",
                    title=title,
                    msg=msg,
                    duration="short"
                )
                toast.show()
            except Exception:
                pass
        threading.Thread(target=_toast, daemon=True).start()
    
    # macOS Notification (osascript)
    if IS_MACOS:
        _enqueue_macos_notification(title, msg)


_macos_notify_queue = queue.SimpleQueue()
_macos_notify_started = False
_macos_notify_lock = threading.Lock()


def _ensure_macos_notify_worker():
    global _macos_notify_started
    if _macos_notify_started:
        return
    with _macos_notify_lock:
        if _macos_notify_started:
            return
        threading.Thread(target=_macos_notify_worker, daemon=True).start()
        _macos_notify_started = True


def _enqueue_macos_notification(title: str, msg: str):
    _ensure_macos_notify_worker()
    _macos_notify_queue.put((str(title), str(msg)))


def _macos_notify_worker():
    while True:
        title, msg = _macos_notify_queue.get()
        try:
            safe_title = str(title).replace("\\", "\\\\").replace('"', '\\"')
            safe_msg = str(msg).replace("\\", "\\\\").replace('"', '\\"')
            script = f'display notification "{safe_msg}" with title "{safe_title}"'
            subprocess.run(["osascript", "-e", script], timeout=2)
        except Exception:
            pass


# ===================== 自动选择可用端口 =====================
def is_port_free(port: int) -> bool:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(("0.0.0.0", port))
            return True
    except OSError:
        return False


def choose_free_port(start_port: int) -> int:
    for p in range(start_port, start_port + MAX_PORT_TRY):
        if is_port_free(p):
            return p
    raise RuntimeError(f"找不到可用端口（从 {start_port} 起尝试 {MAX_PORT_TRY} 个）")


# ===================== IP & 网卡枚举 =====================
def get_lan_ip_best_effort() -> str:
    """通过 UDP “假连接”拿到默认出口网卡 IP（不真正发包）。"""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = "127.0.0.1"
    finally:
        s.close()
    return ip


def is_valid_ipv4(ip: str) -> bool:
    if not ip:
        return False
    if not re.match(r"^\d{1,3}(\.\d{1,3}){3}$", ip):
        return False
    parts = ip.split(".")
    try:
        nums = [int(x) for x in parts]
    except Exception:
        return False
    return all(0 <= n <= 255 for n in nums)


def is_candidate_ipv4(ip: str) -> bool:
    if not is_valid_ipv4(ip):
        return False
    if ip.startswith("127.") or ip.startswith("0.") or ip.startswith("169.254."):
        return False
    return True


def parse_windows_ipconfig() -> List[Tuple[str, str]]:
    """
    Windows：解析 ipconfig，尽量拿到 "网卡名 + IPv4"
    返回 [(label, ip), ...]
    """
    if not IS_WINDOWS:
        return []

    out = ""
    for enc in ("gbk", "utf-8"):
        try:
            out = subprocess.check_output(
                ["ipconfig"], stderr=subprocess.STDOUT, text=True, encoding=enc, errors="ignore"
            )
            if out:
                break
        except Exception:
            continue
    if not out:
        return []

    results: List[Tuple[str, str]] = []
    current_iface = "未知网卡"

    iface_pat = re.compile(r"^\s*([^\r\n:]{3,}adapter\s+.+):\s*$", re.IGNORECASE)
    ipv4_pat = re.compile(r"IPv4.*?:\s*([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)")

    for line in out.splitlines():
        m_iface = iface_pat.match(line.strip())
        if m_iface:
            current_iface = m_iface.group(1).strip()
            continue

        m_ip = ipv4_pat.search(line)
        if m_ip:
            ip = m_ip.group(1).strip()
            if is_candidate_ipv4(ip):
                results.append((f"{current_iface} - {ip}", ip))

    seen = set()
    dedup = []
    for label, ip in results:
        if ip not in seen:
            seen.add(ip)
            dedup.append((label, ip))
    return dedup


def get_ipv4_candidates() -> List[Tuple[str, str]]:
    """
    综合获取候选 IP：
    1) Windows: ipconfig（含网卡名）
    2) hostname 的 IPv4
    3) 自动推荐（默认出口）
    """
    candidates: List[Tuple[str, str]] = []
    if IS_WINDOWS:
        candidates.extend(parse_windows_ipconfig())


    try:
        hostname = socket.gethostname()
        infos = socket.getaddrinfo(hostname, None, family=socket.AF_INET, type=socket.SOCK_STREAM)
        for info in infos:
            ip = info[4][0]
            if is_candidate_ipv4(ip):
                candidates.append((f"{hostname} - {ip}", ip))
    except Exception:
        pass

    ip2 = get_lan_ip_best_effort()
    if is_candidate_ipv4(ip2):
        candidates.append((f"自动推荐（默认出口） - {ip2}", ip2))

    seen = set()
    dedup: List[Tuple[str, str]] = []
    for label, ip in candidates:
        if ip not in seen:
            seen.add(ip)
            dedup.append((label, ip))

    if not dedup:
        dedup = [("本机回环（仅本机可用） - 127.0.0.1", "127.0.0.1")]
    return dedup


# ===================== URL 构建 =====================
def get_effective_ip() -> str:
    global USER_IP
    if USER_IP and USER_IP.strip():
        return USER_IP.strip()
    return get_lan_ip_best_effort()


def build_urls(ip: str):
    global QR_URL, QR_PAYLOAD_URL
    QR_URL = f"http://{ip}:{HTTP_PORT}"
    QR_PAYLOAD_URL = f"{QR_URL}?ws={WS_PORT}"


# ===================== Tk 二维码窗口（内置网卡选择 + 同步刷新）=====================
# macOS 不使用 Tkinter，避免主线程冲突和崩坏
# 替代方案：生成二维码图片并调用系统预览打开，或仅在终端输出

def open_qr_image(url):
    try:
        qr = qrcode.QRCode(box_size=10, border=4)
        qr.add_data(url)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        
        base_dir = get_exe_dir()
        path = os.path.join(base_dir, "qr_code.png")
        try:
            img.save(path)
        except Exception:
            path = os.path.join(tempfile.gettempdir(), "lan_voice_input_qr_code.png")
            img.save(path)
        
        # 打开图片
        if IS_MACOS:
            subprocess.run(["open", path])
        elif IS_WINDOWS:
            os.startfile(path)
        
    except Exception as e:
        print("无法打开二维码图片：", e)

# ===================== Input Control (Cross Platform) =====================

# --- Windows Implementation ---
if IS_WINDOWS:
    if not hasattr(wintypes, "ULONG_PTR"):
        wintypes.ULONG_PTR = ctypes.c_size_t

    user32 = ctypes.WinDLL("user32", use_last_error=True)

    INPUT_KEYBOARD = 1
    KEYEVENTF_KEYUP = 0x0002
    KEYEVENTF_UNICODE = 0x0004

    VK_BACK = 0x08
    VK_RETURN = 0x0D

    class MOUSEINPUT(ctypes.Structure):
        _fields_ = [
            ("dx", wintypes.LONG),
            ("dy", wintypes.LONG),
            ("mouseData", wintypes.DWORD),
            ("dwFlags", wintypes.DWORD),
            ("time", wintypes.DWORD),
            ("dwExtraInfo", wintypes.ULONG_PTR),
        ]

    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [
            ("wVk", wintypes.WORD),
            ("wScan", wintypes.WORD),
            ("dwFlags", wintypes.DWORD),
            ("time", wintypes.DWORD),
            ("dwExtraInfo", wintypes.ULONG_PTR),
        ]

    class HARDWAREINPUT(ctypes.Structure):
        _fields_ = [
            ("uMsg", wintypes.DWORD),
            ("wParamL", wintypes.WORD),
            ("wParamH", wintypes.WORD),
        ]

    class _INPUTunion(ctypes.Union):
        _fields_ = [
            ("mi", MOUSEINPUT),
            ("ki", KEYBDINPUT),
            ("hi", HARDWAREINPUT),
        ]

    class INPUT(ctypes.Structure):
        _anonymous_ = ("union",)
        _fields_ = [("type", wintypes.DWORD), ("union", _INPUTunion)]

    def _send_input(inputs):
        n = len(inputs)
        arr = (INPUT * n)(*inputs)
        cb = ctypes.sizeof(INPUT)
        sent = user32.SendInput(n, arr, cb)
        if sent != n:
            err = ctypes.get_last_error()
            raise ctypes.WinError(err)

    def send_unicode_text(text: str):
        inputs = []
        for ch in text:
            code = ord(ch)
            inputs.append(INPUT(
                type=INPUT_KEYBOARD,
                ki=KEYBDINPUT(wVk=0, wScan=code, dwFlags=KEYEVENTF_UNICODE, time=0, dwExtraInfo=0)
            ))
            inputs.append(INPUT(
                type=INPUT_KEYBOARD,
                ki=KEYBDINPUT(wVk=0, wScan=code, dwFlags=KEYEVENTF_UNICODE | KEYEVENTF_KEYUP, time=0, dwExtraInfo=0)
            ))
        _send_input(inputs)

    def press_vk(vk_code: int, times: int = 1):
        for _ in range(times):
            down = INPUT(type=INPUT_KEYBOARD, ki=KEYBDINPUT(wVk=vk_code, wScan=0, dwFlags=0, time=0, dwExtraInfo=0))
            up = INPUT(type=INPUT_KEYBOARD, ki=KEYBDINPUT(wVk=vk_code, wScan=0, dwFlags=KEYEVENTF_KEYUP, time=0, dwExtraInfo=0))
            _send_input([down, up])

    def backspace(n: int):
        if n > 0:
            press_vk(VK_BACK, times=n)

    def press_enter():
        press_vk(VK_RETURN, times=1)

# --- macOS / Other Implementation ---
else:
    # 依赖 pyautogui / pyperclip
    # 确保已安装: pip install pyperclip
    import pyperclip

    def send_unicode_text(text: str):
        """
        macOS 下模拟键盘输入 Unicode 最稳妥的方式：
        复制到剪贴板 -> 模拟 Cmd+V
        """
        if not text:
            return
        
        try:
            pyperclip.copy(text)
            # macOS 使用 command+v
            pyautogui.hotkey('command', 'v')
        except Exception as e:
            print(f"Error sending text: {e}")

    def backspace(n: int):
        if n > 0:
            pyautogui.press('backspace', presses=n)

    def press_enter():
        pyautogui.press('enter')


# ===================== 指令系统 =====================
@dataclass
class CommandResult:
    handled: bool
    display_text: str = ""
    output: object = ""


class CommandProcessor:
    def __init__(self):
        self.paused = False
        self.history = deque(maxlen=HISTORY_MAX_LEN)
        self.alias = {"豆号": "逗号", "都好": "逗号", "据号": "句号", "聚好": "句号", "句点": "句号"}
        self.punc_map = {"逗号": "，", "句号": "。", "问号": "？", "感叹号": "！", "冒号": "：", "分号": "；", "顿号": "、"}

    def normalize(self, text: str) -> str:
        text = (text or "").strip()
        for k, v in self.alias.items():
            text = text.replace(k, v)
        return text

    def parse_delete_n(self, text: str):
        m = re.search(r"(删除|退格)\s*(\d+)\s*(个字|次)?", text)
        return int(m.group(2)) if m else None

    def handle(self, raw_text: str) -> CommandResult:
        text = self.normalize(raw_text)

        if text in ["暂停输入", "暂停", "停止输入"]:
            self.paused = True
            return CommandResult(True, "⏸ 已暂停输入", "")

        if text in ["继续输入", "继续", "恢复输入"]:
            self.paused = False
            return CommandResult(True, "▶️ 已恢复输入", "")

        if self.paused:
            return CommandResult(True, f"⏸(暂停中) {raw_text}", "")

        if text in ["换行", "回车", "下一行"]:
            return CommandResult(True, "↩️ 换行", ("__ENTER__", 1))

        if text in self.punc_map:
            return CommandResult(True, f"⌨️ {text}", self.punc_map[text])

        if text in ["删除上一句", "撤回上一句", "撤销上一句", "删掉上一句"]:
            if not self.history:
                return CommandResult(True, "⚠️ 没有可删除的内容", "")
            last = self.history.pop()
            return CommandResult(True, f"⌫ 删除上一句：{last}", ("__BACKSPACE__", len(last)))

        n = self.parse_delete_n(text)
        if n is not None:
            return CommandResult(True, f"⌫ 删除 {n} 个字", ("__BACKSPACE__", n))

        if text in ["清空", "清除全部", "全部删除"]:
            return CommandResult(True, "🧹 清空", ("__BACKSPACE__", CLEAR_BACKSPACE_MAX))

        return CommandResult(False, raw_text, raw_text)

    def record_output(self, out: str):
        if out and out != "\n":
            out = str(out)
            if len(out) > 4000:
                out = out[:4000]
            self.history.append(out)


processor = CommandProcessor()


def execute_output(out):
    if out == "":
        return
    if isinstance(out, tuple):
        if out[0] == "__BACKSPACE__":
            backspace(int(out[1]))
            return
        if out[0] == "__ENTER__":
            press_enter()
            return
    if isinstance(out, str):
        send_unicode_text(out)


def focus_target():
    if not FORCE_CLICK_BEFORE_TYPE:
        return
    try:
        x, y = pyautogui.position()
        pyautogui.click(x, y)
        time.sleep(FOCUS_SETTLE_DELAY)
    except Exception:
        pass


_last_msg = ""
_last_time = 0.0


def server_dedup(text: str) -> bool:
    global _last_msg, _last_time
    now = time.time()
    if text == _last_msg and (now - _last_time) < SERVER_DEDUP_WINDOW_SEC:
        return True
    _last_msg = text
    _last_time = now
    return False


def handle_text(text: str):
    text = (text or "").strip()
    if not text:
        return

    if server_dedup(text):
        print("⏭️ 服务器去重：", text)
        return

    if text == "__TEST_INJECT__":
        notify("测试注入", "请将鼠标放在记事本输入区，正在注入测试文本…")
        focus_target()
        try:
            send_unicode_text(TEST_INJECT_TEXT)
            press_enter()
            send_unicode_text("✅ 如果你看到这行文字，说明 SendInput 注入成功！")
            press_enter()
            notify("测试注入成功", "请查看记事本是否出现两行测试文本。")
        except Exception as e:
            notify("测试注入失败", str(e))
        return

    result = processor.handle(text)
    if result.output == "":
        notify("指令执行", result.display_text)
        return

    focus_target()
    execute_output(result.output)

    if not result.handled and isinstance(result.output, str):
        processor.record_output(result.output)


def _build_command_args(command, args) -> List[str]:
    if isinstance(command, str) and command.strip():
        parts = shlex.split(command, posix=False)
    elif isinstance(command, list):
        parts = [str(x) for x in command if str(x).strip()]
    else:
        parts = []

    if isinstance(args, list):
        parts.extend([str(x) for x in args if str(x).strip()])
    return parts


def _match_command(text: str) -> Optional[dict]:
    text = (text or "").strip()
    if not text:
        return None
    for cmd in COMMANDS:
        match_string = (cmd.get("match-string") or "").strip()
        if match_string and match_string == text:
            return cmd
    return None


def execute_command(text: str) -> CommandResult:
    cmd = _match_command(text)
    if not cmd:
        return CommandResult(True, f"未找到匹配指令：{text}", {"ok": False, "message": "未找到匹配指令"})

    args = _build_command_args(cmd.get("command"), cmd.get("args"))
    if not args:
        return CommandResult(True, f"命令配置错误：{text}", {"ok": False, "message": "命令配置错误"})

    try:
        completed = subprocess.run(args, capture_output=True, text=True)
        ok = completed.returncode == 0
        stderr = (completed.stderr or "").strip()
        if ok:
            msg = f"指令执行成功：{text}"
        else:
            msg = f"指令执行失败：{text}（exit {completed.returncode}）"
            if stderr:
                msg = f"{msg} - {stderr}"
        return CommandResult(True, msg, {"ok": ok, "message": msg})
    except Exception as e:
        return CommandResult(True, f"指令执行异常：{text} - {e}", {"ok": False, "message": f"指令执行异常：{e}"})


# ===================== WebSocket =====================
async def ws_handler(websocket):
    global CLIENT_COUNT

    with CLIENT_LOCK:
        CLIENT_COUNT += 1
        c = CLIENT_COUNT
    notify("手机已连接", f"连接数：{c}（HTTP:{HTTP_PORT} WS:{WS_PORT}）")

    try:
        async for msg in websocket:
            msg = msg.strip()
            if not msg:
                continue
            print("收到：", msg)
            msg_type = "text"
            content = msg
            if msg.startswith("{"):
                try:
                    payload = json.loads(msg)
                    if isinstance(payload, dict):
                        msg_type = (payload.get("type") or "text").strip()
                        content = payload.get("string")
                except Exception:
                    msg_type = "text"
                    content = msg

            if msg_type == "cmd":
                result = execute_command(str(content or "").strip())
                resp = {
                    "type": "cmd_result",
                    "string": str(content or "").strip(),
                    "ok": bool(result.output.get("ok")) if isinstance(result.output, dict) else False,
                    "message": result.output.get("message") if isinstance(result.output, dict) else result.display_text,
                }
                await websocket.send(json.dumps(resp, ensure_ascii=False))
            else:
                handle_text(str(content or ""))

    except (ConnectionClosedOK, ConnectionClosedError, ConnectionClosed, ConnectionResetError, OSError):
        pass

    finally:
        with CLIENT_LOCK:
            CLIENT_COUNT -= 1
            c = CLIENT_COUNT
        notify("手机已断开", f"连接数：{c}")


async def ws_main():
    async with websockets.serve(
        ws_handler, "0.0.0.0", WS_PORT,
        ping_interval=WS_PING_INTERVAL,
        ping_timeout=WS_PING_TIMEOUT,
        max_size=1_000_000,
        max_queue=32,
        compression=None,
    ):
        print(f"WebSocket running at ws://0.0.0.0:{WS_PORT}")
        await asyncio.Future()


# ===================== HTTP =====================
app = Flask(__name__)


@app.route("/")
def index():
    # 打包后 index.html 在 sys._MEIPASS（onefile 临时解压目录）
    path = resource_path("index.html")
    response = send_file(path)
    # 禁止缓存，确保前端更新立即可见
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.route("/config")
def config():
    return jsonify({"ws_port": WS_PORT, "http_port": HTTP_PORT, "url": QR_PAYLOAD_URL})


def run_http():
    app.run(host="0.0.0.0", port=HTTP_PORT, debug=False, use_reloader=False)


def _run_http_server_forever():
    global HTTP_SERVER
    if HTTP_PORT is None:
        raise RuntimeError("HTTP_PORT 未初始化")
    HTTP_SERVER = make_server("0.0.0.0", HTTP_PORT, app)
    HTTP_SERVER.serve_forever()


def _ws_thread_main(ready_evt: threading.Event):
    global WS_LOOP, WS_SERVER
    if WS_PORT is None:
        raise RuntimeError("WS_PORT 未初始化")
    loop = asyncio.new_event_loop()
    WS_LOOP = loop
    asyncio.set_event_loop(loop)

    async def _start_ws_server():
        return await websockets.serve(
            ws_handler, "0.0.0.0", WS_PORT,
            ping_interval=WS_PING_INTERVAL,
            ping_timeout=WS_PING_TIMEOUT,
            max_size=1_000_000,
            max_queue=32,
            compression=None,
        )

    WS_SERVER = loop.run_until_complete(_start_ws_server())
    ready_evt.set()
    try:
        loop.run_forever()
    finally:
        try:
            if WS_SERVER:
                WS_SERVER.close()
                loop.run_until_complete(WS_SERVER.wait_closed())
        except Exception:
            pass
        try:
            pending = asyncio.all_tasks(loop)
            for task in pending:
                task.cancel()
            if pending:
                loop.run_until_complete(asyncio.gather(*pending, return_exceptions=True))
        except Exception:
            pass
        try:
            loop.close()
        except Exception:
            pass


def is_service_running() -> bool:
    with SERVICE_LOCK:
        return bool(SERVICE_RUNNING)


def start_services(open_qr: bool = False):
    global HTTP_PORT, WS_PORT, QR_URL, QR_PAYLOAD_URL
    global HTTP_THREAD, WS_THREAD
    global SERVICE_RUNNING

    with SERVICE_LOCK:
        if SERVICE_RUNNING:
            return

        HTTP_PORT = choose_free_port(DEFAULT_HTTP_PORT)
        WS_PORT = choose_free_port(DEFAULT_WS_PORT)
        build_urls(get_effective_ip())
        SERVICE_RUNNING = True

    print("\n======================================")
    print("✅ 已启动（服务已开启）")
    print("📱 手机打开：", QR_PAYLOAD_URL)
    print("HTTP:", HTTP_PORT, "WS:", WS_PORT)
    print("======================================\n")

    HTTP_THREAD = threading.Thread(target=_run_http_server_forever, daemon=True)
    HTTP_THREAD.start()

    ws_ready = threading.Event()
    WS_THREAD = threading.Thread(target=lambda: _ws_thread_main(ws_ready), daemon=True)
    WS_THREAD.start()
    ws_ready.wait(timeout=3)

    notify("LANVoiceInput 服务已启动", f"URL:\n{QR_PAYLOAD_URL}\n\nHTTP:{HTTP_PORT}  WS:{WS_PORT}")
    if open_qr and QR_PAYLOAD_URL:
        threading.Timer(0.3, lambda: open_qr_image(QR_PAYLOAD_URL)).start()


def stop_services():
    global HTTP_PORT, WS_PORT, QR_URL, QR_PAYLOAD_URL
    global HTTP_SERVER, HTTP_THREAD
    global WS_SERVER, WS_LOOP, WS_THREAD
    global SERVICE_RUNNING

    with SERVICE_LOCK:
        if not SERVICE_RUNNING:
            return
        SERVICE_RUNNING = False

    try:
        if HTTP_SERVER:
            HTTP_SERVER.shutdown()
            HTTP_SERVER.server_close()
    except Exception:
        pass
    HTTP_SERVER = None
    HTTP_THREAD = None

    try:
        if WS_LOOP:
            async def _shutdown_ws():
                try:
                    if WS_SERVER:
                        WS_SERVER.close()
                        await WS_SERVER.wait_closed()
                except Exception:
                    pass

            fut = asyncio.run_coroutine_threadsafe(_shutdown_ws(), WS_LOOP)
            try:
                fut.result(timeout=3)
            except Exception:
                pass
            try:
                WS_LOOP.call_soon_threadsafe(WS_LOOP.stop)
            except Exception:
                pass
    except Exception:
        pass

    try:
        if WS_THREAD and WS_THREAD.is_alive():
            WS_THREAD.join(timeout=3)
    except Exception:
        pass

    WS_SERVER = None
    WS_LOOP = None
    WS_THREAD = None

    HTTP_PORT = None
    WS_PORT = None
    QR_URL = None
    QR_PAYLOAD_URL = None

    notify("LANVoiceInput 服务已停止", "已关闭 HTTP/WebSocket 监听，释放端口资源")


def tray_show_qr(icon, _):
    if not is_service_running() or not QR_PAYLOAD_URL:
        notify("服务未启动", "请先在菜单栏选择“启动服务”")
        return
    open_qr_image(QR_PAYLOAD_URL)


def tray_toggle_service(icon, _):
    if is_service_running():
        stop_services()
    else:
        start_services(open_qr=False)
    try:
        icon.update_menu()
    except Exception:
        pass


def tray_start_stop_text(_=None):
    return "停止服务" if is_service_running() else "启动服务"


def tray_copy_url(icon, _):
    if not is_service_running() or not QR_PAYLOAD_URL:
        notify("服务未启动", "请先在菜单栏选择“启动服务”")
        return
    ok = copy_text_to_clipboard(QR_PAYLOAD_URL)
    if ok:
        notify("已复制 URL", QR_PAYLOAD_URL)
    else:
        notify("复制失败", "当前系统不支持自动复制")


def copy_text_to_clipboard(text: str) -> bool:
    try:
        if IS_MACOS:
            subprocess.run(["pbcopy"], input=str(text), text=True, check=False)
            return True
        if IS_WINDOWS:
            subprocess.run(["cmd", "/c", "clip"], input=str(text), text=True, check=False)
            return True
        return False
    except Exception:
        return False


def set_dock_icon_hidden(hidden: bool) -> bool:
    if not IS_MACOS:
        return False
    try:
        from rubicon.objc.runtime import load_library
        load_library("AppKit")
        from rubicon.objc import ObjCClass
        NSApplication = ObjCClass("NSApplication")
        app = NSApplication.sharedApplication
        policy = 1 if hidden else 0
        return bool(app.setActivationPolicy_(policy))
    except Exception:
        return False


def tray_toggle_dock_icon(icon, _):
    global DOCK_ICON_HIDDEN
    target = not bool(DOCK_ICON_HIDDEN)
    ok = set_dock_icon_hidden(target)
    if ok:
        DOCK_ICON_HIDDEN = target
        try:
            icon.update_menu()
        except Exception:
            pass
    else:
        notify("切换失败", "当前环境不支持隐藏/显示 Dock 图标")


def tray_dock_checked(_=None):
    return bool(DOCK_ICON_HIDDEN)


def tray_dock_text(_=None):
    return "显示 Dock 图标" if bool(DOCK_ICON_HIDDEN) else "隐藏 Dock 图标"



def tray_quit(icon, _):
    try:
        stop_services()
    except Exception:
        pass
    notify("退出", "LAN Voice Input 已退出")
    icon.stop()
    os._exit(0)


def run_tray():
    global tray_icon
    candidate_paths = []
    if IS_MACOS:
        candidate_paths.append(resource_path("icon.icns"))
    candidate_paths.append(resource_path("logo.png"))
    candidate_paths.append(resource_path("icon.ico"))
    imagePath = next((p for p in candidate_paths if os.path.exists(p)), None)
        
    menu = (
        item(tray_start_stop_text, tray_toggle_service),
        item("复制 URL", tray_copy_url, enabled=lambda _: is_service_running() and bool(QR_PAYLOAD_URL)),
        item("显示二维码", tray_show_qr),
        item(tray_dock_text, tray_toggle_dock_icon, checked=lambda _: tray_dock_checked(), enabled=lambda _: IS_MACOS),
        item("退出", tray_quit),
    )
    image = Image.open(imagePath) if imagePath else Image.new("RGB", (64, 64), (0, 0, 0))
    tray_icon = pystray.Icon("LANVoiceInput", image, "LAN Voice Input", menu)
    tray_icon.on_double_click = tray_show_qr
    tray_icon.run()


# ===================== main =====================
if __name__ == "__main__":
    # ✅ 启动即读取/创建 config（打包后优先 exe 同级 config.json）
    load_config()
    print("\n======================================")
    print("CONFIG(primary):", CONFIG_PATH_PRIMARY)
    print("CONFIG(fallback):", CONFIG_PATH_FALLBACK)
    print("CONFIG(in use):", CONFIG_PATH_IN_USE)
    print("======================================\n")

    start_services(open_qr=False)

    run_tray()
