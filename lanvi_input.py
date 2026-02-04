import platform
import re
import shlex
import subprocess
import time
from collections import deque
from dataclasses import dataclass

import pyautogui

IS_WINDOWS = platform.system() == "Windows"
IS_MACOS = platform.system() == "Darwin"

FORCE_CLICK_BEFORE_TYPE = False
FOCUS_SETTLE_DELAY = 0.06

CLEAR_BACKSPACE_MAX = 200
TEST_INJECT_TEXT = "[SendInput Test] 123 ABC 中文 测试"

SERVER_DEDUP_WINDOW_SEC = 1.2
HISTORY_MAX_LEN = 300

if IS_WINDOWS:
    import ctypes
    from ctypes import wintypes

    if not hasattr(wintypes, "ULONG_PTR"):
        wintypes.ULONG_PTR = ctypes.c_size_t

    user32 = ctypes.WinDLL("user32", use_last_error=True)

    INPUT_KEYBOARD = 1
    KEYEVENTF_KEYUP = 0x0002
    KEYEVENTF_UNICODE = 0x0004

    VK_BACK = 0x08
    VK_TAB = 0x09
    VK_RETURN = 0x0D
    VK_ESCAPE = 0x1B

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

    def press_tab():
        press_vk(VK_TAB, times=1)

    def press_esc():
        press_vk(VK_ESCAPE, times=1)

else:
    import pyperclip

    def send_unicode_text(text: str):
        if not text:
            return
        try:
            pyperclip.copy(text)
            pyautogui.keyDown("command")
            pyautogui.press("v")
            time.sleep(0.03)
        except Exception:
            pass
        finally:
            try:
                pyautogui.keyUp("command")
            except Exception:
                pass

    def backspace(n: int):
        if n > 0:
            pyautogui.press("backspace", presses=n)

    def press_enter():
        pyautogui.press("enter")

    def press_tab():
        pyautogui.press("tab")

    def press_esc():
        pyautogui.press("esc")


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

        if text in ["换行", "回车", "下一行", "enter", "ENTER", "回车键", "enter键", "Enter"]:
            return CommandResult(True, "↩️ 换行", ("__ENTER__", 1))

        if text in ["tab", "TAB", "制表符", "制表", "tab键", "TAB键", "Tab"]:
            return CommandResult(True, "↹ TAB", ("__TAB__", 1))

        if text in ["esc", "ESC", "escape", "ESC键", "esc键", "Escape"]:
            return CommandResult(True, "⎋ ESC", ("__ESC__", 1))

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
        if out[0] == "__TAB__":
            press_tab()
            return
        if out[0] == "__ESC__":
            press_esc()
            return
    if isinstance(out, str):
        send_unicode_text(out)


def focus_target():
    if not FORCE_CLICK_BEFORE_TYPE:
        return
    try:
        pyautogui.click()
        time.sleep(FOCUS_SETTLE_DELAY)
    except Exception:
        pass


def _build_command_args(command, args):
    if isinstance(command, str) and command.strip():
        parts = shlex.split(command, posix=False)
    elif isinstance(command, list):
        parts = [str(x) for x in command if str(x).strip()]
    else:
        parts = []

    if isinstance(args, list):
        parts.extend([str(x) for x in args if str(x).strip()])
    return parts


def _match_command(text, commands):
    text = (text or "").strip()
    if not text:
        return None
    for cmd in commands:
        match_string = (cmd.get("match-string") or "").strip()
        if match_string and match_string == text:
            return cmd
    return None


class InputService:
    def __init__(self, notify, get_commands):
        self.notify = notify
        self.get_commands = get_commands
        self.processor = CommandProcessor()
        self._last_msg = ""
        self._last_time = 0.0

    def _server_dedup(self, text: str) -> bool:
        now = time.time()
        if text == self._last_msg and (now - self._last_time) < SERVER_DEDUP_WINDOW_SEC:
            return True
        self._last_msg = text
        self._last_time = now
        return False

    def handle_text(self, text: str):
        text = (text or "").strip()
        if not text:
            return

        if self._server_dedup(text):
            print("⏭️ 服务器去重：", text)
            return

        if text == "__TEST_INJECT__":
            self.notify("测试注入", "请将鼠标放在记事本输入区，正在注入测试文本…")
            focus_target()
            try:
                send_unicode_text(TEST_INJECT_TEXT)
                press_enter()
                send_unicode_text("✅ 如果你看到这行文字，说明 SendInput 注入成功！")
                press_enter()
                self.notify("测试注入成功", "请查看记事本是否出现两行测试文本。")
            except Exception as e:
                self.notify("测试注入失败", str(e))
            return

        result = self.processor.handle(text)
        if result.output == "":
            self.notify("指令执行", result.display_text)
            return

        if isinstance(result.output, str):
            focus_target()
        execute_output(result.output)

        if not result.handled and isinstance(result.output, str):
            self.processor.record_output(result.output)

    def execute_command(self, text: str) -> CommandResult:
        cmd = _match_command(text, self.get_commands() or [])
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
