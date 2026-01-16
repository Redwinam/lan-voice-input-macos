# 局域网语音输入 (LAN Voice Input) - macOS 适配版

一个简单的局域网工具，让你能通过手机的语音输入法，实时将文字发送到电脑端进行输入。

> 适配 macOS 系统，支持自动/手动发送模式，支持语音指令控制。

## ✨ 主要功能

*   **🎙️ 跨设备输入**：利用手机优秀的语音识别能力（如 iOS 自带键盘），直接向电脑输入文字。
*   **🍎 macOS 深度适配**：
    *   使用 macOS 原生通知中心。
    *   利用 `Clipboard` + `Cmd+V` 模拟粘贴，兼容性更强，支持 Emoji 和特殊符号。
    *   启动时自动调用系统“预览”显示连接二维码。
*   **🔒 手动发送模式 (新)**：
    *   支持“仅手动发送”开关。
    *   在手机上确认文字无误后，再一键发送上屏，避免语音识别错误的尴尬。
*   **⚡ 语音指令**：
    *   支持“换行”、“回车”、“删除上一句”等语音控制指令。
*   **🚀 零额外依赖**：无需安装专用 APP，手机浏览器扫码即用。

## 🛠️ 安装与运行

### 环境要求
*   Python 3.8+ (建议 3.10+)
*   macOS (本项目针对 macOS 进行了重构适配)

### 1. 克隆项目
```bash
git clone https://github.com/Redwinam/lan-voice-input-macos.git
cd lan-voice-input-macos
```

### 2. 创建虚拟环境 (推荐)
```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. 安装依赖
```bash
pip install -r requirements.txt
```

### 4. 启动服务
```bash
python server.py
```

启动成功后，会通过系统通知显示连接 URL。你也可以在右上角菜单栏图标中复制 URL 或显示二维码。

## 📦 打包成 macOS 菜单栏常驻应用（.app）

### 1. 安装打包工具
```bash
python3 -m pip install pyinstaller
```

### 2. 生成 .app
在项目根目录执行：
```bash
pyinstaller server.py \
  --name "LAN Voice Input" \
  --windowed \
  --icon icon.icns \
  --add-data "index.html:." \
  --add-data "icon.icns:." \
  --add-data "icon.png:." \
  --clean \
  --noconfirm
```

### 3. 运行
产物在：
* `dist/LAN Voice Input.app`

双击打开后会常驻在右上角菜单栏。

### 4. 菜单栏功能
* 启动服务 / 停止服务（停止会释放端口与资源）
* 复制 URL（用于手机扫码/打开）
* 显示二维码
* 隐藏/显示 Dock 图标

### 5. 开机自启动（可选）
系统设置 → 通用 → 登录项 → 将 `LAN Voice Input.app` 添加到“在登录时打开”。

## 📱 使用方法

1.  **扫码连接**：确保手机和电脑连接在**同一个 WiFi** 下，用手机扫描电脑弹出的二维码。
2.  **开始输入**：
    *   **自动模式**：手机输入框打字/说话，电脑端光标处实时上屏。
    *   **手动模式**：勾选页面上的 `☑️ 仅手动发送`，说完后点击 `📤 手动发送` 按钮上屏。
3.  **语音指令**：
    *   点击 `⚡ 命令` 模式，或直接说出指令（如“换行”、“清空”）。

## ⚠️ 注意事项

*   **输入法设置**：建议电脑端切换到**英文输入法**，避免粘贴时触发中文输入法的联想框。
*   **权限授予**：首次运行时，macOS 可能会请求**“辅助功能”**或**“控制电脑”**的权限，请务必**允许**，否则无法模拟键盘粘贴。
*   **剪贴板占用**：程序通过剪贴板粘贴文字，使用过程中会覆盖你当前的剪贴板内容。

## 📝 协议与致谢
本项目基于 [lan-voice-input](https://github.com/bfilestor/lan-voice-input) 进行 macOS 适配与功能增强。
仅供学习交流使用。
