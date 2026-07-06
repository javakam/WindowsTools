# AGENTS.md

## 项目规则

- 中文回复用户。
- 这是一个 Python + Tkinter 的 Windows 本地小工具，目标是简单、稳定、轻量。
- 优先使用 Python 标准库、Tkinter、ctypes、pywin32 和现有代码风格。
- 不要过度设计，不要把项目扩展成复杂的软件管理器。
- 不要引入 Web 框架、前端框架、数据库、后台服务等重型依赖。
- 不要随意创建新的 Markdown/TXT 文档，除非用户明确要求。
- 修改主程序或打包可执行文件时，同步更新 `windows_tools.py` 中的 `APP_TITLE` 日期，格式为 `WindowsToolsYYYYMMDDV1`。
- 每次修改内容后，提交 Git；只提交本次相关文件，不混入无关改动。

## 功能边界

- 只关注开始菜单和桌面扫描、搜索、来源筛选、快捷方式解析、管理员启动和明确错误提示。
- 默认不要增加卸载管理、注册表深度扫描、进程管理、自动更新、账号系统、复杂分类、排序或插件系统。

## Windows 注意事项

- `.lnk` 快捷方式应使用 `pywin32` / COM 解析。
- 管理员启动应使用 Windows `ShellExecuteW` 的 `runas`。
- 文档类目标应尊重 Windows 默认关联程序。
- 不要假设快捷方式目标一定存在，也不要假设所有文件类型都有默认打开程序。

## 验证要求

- 修改后至少检查程序能启动，搜索和来源筛选可用，列表可滚动，启动失败有明确提示。
- 如修改打包相关内容，确认 PyInstaller spec 仍包含 `pythoncom`、`pywintypes`、`win32com.client`。
