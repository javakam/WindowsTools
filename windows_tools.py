import ctypes
import json
import os
import re
import subprocess
import tempfile
import threading
from ctypes import wintypes
from dataclasses import dataclass
from pathlib import Path
from tkinter import BOTH, END, LEFT, Menu, VERTICAL, X, Y, StringVar, Tk, messagebox
from tkinter import font as tkfont
from tkinter import ttk
from xml.etree import ElementTree

import pythoncom
import win32com.client


APP_TITLE = "WindowsTools20260702V1"
SUPPORTED_EXTENSIONS = {".lnk", ".exe", ".appref-ms"}
FILTER_ALL = "全部"
SOURCE_START_MENU = "开始菜单"
SOURCE_DESKTOP = "桌面"
SOURCE_START_PINNED = "开始固定"
SOURCE_FILTERS = (FILTER_ALL, SOURCE_START_MENU, SOURCE_START_PINNED, SOURCE_DESKTOP)
SETTINGS_DIR_NAME = "WindowsTools"
SETTINGS_FILE_NAME = "settings.json"
WINDOW_WIDTH = 1280
WINDOW_HEIGHT = 720
MIN_WINDOW_WIDTH = 1080
MIN_WINDOW_HEIGHT = 560
NAME_COLUMN_WIDTH = 300
SOURCE_COLUMN_WIDTH = 130
PATH_MIN_COLUMN_WIDTH = 760
UI_FONT_SIZE = 10
TREE_ROW_HEIGHT = 28
TOP_LEFT_RESTORE_MARGIN = 80
WINDOW_GEOMETRY_PATTERN = re.compile(r"^(\d+)x(\d+)([+-]\d+)([+-]\d+)$")


@dataclass(frozen=True)
class AppEntry:
    name: str
    path: str
    source: str
    source_group: str
    drive: str = FILTER_ALL


@dataclass(frozen=True)
class LaunchTarget:
    file_path: str
    arguments: str
    working_directory: str | None


def get_settings_path() -> Path:
    appdata = os.environ.get("APPDATA")
    base_dir = Path(appdata) if appdata else Path.home()
    return base_dir / SETTINGS_DIR_NAME / SETTINGS_FILE_NAME


def load_saved_geometry() -> str | None:
    settings_path = get_settings_path()
    try:
        settings = json.loads(settings_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None

    geometry = settings.get("geometry")
    if isinstance(geometry, str) and WINDOW_GEOMETRY_PATTERN.match(geometry):
        return geometry
    return None


def save_window_geometry(geometry: str) -> None:
    if not WINDOW_GEOMETRY_PATTERN.match(geometry):
        return

    settings_path = get_settings_path()
    try:
        settings_path.parent.mkdir(parents=True, exist_ok=True)
        settings_path.write_text(json.dumps({"geometry": geometry}, ensure_ascii=False), encoding="utf-8")
    except OSError:
        return


def clamp_window_geometry(geometry: str, screen_width: int, screen_height: int) -> str:
    match = WINDOW_GEOMETRY_PATTERN.match(geometry)
    if not match:
        return geometry

    width = min(max(int(match.group(1)), MIN_WINDOW_WIDTH), screen_width)
    height = min(max(int(match.group(2)), MIN_WINDOW_HEIGHT), screen_height)
    x = max(min(int(match.group(3)), max(screen_width - 100, 0)), 0)
    y = max(min(int(match.group(4)), max(screen_height - 100, 0)), 0)
    return f"{width}x{height}+{x}+{y}"


def centered_window_geometry(width: int, height: int, screen_width: int, screen_height: int) -> str:
    x = max((screen_width - width) // 2, 0)
    y = max((screen_height - height) // 2, 0)
    return f"{width}x{height}+{x}+{y}"


def safe_restore_window_geometry(geometry: str, screen_width: int, screen_height: int) -> str:
    match = WINDOW_GEOMETRY_PATTERN.match(geometry)
    if not match:
        return geometry

    width = min(max(int(match.group(1)), MIN_WINDOW_WIDTH), screen_width)
    height = min(max(int(match.group(2)), MIN_WINDOW_HEIGHT), screen_height)
    x = int(match.group(3))
    y = int(match.group(4))
    if x <= TOP_LEFT_RESTORE_MARGIN and y <= TOP_LEFT_RESTORE_MARGIN:
        return centered_window_geometry(width, height, screen_width, screen_height)

    return clamp_window_geometry(f"{width}x{height}{match.group(3)}{match.group(4)}", screen_width, screen_height)


def drive_filter_from_path(path: str) -> str:
    drive = os.path.splitdrive(path)[0].upper()
    if not drive:
        return FILTER_ALL
    return f"{drive[0]}盘"


def drive_filter_for_item(path: Path, shell=None) -> str:
    if path.suffix.lower() != ".lnk" or shell is None:
        return drive_filter_from_path(str(path))

    shortcut = None
    try:
        shortcut = shell.CreateShortcut(str(path))
        target_path = os.path.expandvars(shortcut.TargetPath or "").strip()
    except Exception:
        return drive_filter_from_path(str(path))
    finally:
        shortcut = None

    return drive_filter_from_path(target_path or str(path))


def cached_drive_filter_for_item(path: Path, shell, cache: dict[str, str]) -> str:
    key = str(path).lower()
    if key not in cache:
        cache[key] = drive_filter_for_item(path, shell)
    return cache[key]


def get_scan_roots() -> list[tuple[str, Path]]:
    user_profile = Path(os.environ.get("USERPROFILE", ""))
    appdata = Path(os.environ.get("APPDATA", ""))
    program_data = Path(os.environ.get("PROGRAMDATA", r"C:\ProgramData"))

    roots = [
        ("系统开始菜单", program_data / "Microsoft" / "Windows" / "Start Menu" / "Programs"),
        ("用户开始菜单", appdata / "Microsoft" / "Windows" / "Start Menu" / "Programs"),
        ("公共桌面", Path(r"C:\Users\Public\Desktop")),
        ("用户桌面", user_profile / "Desktop"),
    ]
    return [(label, path) for label, path in roots if path.exists()]


def display_name_from_path(path: Path) -> str:
    name = path.stem.strip()
    for suffix in (" - 快捷方式", " 快捷方式", " Shortcut"):
        if name.endswith(suffix):
            name = name[: -len(suffix)].strip()
    return name or path.name


def source_group_from_source(source: str) -> str:
    if source in ("系统开始菜单", "用户开始菜单"):
        return SOURCE_START_MENU
    if source == SOURCE_START_PINNED:
        return SOURCE_START_PINNED
    return SOURCE_DESKTOP


def powershell_quote(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def export_start_layout_text() -> str | None:
    creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0

    with tempfile.TemporaryDirectory(prefix="windowstools_start_layout_") as temp_dir:
        for file_name in ("layout.json", "layout.xml"):
            layout_path = Path(temp_dir) / file_name
            command = [
                "powershell.exe",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                f"Export-StartLayout -Path {powershell_quote(str(layout_path))}",
            ]

            try:
                result = subprocess.run(
                    command,
                    capture_output=True,
                    creationflags=creationflags,
                    encoding="utf-8",
                    errors="replace",
                    text=True,
                    timeout=10,
                    check=False,
                )
            except (OSError, subprocess.TimeoutExpired):
                continue

            if result.returncode == 0 and layout_path.exists():
                try:
                    return layout_path.read_text(encoding="utf-8-sig")
                except OSError:
                    return None

    return None


def desktop_links_from_start_layout(layout_text: str) -> list[str]:
    text = layout_text.strip()
    if not text:
        return []

    if text.startswith("{"):
        return desktop_links_from_start_layout_json(text)
    return desktop_links_from_start_layout_xml(text)


def desktop_links_from_start_layout_json(layout_text: str) -> list[str]:
    try:
        layout = json.loads(layout_text)
    except json.JSONDecodeError:
        return []

    links: list[str] = []
    for item in layout.get("pinnedList", []):
        if not isinstance(item, dict):
            continue
        link = item.get("desktopAppLink")
        if isinstance(link, str) and link.strip():
            links.append(link.strip())
    return links


def desktop_links_from_start_layout_xml(layout_text: str) -> list[str]:
    try:
        root = ElementTree.fromstring(layout_text)
    except ElementTree.ParseError:
        return []

    links: list[str] = []
    for element in root.iter():
        for name, value in element.attrib.items():
            if name.endswith("DesktopApplicationLinkPath") and value.strip():
                links.append(value.strip())
    return links


def scan_start_pinned_apps(shell=None, drive_cache: dict[str, str] | None = None) -> list[AppEntry]:
    layout_text = export_start_layout_text()
    if not layout_text:
        return []

    entries: dict[str, AppEntry] = {}
    for link in desktop_links_from_start_layout(layout_text):
        path = Path(os.path.expandvars(link))
        if not path.exists() or not path.is_file() or path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue

        name = display_name_from_path(path)
        key = f"{name.lower()}|{str(path).lower()}|{SOURCE_START_PINNED}"
        entries[key] = AppEntry(
            name=name,
            path=str(path),
            source=SOURCE_START_PINNED,
            source_group=SOURCE_START_PINNED,
            drive=(
                cached_drive_filter_for_item(path, shell, drive_cache)
                if drive_cache is not None
                else drive_filter_for_item(path, shell)
            ),
        )

    return list(entries.values())


def scan_apps() -> list[AppEntry]:
    entries: dict[str, AppEntry] = {}
    drive_cache: dict[str, str] = {}
    shell = None
    co_initialized = False

    try:
        pythoncom.CoInitialize()
        co_initialized = True
        shell = win32com.client.Dispatch("WScript.Shell")
    except Exception:
        shell = None

    try:
        for source, root in get_scan_roots():
            try:
                for item in root.rglob("*"):
                    if item.suffix.lower() not in SUPPORTED_EXTENSIONS or not item.is_file():
                        continue

                    name = display_name_from_path(item)
                    key = f"{name.lower()}|{str(item).lower()}|{source}"
                    entries[key] = AppEntry(
                        name=name,
                        path=str(item),
                        source=source,
                        source_group=source_group_from_source(source),
                        drive=cached_drive_filter_for_item(item, shell, drive_cache),
                    )
            except OSError:
                continue

        for app in scan_start_pinned_apps(shell, drive_cache):
            key = f"{app.name.lower()}|{app.path.lower()}|{app.source}"
            entries[key] = app
    finally:
        shell = None
        if co_initialized:
            pythoncom.CoUninitialize()

    return sorted(entries.values(), key=lambda entry: entry.name.lower())


def associated_executable_for_path(path: str) -> tuple[str | None, str | None]:
    suffix = Path(path).suffix
    if not suffix:
        return None, "文件没有扩展名，无法查找默认打开程序。"

    length = wintypes.DWORD(0)
    result = ctypes.windll.shlwapi.AssocQueryStringW(
        0,
        2,
        suffix,
        None,
        None,
        ctypes.byref(length),
    )
    if result not in (0, 1) or length.value == 0:
        return None, f"没有找到 {suffix} 的默认打开程序。"

    buffer = ctypes.create_unicode_buffer(length.value)
    result = ctypes.windll.shlwapi.AssocQueryStringW(
        0,
        2,
        suffix,
        None,
        buffer,
        ctypes.byref(length),
    )
    if result != 0 or not buffer.value.strip():
        return None, f"无法读取 {suffix} 的默认打开程序。"

    executable = os.path.expandvars(buffer.value.strip())
    if not Path(executable).exists():
        return None, f"默认打开程序不存在：{executable}"

    return executable, None


def resolve_document_target(
    document_path: str,
    arguments: str = "",
    working_directory: str | None = None,
) -> tuple[LaunchTarget | None, str | None]:
    executable, error = associated_executable_for_path(document_path)
    if error:
        return None, error

    document_argument = subprocess.list2cmdline([document_path])
    final_arguments = document_argument if not arguments else f"{document_argument} {arguments}"
    cwd = working_directory if working_directory and Path(working_directory).exists() else str(Path(document_path).parent)
    return LaunchTarget(executable, final_arguments, cwd), None


def resolve_shortcut(path: str) -> tuple[LaunchTarget | None, str | None]:
    shortcut_path = Path(path)
    if not shortcut_path.exists():
        return None, "快捷方式不存在，可能已经被移动或删除。"

    suffix = shortcut_path.suffix.lower()
    if suffix != ".lnk":
        if suffix == ".exe":
            return LaunchTarget(str(shortcut_path), "", str(shortcut_path.parent)), None
        return resolve_document_target(str(shortcut_path))

    shell = None
    shortcut = None
    co_initialized = False
    try:
        pythoncom.CoInitialize()
        co_initialized = True
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(str(shortcut_path))
        target_path = os.path.expandvars(shortcut.TargetPath or "").strip()
        arguments = os.path.expandvars(shortcut.Arguments or "").strip()
        working_directory = os.path.expandvars(shortcut.WorkingDirectory or "").strip()
    except Exception as exc:
        return None, f"无法解析快捷方式：{exc}"
    finally:
        shortcut = None
        shell = None
        if co_initialized:
            pythoncom.CoUninitialize()

    if not target_path:
        return None, "快捷方式没有可启动的目标程序。"

    target = Path(target_path)
    if not target.exists():
        return None, f"快捷方式目标不存在：{target_path}"

    cwd = working_directory if working_directory and Path(working_directory).exists() else str(target.parent)
    if target.suffix.lower() != ".exe":
        return resolve_document_target(str(target), arguments, cwd)

    return LaunchTarget(str(target), arguments, cwd), None


def launch_as_admin(app: AppEntry) -> tuple[bool, str]:
    target, error = resolve_shortcut(app.path)
    if error:
        return False, f"{error}\n快捷方式路径：{app.path}"

    if target is None:
        return False, f"无法解析启动目标。\n快捷方式路径：{app.path}"

    file_path = Path(target.file_path)
    if not file_path.exists():
        return False, "文件不存在，可能已经被移动或卸载。"

    result = ctypes.windll.shell32.ShellExecuteW(
        None,
        "runas",
        str(file_path),
        target.arguments or None,
        target.working_directory,
        1,
    )
    if result > 32:
        return True, "已触发管理员权限启动，请在 UAC 弹窗中确认。"

    errors = {
        2: "找不到指定文件。",
        3: "找不到指定路径。",
        5: "权限被拒绝或取消了 UAC。",
        26: "共享冲突，文件可能正在被占用。",
        27: "文件关联不完整。",
        31: "没有可用于启动该文件类型的程序。",
    }
    return False, errors.get(result, f"启动失败，ShellExecute 返回码：{result}")


class WindowsToolsApp:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title(APP_TITLE)
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.root.minsize(MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT)
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self._configure_fonts()

        self.query = StringVar()
        self.source_filter = StringVar(value=FILTER_ALL)
        self.drive_filter = StringVar(value=FILTER_ALL)
        self.status = StringVar(value="正在扫描当前 Windows 开始菜单、桌面和开始固定项...")
        self.apps: list[AppEntry] = []
        self.filtered_apps: list[AppEntry] = []
        self.drive_filters: tuple[str, ...] = (FILTER_ALL,)
        self.is_closing = False

        self._build_ui()
        self._restore_or_center_window()
        self.refresh()

    def _configure_fonts(self) -> None:
        for font_name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            try:
                named_font = tkfont.nametofont(font_name)
            except Exception:
                continue

            named_font.configure(size=UI_FONT_SIZE)
            if font_name == "TkHeadingFont":
                named_font.configure(weight="bold")

    def _center_window(self) -> None:
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        self.root.geometry(centered_window_geometry(WINDOW_WIDTH, WINDOW_HEIGHT, screen_width, screen_height))

    def _restore_or_center_window(self) -> None:
        geometry = load_saved_geometry()
        if not geometry:
            self._center_window()
            return

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        self.root.geometry(safe_restore_window_geometry(geometry, screen_width, screen_height))

    def _build_ui(self) -> None:
        style = ttk.Style(self.root)
        style.configure("SourceTitle.TLabel", font=("Microsoft YaHei UI", UI_FONT_SIZE, "bold"), foreground="#1f5fa8")
        style.configure("DriveTitle.TLabel", font=("Microsoft YaHei UI", UI_FONT_SIZE, "bold"), foreground="#5f6470")
        style.configure("Treeview", rowheight=TREE_ROW_HEIGHT)
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", UI_FONT_SIZE, "bold"))
        style.map("Treeview", background=[("selected", "#cfe8ff")], foreground=[("selected", "#0f172a")])

        main = ttk.Frame(self.root, padding=12)
        main.pack(fill=BOTH, expand=True)

        top = ttk.Frame(main, padding=(0, 0, 0, 4))
        top.pack(fill=X)
        top.bind("<Configure>", self.update_toolbar_wrap)
        self.toolbar = top
        self.toolbar_drive_wrapped = False

        self.search_label = ttk.Label(top, text="搜索")
        self.search_label.grid(row=0, column=0, sticky="w", padx=(0, 0), pady=(0, 4))
        self.search_entry = ttk.Entry(top, textvariable=self.query)
        search = self.search_entry
        search.configure(width=32)
        search.grid(row=0, column=1, sticky="w", padx=(8, 18), pady=(0, 4))
        search.bind("<KeyRelease>", lambda _event: self.apply_filter())
        search.bind("<Escape>", lambda _event: self.clear_search())

        self.source_group = ttk.Frame(top)
        self.source_group.grid(row=0, column=2, sticky="w", padx=(0, 14), pady=(0, 4))
        ttk.Label(self.source_group, text="来源", style="SourceTitle.TLabel").pack(side=LEFT, padx=(0, 10))
        for source in SOURCE_FILTERS:
            ttk.Radiobutton(
                self.source_group,
                text=source,
                value=source,
                variable=self.source_filter,
                command=self.apply_filter,
            ).pack(side=LEFT, padx=(0, 10))

        self.filter_separator = ttk.Separator(top, orient=VERTICAL)
        self.filter_separator.grid(row=0, column=3, sticky="ns", padx=(0, 14), pady=(2, 6))

        self.drive_group = ttk.Frame(top)
        self.drive_group.grid(row=0, column=4, sticky="w", padx=(0, 0), pady=(0, 4))
        ttk.Label(self.drive_group, text="盘符", style="DriveTitle.TLabel").pack(side=LEFT, padx=(0, 10))
        self.drive_filter_frame = ttk.Frame(self.drive_group)
        self.drive_filter_frame.pack(side=LEFT)

        body = ttk.Frame(main)
        body.pack(fill=BOTH, expand=True, pady=(12, 8))
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        columns = ("name", "source", "path")
        self.tree = ttk.Treeview(body, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("name", text="软件")
        self.tree.heading("source", text="来源")
        self.tree.heading("path", text="路径")
        self.tree.column("name", width=NAME_COLUMN_WIDTH, minwidth=180, anchor="w", stretch=False)
        self.tree.column("source", width=SOURCE_COLUMN_WIDTH, minwidth=100, anchor="w", stretch=False)
        self.tree.column("path", width=PATH_MIN_COLUMN_WIDTH, minwidth=560, anchor="w", stretch=True)
        self.tree.bind("<Double-1>", lambda _event: self.launch_selected())
        self.tree.bind("<Configure>", lambda _event: self.update_column_widths())
        self.tree.bind("<Button-3>", self.show_context_menu)
        self.tree.grid(row=0, column=0, sticky="nsew")

        self.context_menu = Menu(self.root, tearoff=False)
        self.context_menu.add_command(label="管理员启动", command=self.launch_selected)
        self.context_menu.add_command(label="打开所在位置", command=self.open_selected_location)
        self.context_menu.add_command(label="复制路径", command=self.copy_selected_path)

        scrollbar = ttk.Scrollbar(body, orient=VERTICAL, command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

        bottom = ttk.Frame(main)
        bottom.pack(fill=X)

        ttk.Label(bottom, textvariable=self.status).pack(side=LEFT, fill=X, expand=True)

    def toolbar_required_width(self) -> int:
        widgets = (self.search_label, self.search_entry, self.source_group, self.filter_separator, self.drive_group)
        return sum(widget.winfo_reqwidth() for widget in widgets) + 90

    def update_toolbar_wrap(self, _event=None) -> None:
        if not hasattr(self, "drive_group"):
            return

        available_width = self.toolbar.winfo_width()
        if available_width <= 1:
            return

        should_wrap = available_width < self.toolbar_required_width()
        if should_wrap == self.toolbar_drive_wrapped:
            return

        self.toolbar_drive_wrapped = should_wrap
        if should_wrap:
            self.filter_separator.grid_remove()
            self.drive_group.grid_configure(row=1, column=0, columnspan=5, sticky="w", padx=(0, 0), pady=(4, 4))
        else:
            self.filter_separator.grid()
            self.drive_group.grid_configure(row=0, column=4, columnspan=1, sticky="w", padx=(0, 0), pady=(0, 4))

    def refresh(self) -> None:
        self.status.set("正在扫描当前 Windows 开始菜单、桌面和开始固定项...")
        self.tree.delete(*self.tree.get_children())
        threading.Thread(target=self._refresh_worker, daemon=True).start()

    def _refresh_worker(self) -> None:
        apps = scan_apps()
        if self.is_closing:
            return

        try:
            self.root.after(0, lambda: self._finish_refresh(apps))
        except RuntimeError:
            return

    def _finish_refresh(self, apps: list[AppEntry]) -> None:
        if self.is_closing:
            return

        self.apps = apps
        self.update_drive_filters()
        self.apply_filter()

    def update_drive_filters(self) -> None:
        drives = sorted(
            {
                drive
                for app in self.apps
                for drive in (app.drive,)
                if drive != FILTER_ALL
            }
        )
        self.drive_filters = (FILTER_ALL, *drives)
        if self.drive_filter.get() not in self.drive_filters:
            self.drive_filter.set(FILTER_ALL)

        for child in self.drive_filter_frame.winfo_children():
            child.destroy()

        for drive in self.drive_filters:
            ttk.Radiobutton(
                self.drive_filter_frame,
                text=drive,
                value=drive,
                variable=self.drive_filter,
                command=self.apply_filter,
            ).pack(side=LEFT, padx=(0, 10))
        self.root.after_idle(self.update_toolbar_wrap)

    def apply_filter(self) -> None:
        keywords = [part for part in self.query.get().strip().lower().split() if part]
        source = self.source_filter.get()
        drive = self.drive_filter.get()
        self.filtered_apps = []

        for app in self.apps:
            if source != FILTER_ALL and app.source_group != source:
                continue
            if drive != FILTER_ALL and app.drive != drive:
                continue
            if keywords and not all(self.match_app(app, keyword) for keyword in keywords):
                continue
            self.filtered_apps.append(app)

        self.tree.delete(*self.tree.get_children())
        for index, app in enumerate(self.filtered_apps):
            self.tree.insert("", END, iid=str(index), values=(app.name, app.source, app.path))

        self.update_column_widths()
        self.status.set(
            f"共 {len(self.apps)} 项，当前显示 {len(self.filtered_apps)} 项，来源：{source}，盘符：{drive}。右键可操作，双击也可以启动。"
        )

    def update_column_widths(self) -> None:
        tree_width = self.tree.winfo_width()
        if tree_width <= 1:
            return

        path_width = max(tree_width - NAME_COLUMN_WIDTH - SOURCE_COLUMN_WIDTH - 8, PATH_MIN_COLUMN_WIDTH)

        self.tree.column("name", width=NAME_COLUMN_WIDTH)
        self.tree.column("source", width=SOURCE_COLUMN_WIDTH)
        self.tree.column("path", width=path_width)

    def show_context_menu(self, event) -> str:
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return "break"

        self.tree.selection_set(item_id)
        self.tree.focus(item_id)
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
        return "break"

    def selected_app(self) -> AppEntry | None:
        selection = self.tree.selection()
        if not selection:
            return None

        try:
            index = int(selection[0])
            return self.filtered_apps[index]
        except (ValueError, IndexError):
            return None

    def match_app(self, app: AppEntry, keyword: str) -> bool:
        return (
            keyword in app.name.lower()
            or keyword in app.source.lower()
            or keyword in app.drive.lower()
            or keyword in app.path.lower()
        )

    def clear_search(self) -> None:
        if self.query.get():
            self.query.set("")
            self.apply_filter()

    def launch_selected(self) -> None:
        app = self.selected_app()
        if app is None:
            messagebox.showinfo(APP_TITLE, "请先选择一个软件。")
            return

        success, message = launch_as_admin(app)
        self.status.set(f"{app.name}: {message}")
        if not success:
            messagebox.showerror(APP_TITLE, message)

    def open_selected_location(self) -> None:
        app = self.selected_app()
        if app is None:
            messagebox.showinfo(APP_TITLE, "请先选择一个软件。")
            return

        path = Path(app.path)
        if path.exists():
            command = ["explorer.exe", f"/select,{str(path)}"]
        elif path.parent.exists():
            command = ["explorer.exe", str(path.parent)]
        else:
            message = f"所在位置不存在：{path.parent}"
            self.status.set(f"{app.name}: {message}")
            messagebox.showerror(APP_TITLE, message)
            return

        try:
            subprocess.Popen(command)
        except OSError as exc:
            message = f"无法打开所在位置：{exc}"
            self.status.set(f"{app.name}: {message}")
            messagebox.showerror(APP_TITLE, message)
            return

        self.status.set(f"{app.name}: 已打开所在位置。")

    def copy_selected_path(self) -> None:
        app = self.selected_app()
        if app is None:
            messagebox.showinfo(APP_TITLE, "请先选择一个软件。")
            return

        self.root.clipboard_clear()
        self.root.clipboard_append(app.path)
        self.status.set(f"{app.name}: 已复制路径。")

    def run(self) -> None:
        self.root.mainloop()

    def close(self) -> None:
        self.is_closing = True
        if self.root.state() == "normal":
            save_window_geometry(self.root.geometry())
        self.root.destroy()


if __name__ == "__main__":
    WindowsToolsApp().run()
