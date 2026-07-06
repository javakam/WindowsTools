import ctypes
import os
import threading
from dataclasses import dataclass
from pathlib import Path
from tkinter import BOTH, END, LEFT, VERTICAL, X, Y, StringVar, Tk, messagebox
from tkinter import ttk

import pythoncom
import win32com.client


APP_TITLE = "WindowsTools20260702V1"
SUPPORTED_EXTENSIONS = {".lnk", ".exe", ".appref-ms"}
FILTER_ALL = "全部"
SOURCE_START_MENU = "开始菜单"
SOURCE_DESKTOP = "桌面"


@dataclass(frozen=True)
class AppEntry:
    name: str
    path: str
    source: str
    source_group: str


@dataclass(frozen=True)
class LaunchTarget:
    file_path: str
    arguments: str
    working_directory: str | None


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
    return SOURCE_DESKTOP


def scan_apps() -> list[AppEntry]:
    entries: dict[str, AppEntry] = {}

    for source, root in get_scan_roots():
        try:
            for item in root.rglob("*"):
                if not item.is_file() or item.suffix.lower() not in SUPPORTED_EXTENSIONS:
                    continue

                name = display_name_from_path(item)
                key = f"{name.lower()}|{str(item).lower()}"
                entries[key] = AppEntry(
                    name=name,
                    path=str(item),
                    source=source,
                    source_group=source_group_from_source(source),
                )
        except OSError:
            continue

    return sorted(entries.values(), key=lambda entry: entry.name.lower())


def resolve_shortcut(path: str) -> tuple[LaunchTarget | None, str | None]:
    shortcut_path = Path(path)
    if not shortcut_path.exists():
        return None, "快捷方式不存在，可能已经被移动或删除。"

    suffix = shortcut_path.suffix.lower()
    if suffix != ".lnk":
        if suffix == ".exe":
            return LaunchTarget(str(shortcut_path), "", str(shortcut_path.parent)), None
        return None, f"当前只支持启动 exe 或解析 lnk 快捷方式，文件类型为：{suffix or '未知'}"

    shell = None
    shortcut = None
    pythoncom.CoInitialize()
    try:
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
        pythoncom.CoUninitialize()

    if not target_path:
        return None, "快捷方式没有可启动的目标程序。"

    target = Path(target_path)
    if not target.exists():
        return None, f"快捷方式目标不存在：{target_path}"
    if target.suffix.lower() != ".exe":
        return None, f"当前只支持以管理员权限启动 exe，目标类型为：{target.suffix or '未知'}"

    cwd = working_directory if working_directory and Path(working_directory).exists() else str(target.parent)
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
        self.root.geometry("900x540")
        self.root.minsize(760, 420)

        self.query = StringVar()
        self.source_filter = StringVar(value=FILTER_ALL)
        self.status = StringVar(value="正在扫描当前 Windows 开始菜单和桌面软件...")
        self.apps: list[AppEntry] = []
        self.filtered_apps: list[AppEntry] = []

        self._build_ui()
        self.refresh()

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill=BOTH, expand=True)

        top = ttk.Frame(main)
        top.pack(fill=X)

        ttk.Label(top, text="搜索").pack(side=LEFT)
        search = ttk.Entry(top, textvariable=self.query)
        search.pack(side=LEFT, fill=X, expand=True, padx=(8, 10))
        search.bind("<KeyRelease>", lambda _event: self.apply_filter())
        search.bind("<Escape>", lambda _event: self.clear_search())

        ttk.Label(top, text="来源").pack(side=LEFT)
        self.source_box = ttk.Combobox(
            top,
            textvariable=self.source_filter,
            values=(FILTER_ALL, SOURCE_START_MENU, SOURCE_DESKTOP),
            state="readonly",
            width=10,
        )
        self.source_box.pack(side=LEFT, padx=(8, 8))
        self.source_box.bind("<<ComboboxSelected>>", lambda _event: self.apply_filter())

        ttk.Button(top, text="刷新", command=self.refresh).pack(side=LEFT, padx=(0, 8))
        ttk.Button(top, text="管理员启动", command=self.launch_selected).pack(side=LEFT)

        body = ttk.Frame(main)
        body.pack(fill=BOTH, expand=True, pady=(12, 8))
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        columns = ("name", "source", "path")
        self.tree = ttk.Treeview(body, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("name", text="软件")
        self.tree.heading("source", text="来源")
        self.tree.heading("path", text="路径")
        self.tree.column("name", width=240, anchor="w")
        self.tree.column("source", width=120, anchor="w")
        self.tree.column("path", width=520, anchor="w")
        self.tree.bind("<Double-1>", lambda _event: self.launch_selected())
        self.tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(body, orient=VERTICAL, command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

        bottom = ttk.Frame(main)
        bottom.pack(fill=X)

        ttk.Label(bottom, textvariable=self.status).pack(side=LEFT, fill=X, expand=True)

    def refresh(self) -> None:
        self.status.set("正在扫描当前 Windows 开始菜单和桌面软件...")
        self.tree.delete(*self.tree.get_children())
        threading.Thread(target=self._refresh_worker, daemon=True).start()

    def _refresh_worker(self) -> None:
        apps = scan_apps()
        self.root.after(0, lambda: self._finish_refresh(apps))

    def _finish_refresh(self, apps: list[AppEntry]) -> None:
        self.apps = apps
        self.apply_filter()

    def apply_filter(self) -> None:
        keywords = [part for part in self.query.get().strip().lower().split() if part]
        source = self.source_filter.get()
        self.filtered_apps = []

        for app in self.apps:
            if source != FILTER_ALL and app.source_group != source:
                continue
            if keywords and not all(self.match_app(app, keyword) for keyword in keywords):
                continue
            self.filtered_apps.append(app)

        self.tree.delete(*self.tree.get_children())
        for index, app in enumerate(self.filtered_apps):
            self.tree.insert("", END, iid=str(index), values=(app.name, app.source, app.path))

        self.status.set(
            f"共 {len(self.apps)} 项，当前显示 {len(self.filtered_apps)} 项，来源：{source}。双击也可以启动。"
        )

    def match_app(self, app: AppEntry, keyword: str) -> bool:
        return (
            keyword in app.name.lower()
            or keyword in app.source.lower()
            or keyword in app.path.lower()
        )

    def clear_search(self) -> None:
        if self.query.get():
            self.query.set("")
            self.apply_filter()

    def launch_selected(self) -> None:
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo(APP_TITLE, "请先选择一个软件。")
            return

        index = int(selection[0])
        app = self.filtered_apps[index]
        success, message = launch_as_admin(app)
        self.status.set(f"{app.name}: {message}")
        if not success:
            messagebox.showerror(APP_TITLE, message)

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    WindowsToolsApp().run()
