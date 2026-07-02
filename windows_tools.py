import ctypes
import os
import threading
from dataclasses import dataclass
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, VERTICAL, X, Y, StringVar, Tk, messagebox
from tkinter import ttk


APP_TITLE = "WindowsTools20260702V1"
SUPPORTED_EXTENSIONS = {".lnk", ".exe", ".appref-ms"}


@dataclass(frozen=True)
class AppEntry:
    name: str
    path: str
    source: str


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


def scan_apps() -> list[AppEntry]:
    entries: dict[str, AppEntry] = {}

    for source, root in get_scan_roots():
        try:
            for item in root.rglob("*"):
                if not item.is_file() or item.suffix.lower() not in SUPPORTED_EXTENSIONS:
                    continue

                name = display_name_from_path(item)
                key = f"{name.lower()}|{str(item).lower()}"
                entries[key] = AppEntry(name=name, path=str(item), source=source)
        except OSError:
            continue

    return sorted(entries.values(), key=lambda entry: entry.name.lower())


def launch_as_admin(path: str) -> tuple[bool, str]:
    file_path = Path(path)
    if not file_path.exists():
        return False, "文件不存在，可能已经被移动或卸载。"

    result = ctypes.windll.shell32.ShellExecuteW(
        None,
        "runas",
        str(file_path),
        None,
        None,
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
        search.pack(side=LEFT, fill=X, expand=True, padx=(8, 8))
        search.bind("<KeyRelease>", lambda _event: self.apply_filter())

        ttk.Button(top, text="刷新", command=self.refresh).pack(side=LEFT, padx=(0, 8))
        ttk.Button(top, text="管理员启动", command=self.launch_selected).pack(side=LEFT)

        body = ttk.Frame(main)
        body.pack(fill=BOTH, expand=True, pady=(12, 8))

        columns = ("name", "source", "path")
        self.tree = ttk.Treeview(body, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("name", text="软件")
        self.tree.heading("source", text="来源")
        self.tree.heading("path", text="路径")
        self.tree.column("name", width=240, anchor="w")
        self.tree.column("source", width=120, anchor="w")
        self.tree.column("path", width=520, anchor="w")
        self.tree.bind("<Double-1>", lambda _event: self.launch_selected())
        self.tree.pack(side=LEFT, fill=BOTH, expand=True)

        scrollbar = ttk.Scrollbar(body, orient=VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
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
        keyword = self.query.get().strip().lower()
        if keyword:
            self.filtered_apps = [
                app
                for app in self.apps
                if keyword in app.name.lower()
                or keyword in app.source.lower()
                or keyword in app.path.lower()
            ]
        else:
            self.filtered_apps = list(self.apps)

        self.tree.delete(*self.tree.get_children())
        for index, app in enumerate(self.filtered_apps):
            self.tree.insert("", END, iid=str(index), values=(app.name, app.source, app.path))

        self.status.set(f"共 {len(self.apps)} 项，当前显示 {len(self.filtered_apps)} 项。双击也可以启动。")

    def launch_selected(self) -> None:
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo(APP_TITLE, "请先选择一个软件。")
            return

        index = int(selection[0])
        app = self.filtered_apps[index]
        success, message = launch_as_admin(app.path)
        self.status.set(f"{app.name}: {message}")
        if not success:
            messagebox.showerror(APP_TITLE, message)

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    WindowsToolsApp().run()
