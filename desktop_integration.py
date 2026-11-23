"""Local preferences and explicit, user-invoked Windows shortcut creation."""
import base64
import ctypes
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import uuid


def enable_dpi_awareness():
    if os.name == 'nt':
        try:
            ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
        except (AttributeError, OSError):
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(2)
            except (AttributeError, OSError):
                pass


def window_scale(root):
    try:
        from ctypes import wintypes
        get = ctypes.windll.user32.GetDpiForWindow
        get.argtypes = [wintypes.HWND]
        get.restype = ctypes.c_uint
        dpi = get(root.winfo_id()) or 96
        root.tk.call('tk', 'scaling', dpi / 72)
        return dpi / 96
    except (AttributeError, OSError):
        return root.winfo_fpixels('1i') / 96


def install_window_icons(root, ico_path):
    """Load native small/large icons at the exact monitor DPI; keep their handles alive."""
    if os.name != 'nt':
        return
    from ctypes import wintypes
    user = ctypes.windll.user32
    user.GetAncestor.argtypes = [wintypes.HWND, ctypes.c_uint]
    user.GetAncestor.restype = wintypes.HWND
    user.GetDpiForWindow.argtypes = [wintypes.HWND]
    user.GetDpiForWindow.restype = ctypes.c_uint
    user.LoadImageW.argtypes = [wintypes.HINSTANCE, wintypes.LPCWSTR, ctypes.c_uint, ctypes.c_int, ctypes.c_int, ctypes.c_uint]
    user.LoadImageW.restype = wintypes.HANDLE
    user.SendMessageW.argtypes = [wintypes.HWND, ctypes.c_uint, ctypes.c_size_t, ctypes.c_ssize_t]
    user.SendMessageW.restype = ctypes.c_ssize_t
    user.DestroyIcon.argtypes = [wintypes.HANDLE]
    user.GetSystemMetricsForDpi.argtypes = [ctypes.c_int, ctypes.c_uint]
    user.GetSystemMetricsForDpi.restype = ctypes.c_int
    state = {'dpi': None, 'handles': []}
    def refresh(event=None):
        if event is not None and event.widget is not root:
            return
        hwnd = user.GetAncestor(root.winfo_id(), 2)
        dpi = user.GetDpiForWindow(hwnd) or 96
        if state['dpi'] == dpi:
            return
        handles = []
        for kind, metric in ((0,49),(1,11)):
            size = user.GetSystemMetricsForDpi(metric, dpi)
            handle = user.LoadImageW(None, str(ico_path), 1, size, size, 0x10)
            if not handle:
                for loaded in handles:user.DestroyIcon(loaded)
                raise OSError('无法载入如意窗口图标。')
            handles.append(handle)
        for kind, handle in enumerate(handles):user.SendMessageW(hwnd, 0x80, kind, handle)
        for handle in state['handles']:user.DestroyIcon(handle)
        state.update(dpi=dpi, handles=handles)
    def close(event):
        if event.widget is root:
            for handle in state['handles']:user.DestroyIcon(handle)
            state['handles'] = []
    root.bind('<Configure>', refresh, add='+')
    root.bind('<Destroy>', close, add='+')
    root.after_idle(refresh)
    root._ruyi_native_icons = state


def settings_path():
    base = Path(os.environ.get('RUYI_CONFIG_DIR') or Path(os.environ.get('LOCALAPPDATA', Path.home())) / 'Ruyi')
    return base / 'settings.json'


def load_settings():
    defaults = {'appearance': '浅色', 'palette': '雾蓝', 'gradient': False, 'libreoffice': ''}
    try:
        saved = json.loads(settings_path().read_text(encoding='utf-8'))
        if isinstance(saved, dict):
            defaults.update({k: saved[k] for k in defaults if k in saved})
    except (OSError, ValueError):
        pass
    if defaults['appearance'] not in ('浅色', '深色', '跟随系统'):
        defaults['appearance'] = '浅色'
    if defaults['palette'] not in ('雾蓝', '石墨', '淡紫', '薄荷', '浅粉', '暮霞'):
        defaults['palette'] = '雾蓝'
    return defaults


def save_settings(values):
    destination = settings_path()
    destination.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix='settings-', suffix='.json', dir=destination.parent)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            json.dump(values, f, ensure_ascii=False, indent=2)
        os.replace(name, destination)
    finally:
        Path(name).unlink(missing_ok=True)


def system_dark():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Themes\Personalize') as key:
            return winreg.QueryValueEx(key, 'AppsUseLightTheme')[0] == 0
    except OSError:
        return False


def desktop_directory():
    if os.environ.get('RUYI_SHORTCUT_DESKTOP'):
        return Path(os.environ['RUYI_SHORTCUT_DESKTOP']).resolve(strict=True)
    from ctypes import wintypes
    class GUID(ctypes.Structure):
        _fields_ = [('Data1', wintypes.DWORD), ('Data2', wintypes.WORD), ('Data3', wintypes.WORD), ('Data4', ctypes.c_ubyte * 8)]
    guid = GUID.from_buffer_copy(uuid.UUID('B4BFCC3A-DB2C-424C-B029-7FE99A87C641').bytes_le)
    value = ctypes.c_wchar_p()
    get = ctypes.windll.shell32.SHGetKnownFolderPath
    get.argtypes = [ctypes.POINTER(GUID), wintypes.DWORD, wintypes.HANDLE, ctypes.POINTER(ctypes.c_wchar_p)]
    get.restype = ctypes.c_long
    if get(ctypes.byref(guid), 0, None, ctypes.byref(value)) != 0:
        raise OSError('无法定位桌面文件夹。')
    try:
        return Path(value.value)
    finally:
        ctypes.windll.ole32.CoTaskMemFree(ctypes.cast(value, ctypes.c_void_p))


def create_desktop_shortcut(executable=None):
    exe = Path(executable or sys.executable).resolve(strict=True)
    if exe.suffix.lower() != '.exe' or (executable is None and not getattr(sys, 'frozen', False)):
        raise ValueError('请使用已打包的如意程序创建桌面快捷方式。')
    destination = desktop_directory() / '如意.lnk'
    quote = lambda s: "'" + str(s).replace("'", "''") + "'"
    script = "$ErrorActionPreference='Stop'; $s=New-Object -ComObject WScript.Shell; "
    script += '$p=' + quote(destination) + '; $x=$s.CreateShortcut($p); '
    script += "if((Test-Path -LiteralPath $p) -and $x.TargetPath -and ([IO.Path]::GetFileName($x.TargetPath) -ne '如意.exe')) {throw '桌面已有同名的其他程序快捷方式，未覆盖。'}; "
    script += '$x.TargetPath=' + quote(exe) + '; $x.WorkingDirectory=' + quote(exe.parent) + '; '
    script += '$x.IconLocation=' + quote(str(exe) + ',0') + "; $x.Description='如意 · 本地文件工具箱'; $x.Save();"
    encoded = base64.b64encode(script.encode('utf-16le')).decode('ascii')
    shell = Path(os.environ.get('SystemRoot', r'C:\Windows')) / 'System32/WindowsPowerShell/v1.0/powershell.exe'
    subprocess.run([str(shell), '-NoProfile', '-NonInteractive', '-EncodedCommand', encoded],
                   check=True, timeout=20, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
    if not destination.is_file():
        raise OSError('快捷方式未成功保存。')
    notify_shell(destination)
    return destination


def notify_shell(path):
    if os.name != 'nt':
        return
    try:
        from ctypes import wintypes
        notify = ctypes.windll.shell32.SHChangeNotify
        notify.argtypes = [ctypes.c_long, ctypes.c_uint, wintypes.LPCWSTR, ctypes.c_void_p]
        # Path-specific create/update notifications, never change file associations.
        notify(0x00000002, 0x0005 | 0x2000, str(Path(path).resolve()), None)
        notify(0x00002000, 0x0005 | 0x2000, str(Path(path).resolve()), None)
    except (OSError, AttributeError):
        pass
