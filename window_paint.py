"""Keep the Windows restore surface transparent until Tk finishes painting."""
import tkinter as tk

RESIZE_SETTLE_MS = 120
RESIZE_ENDED = '<<RuyiResizeEnded>>'


def window_is_resizing(widget):
    root = widget.winfo_toplevel()
    active = getattr(root, '_ruyi_resize_active', False)
    if active and root._ruyi_resize_pending is None:
        root._ruyi_resize_pending = root.after(8, _poll_resize_end, root)
    return active


def _poll_resize_end(root):
    root._ruyi_resize_pending = None
    if not window_is_resizing(root):
        root.event_generate(RESIZE_ENDED, when='tail')


def track_window_resize(root):
    if root.tk.call('tk', 'windowingsystem') != 'win32':
        return
    from ctypes import POINTER, WinDLL, WINFUNCTYPE, byref, c_size_t, c_ssize_t
    from ctypes.wintypes import BOOL, DWORD, HBRUSH, HDC, HWND, INT, LPARAM, RECT, UINT, WPARAM

    user32 = WinDLL('user32')
    user32.GetAncestor.argtypes = (HWND, UINT)
    user32.GetAncestor.restype = HWND
    user32.GetForegroundWindow.restype = HWND
    user32.GetWindowRect.argtypes = (HWND, POINTER(RECT))
    user32.GetWindowRect.restype = BOOL
    user32.GetClientRect.argtypes = (HWND, POINTER(RECT))
    user32.GetClientRect.restype = BOOL
    user32.FillRect.argtypes = (HDC, POINTER(RECT), HBRUSH)
    user32.FillRect.restype = INT
    gdi = WinDLL('gdi32')
    gdi.GetStockObject.argtypes = (INT,)
    gdi.GetStockObject.restype = HBRUSH
    gdi.SetDCBrushColor.argtypes = (HDC, DWORD)
    gdi.SetDCBrushColor.restype = DWORD
    brush = gdi.GetStockObject(18)  # DC_BRUSH
    client = RECT()
    controls = WinDLL('comctl32')
    subclass = WINFUNCTYPE(c_ssize_t, HWND, UINT, WPARAM, LPARAM, c_size_t, c_size_t)
    controls.SetWindowSubclass.argtypes = (HWND, subclass, c_size_t, c_size_t)
    controls.SetWindowSubclass.restype = BOOL
    controls.RemoveWindowSubclass.argtypes = (HWND, subclass, c_size_t)
    controls.RemoveWindowSubclass.restype = BOOL
    controls.DefSubclassProc.argtypes = (HWND, UINT, WPARAM, LPARAM)
    controls.DefSubclassProc.restype = c_ssize_t
    root._ruyi_resize_active = False
    root._ruyi_resize_pending = None
    root._ruyi_resize_snapshot = None
    captured = False

    def capture_page(hwnd):
        page = root._ruyi_snapshot_hwnd
        if not page or user32.GetForegroundWindow() != hwnd:
            return
        # Capture only the visible client area belonging to this app, before
        # Windows clips it to the new size. Keep the result in memory.
        bounds = RECT()
        if not user32.GetWindowRect(page, byref(bounds)):
            return
        x, y, width, height = (user32.GetSystemMetrics(i) for i in (76, 77, 78, 79))
        if not (x <= bounds.left < bounds.right <= x + width
                and y <= bounds.top < bounds.bottom <= y + height
                and (bounds.right-bounds.left)*(bounds.bottom-bounds.top) <= 8_500_000):
            return
        try:
            from PIL import ImageGrab
            image = ImageGrab.grab(window=page)
            if image.getbbox():
                root._ruyi_resize_snapshot = image
            else:
                image.close()
        except (OSError, ValueError):
            pass

    root._ruyi_capture_resize_page = capture_page

    @subclass
    def window_proc(hwnd, message, wparam, lparam, identifier, _data):
        nonlocal captured
        if message == 0x0014 and wparam and brush:  # WM_ERASEBKGND
            # Tk's outer window leaves erasure to its children. Paint any
            # exposed area with the current theme while they catch up.
            if user32.GetClientRect(hwnd, byref(client)):
                previous = gdi.SetDCBrushColor(wparam, root._ruyi_background)
                painted = user32.FillRect(wparam, byref(client), brush)
                gdi.SetDCBrushColor(wparam, previous)
                if painted:
                    return 1
        elif message == 0x0231:  # WM_ENTERSIZEMOVE
            root._ruyi_resize_active = True
            captured = False
            if root._ruyi_resize_snapshot is not None:
                root._ruyi_resize_snapshot.close()
                root._ruyi_resize_snapshot = None
        elif message == 0x0214 and not captured:  # WM_SIZING
            captured = True
            capture_page(hwnd)
        elif message == 0x0232:  # WM_EXITSIZEMOVE
            root._ruyi_resize_active = False
        elif message == 0x0082:  # WM_NCDESTROY
            controls.RemoveWindowSubclass(hwnd, window_proc, identifier)
        return controls.DefSubclassProc(hwnd, message, wparam, lparam)

    def install(event):
        if event.widget is root:
            hwnd = user32.GetAncestor(root.winfo_id(), 2)
            if hwnd:
                controls.SetWindowSubclass(hwnd, window_proc, 1, 0)

    root.bind('<Map>', install, add='+')
    def cancel_poll(event):
        if event.widget is root and root._ruyi_resize_pending is not None:
            root.after_cancel(root._ruyi_resize_pending)
            root._ruyi_resize_pending = None
    root.bind('<Destroy>', cancel_poll, add='+')


def set_window_layered(root, enabled):
    from ctypes import WinDLL
    from ctypes.wintypes import BOOL, HWND, INT, LONG, LPVOID, UINT

    user32 = WinDLL('user32')
    user32.GetAncestor.argtypes = (HWND, UINT)
    user32.GetAncestor.restype = HWND
    user32.GetWindowLongW.argtypes = (HWND, INT)
    user32.GetWindowLongW.restype = LONG
    user32.SetWindowLongW.argtypes = (HWND, INT, LONG)
    user32.SetWindowLongW.restype = LONG
    user32.RedrawWindow.argtypes = (HWND, LPVOID, LPVOID, UINT)
    user32.RedrawWindow.restype = BOOL
    hwnd = user32.GetAncestor(root.winfo_id(), 2)  # GA_ROOT, Tk's outer window
    ex_style, layered = -20, 0x80000  # GWL_EXSTYLE, WS_EX_LAYERED
    style = user32.GetWindowLongW(hwnd, ex_style)
    updated = (style | layered) if enabled else (style & ~layered)
    if not hwnd or updated == style:
        return
    user32.SetWindowLongW(hwnd, ex_style, updated)
    if not enabled:
        # RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_FRAME
        user32.RedrawWindow(hwnd, None, None, 0x485)


class WindowPaintGate:
    def __init__(self, root):
        self.root = root
        self.pending = None
        self.windows = root.tk.call('tk', 'windowingsystem') == 'win32'
        if self.windows:
            root.attributes('-alpha', 0.0)
        root.bind('<Unmap>', self.unmapped, add='+')
        root.bind('<Map>', self.mapped, add='+')
        root.bind('<Destroy>', self.destroyed, add='+')

    def unmapped(self, event):
        if event.widget is self.root and self.windows and self.root.state() in ('iconic', 'withdrawn'):
            # Do this while minimized, before the taskbar restores the HWND.
            # Tk remembers its layered style even after we remove it natively.
            set_window_layered(self.root, True)
            self.root.attributes('-alpha', 0.0)

    def mapped(self, event):
        if event.widget is not self.root:
            return
        if self.pending is not None:
            self.root.after_cancel(self.pending)
        self.pending = self.root.after_idle(self.painted)

    def painted(self):
        self.pending = None
        if self.windows:
            # Idle work submits child paints; allow the compositor one more
            # frame before revealing them, rather than exposing black HWNDs.
            self.pending = self.root.after(16, self.reveal)

    def reveal(self):
        self.pending = None
        if self.root.state() not in ('normal', 'zoomed'):
            return
        self.root.attributes('-alpha', 1.0)
        set_window_layered(self.root, False)

    def destroyed(self, event):
        if event.widget is self.root and self.pending is not None:
            try:
                self.root.after_cancel(self.pending)
            except tk.TclError:
                pass
            self.pending = None
