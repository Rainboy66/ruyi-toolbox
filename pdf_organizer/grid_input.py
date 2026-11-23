"""Suspend IME only on the non-text page canvas; restore on focus-out/close."""
import os


class CanvasInput:
    def __init__(self,canvas):
        self.canvas=canvas;self.context=None;self.imm=None
        if os.name=='nt':
            import ctypes
            from ctypes import wintypes
            self.imm=ctypes.WinDLL('imm32',use_last_error=True)
            self.imm.ImmAssociateContext.argtypes=[wintypes.HWND,wintypes.HANDLE]
            self.imm.ImmAssociateContext.restype=wintypes.HANDLE
            canvas.bind('<FocusIn>',self.enter,add='+');canvas.bind('<FocusOut>',self.leave,add='+')

    def enter(self,event=None):
        if self.imm is not None and self.context is None:
            old=self.imm.ImmAssociateContext(self.canvas.winfo_id(),None)
            if old:self.context=old

    def leave(self,event=None):
        if self.imm is not None and self.context is not None:
            self.imm.ImmAssociateContext(self.canvas.winfo_id(),self.context);self.context=None
