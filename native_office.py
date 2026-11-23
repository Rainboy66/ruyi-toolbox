"""Owned local PowerPoint/Word export; never attach to a user's open session."""
import ctypes
import json
import os
from pathlib import Path
import subprocess
import sys
import time


def find_powerpoint():
    return _find_application('PowerPoint.Application')


def find_word():
    return _find_application('Word.Application')


def find_excel():
    return _find_application('Excel.Application')


WPS_PROGIDS = {'word': 'KWPS.Application', 'sheet': 'KET.Application', 'slides': 'KWPP.Application'}


def _wps_registration(kind):
    """WPS frequently registers its local server only in the 32-bit view."""
    if os.name != 'nt':
        return None
    import winreg
    progid = WPS_PROGIDS[kind]
    for view, bits in ((winreg.KEY_WOW64_32KEY, 32), (winreg.KEY_WOW64_64KEY, 64)):
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + r'\CLSID', 0, winreg.KEY_READ | view) as key:
                clsid = winreg.QueryValueEx(key, '')[0]
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf'CLSID\{clsid}\LocalServer32', 0, winreg.KEY_READ | view) as key:
                command = winreg.QueryValueEx(key, '')[0]
            end = command.lower().find('.exe')
            path = Path(command[:end + 4].strip('" ')) if end >= 0 else None
            if path and path.is_file():
                return {'path': path, 'bits': bits, 'progid': progid}
        except OSError:
            continue
    return None


def find_wps(kind='word'):
    registration = _wps_registration(kind)
    return registration['path'] if registration else None


def _find_application(progid):
    if os.name != 'nt':
        return None
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + r'\CLSID') as key:
            clsid=winreg.QueryValueEx(key,'')[0]
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf'CLSID\{clsid}\LocalServer32') as key:
            command=winreg.QueryValueEx(key,'')[0]
        end=command.lower().find('.exe')
        path=Path(command[:end+4].strip('" ')) if end>=0 else None
        return path if path and path.is_file() else None
    except OSError:
        return None


def _owned_handle(owner_file, process_name='powerpnt.exe', *, owner_data=None):
    """Keep a process handle, checked against creation time, to avoid PID reuse."""
    if os.name != 'nt' or (owner_data is None and not owner_file.is_file()):
        return None
    from ctypes import wintypes as W
    try:
        data=owner_data if owner_data is not None else json.loads(owner_file.read_text(encoding='utf-8-sig'))
        kernel=ctypes.WinDLL('kernel32',use_last_error=True)
        kernel.OpenProcess.argtypes=[W.DWORD,W.BOOL,W.DWORD];kernel.OpenProcess.restype=W.HANDLE
        kernel.CloseHandle.argtypes=[W.HANDLE]
        kernel.GetProcessTimes.argtypes=[W.HANDLE]+[ctypes.POINTER(W.FILETIME)]*4
        kernel.QueryFullProcessImageNameW.argtypes=[W.HANDLE,W.DWORD,W.LPWSTR,ctypes.POINTER(W.DWORD)]
        kernel.TerminateProcess.argtypes=[W.HANDLE,W.UINT]
        kernel.WaitForSingleObject.argtypes=[W.HANDLE,W.DWORD]
        handle=kernel.OpenProcess(0x00100000|0x1000|0x0001,False,int(data['pid']))
        if not handle:return None
        created,exit_time,cpu,wall=(W.FILETIME() for _ in range(4))
        name=ctypes.create_unicode_buffer(32768);length=W.DWORD(len(name))
        valid=kernel.GetProcessTimes(handle,ctypes.byref(created),ctypes.byref(exit_time),ctypes.byref(cpu),ctypes.byref(wall))
        valid=valid and ((created.dwHighDateTime<<32)|created.dwLowDateTime)==int(data['created'])
        valid=valid and kernel.QueryFullProcessImageNameW(handle,0,name,ctypes.byref(length)) and Path(name.value).name.lower()==process_name.lower()
        if data.get('path'):
            valid=valid and os.path.normcase(os.path.abspath(name.value))==os.path.normcase(os.path.abspath(data['path']))
        if not valid:
            kernel.CloseHandle(handle);return None
        return kernel,handle
    except (OSError,ValueError,KeyError):
        return None


def export_powerpoint(source,output,workspace,check_cancel,timeout=45):
    return _export_native(source,output,workspace,check_cancel,timeout)


def export_word(source,output,workspace,check_cancel,timeout=45):
    return _export_native(source,output,workspace,check_cancel,timeout,word=True)


def export_excel(source,output,workspace,check_cancel,timeout=120):
    return _export_native(source,output,workspace,check_cancel,timeout,excel=True)


def export_wps(source,output,workspace,check_cancel,timeout=120,*,kind='word'):
    registration = _wps_registration(kind)
    if not registration:
        raise ValueError('未找到可用的 WPS 本机导出接口。')
    return _export_native(source,output,workspace,check_cancel,timeout,wps=registration,kind=kind)


def _export_native(source,output,workspace,check_cancel,timeout,*,word=False,excel=False,wps=None,kind=None):
    """The caller owns fallback and cancellation semantics. No shell string interpolation."""
    owner=workspace/'native-owner.json'
    result=workspace/'native-result.json'
    job=workspace/'native-job.json'
    script_name='native_wps_export.ps1' if wps else 'native_excel_export.ps1' if excel else 'native_word_export.ps1' if word else 'native_export.ps1'
    script=Path(getattr(sys,'_MEIPASS',Path(__file__).resolve().parent))/script_name
    process_name=wps['path'].name.lower() if wps else 'excel.exe' if excel else 'winword.exe' if word else 'powerpnt.exe'
    application='WPS' if wps else 'Excel' if excel else 'Word' if word else 'PowerPoint'
    payload={'source':str(source),'output':str(output),'owner':str(owner),'result':str(result)}
    if wps:payload.update(kind=kind,progid=wps['progid'],executable=str(wps['path']))
    job.write_text(json.dumps(payload),encoding='utf-8')
    shell_folder='SysWOW64' if wps and wps['bits']==32 and os.environ.get('PROCESSOR_ARCHITECTURE','').endswith('64') else 'System32'
    powershell=Path(os.environ.get('SystemRoot','C:/Windows'))/shell_folder/'WindowsPowerShell/v1.0/powershell.exe'
    started=time.monotonic()
    native_handles={}
    def collect_handles():
        if not owner.is_file():return
        try:
            data=json.loads(owner.read_text(encoding='utf-8-sig'))
            for entry in data.get('owners',[data]):
                identity=(int(entry['pid']),str(entry['created']))
                if identity not in native_handles:
                    handle=_owned_handle(owner,process_name,owner_data=entry)
                    if handle:native_handles[identity]=handle
        except (OSError,ValueError,KeyError):
            pass
    abnormal=True
    with (workspace/'native.log').open('w+b') as log:
        bundled=getattr(sys,'frozen',False)
        if bundled:ctypes.windll.kernel32.SetDllDirectoryW(None)
        try:
            process=subprocess.Popen([str(powershell),'-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',str(script),'-JobPath',str(job)],
                                     stdout=log,stderr=log,stdin=subprocess.DEVNULL,creationflags=subprocess.CREATE_NO_WINDOW)
        finally:
            if bundled:ctypes.windll.kernel32.SetDllDirectoryW(sys._MEIPASS)
        try:
            while process.poll() is None:
                collect_handles()
                check_cancel()
                if time.monotonic()-started>timeout:raise TimeoutError(application+' 本地导出超时')
                time.sleep(.1)
            check_cancel()
            if process.returncode!=0 or not output.is_file():
                log.seek(0)
                message=log.read().decode('utf-8',errors='replace').strip()[-500:]
                # The PowerShell workers stay ASCII-only (32-bit PowerShell reads
                # BOM-less scripts as ANSI), so user-facing text is produced here.
                friendly=message
                if 'SESSION_CONFLICT' in message:
                    friendly=('无法确认一个独立且空白的{}会话；为不影响你正在编辑的文件，本次未导出。'
                              '请关闭{}后重试，或在“排版基准”中选择另一个办公软件。').format(application,application)
                elif 'Could not confirm a dedicated empty Word session' in message:
                    friendly=('无法确认一个独立的空白 Word 会话；为不影响你正在编辑的文件，本次未导出。'
                              '请关闭 Word 后重试，或在“排版基准”中选择 WPS 文字。')
                raise ValueError(application+' 未完成导出：'+friendly)
            if not output.read_bytes().startswith(b'%PDF-'):
                raise ValueError(application+' 输出不是有效 PDF')
            abnormal=False
        finally:
            # Allow COM activation to publish ownership before terminating the worker.
            # Never kill an Office process identified merely by its name.
            if abnormal and not native_handles:
                cleanup_deadline=time.monotonic()+10
                while process.poll() is None and time.monotonic()<cleanup_deadline:
                    collect_handles()
                    if native_handles:break
                    time.sleep(.1)
            collect_handles()
            if abnormal:
                for kernel,handle in reversed(list(native_handles.values())):
                    kernel.TerminateProcess(handle,1)
                    kernel.WaitForSingleObject(handle,5000)
            if process.poll() is None:
                process.kill();process.wait(timeout=10)
            # Quit returns before Office's process has necessarily exited.
            # Allow the owned worker to finish closing before the next batch
            # item checks for an existing user session. Never kill on success.
            if not abnormal:
                for kernel,handle in native_handles.values():kernel.WaitForSingleObject(handle,10000)
            for kernel,handle in native_handles.values():kernel.CloseHandle(handle)
    return json.loads(result.read_text(encoding='utf-8-sig'))
