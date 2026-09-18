"""Origin 自动化公共模块：冷启动 + attach + 常用封装。

用法:
    from origin_helper import ensure_origin
    op = ensure_origin()          # 确保 OriginPro 在运行并已连接
"""
import os
import subprocess
import time

import originpro as op
import win32com.client


def _find_origin_exe():
    """定位 Origin64.exe：环境变量 > 注册表 App Paths > 常见安装目录。"""
    env = os.environ.get("ORIGIN_EXE")
    if env and os.path.exists(env):
        return env
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Origin64.exe") as k:
            v = winreg.QueryValueEx(k, "")[0]
            if v and os.path.exists(v):
                return v
    except OSError:
        pass
    import glob
    for pat in (r"C:\Program Files\OriginLab\Origin*\Origin64.exe",
                r"D:\Program Files\OriginLab\Origin*\Origin64.exe"):
        hits = sorted(glob.glob(pat), reverse=True)     # 新版本优先
        if hits:
            return hits[0]
    raise RuntimeError("Origin64.exe not found; set ORIGIN_EXE env var to its full path")


try:
    ORIGIN_EXE = _find_origin_exe()
except RuntimeError:
    ORIGIN_EXE = None   # 已开实例 attach 仍可用；冷启动时才报错


def _connect() -> bool:
    """Origin 不注册 ROT；进程在跑时 Dispatch 会附着到已有实例。"""
    try:
        win32com.client.Dispatch("Origin.ApplicationSI")
        return True
    except Exception:
        return False


def ensure_origin(timeout: float = 120.0) -> "op":
    """保证有可连接的 Origin 实例；没有就拉起，返回已 attach 的 originpro 模块。"""
    if not _connect():
        if not ORIGIN_EXE:
            raise RuntimeError("Origin is not running and Origin64.exe was not found; "
                               "start Origin manually or set ORIGIN_EXE env var")
        subprocess.Popen([ORIGIN_EXE])
        deadline = time.time() + timeout
        while time.time() < deadline:
            time.sleep(4)
            if _connect():
                break
        else:
            raise RuntimeError(f"Origin failed to become ready within {timeout:.0f}s")
    # COM 就绪 ≠ LabTalk/originpro 就绪：等窗口出现并额外缓冲
    for _ in range(15):
        p = subprocess.run(
            ['tasklist', '/FI', 'IMAGENAME eq Origin64.exe', '/FO', 'CSV'],
            capture_output=True, text=True)
        if "Origin64.exe" in p.stdout:
            break
        time.sleep(2)
    time.sleep(8)
    op.attach()
    return op


def new_sheet(wb_name: str, cols: dict[str, list], units: dict[str, str] | None = None):
    """新建 workbook 并写入多列数据。cols: {列名: [值,...]}"""
    o = ensure_origin()
    wb = o.new_book("w", wb_name)
    ws = wb[0]
    for i, (name, data) in enumerate(cols.items()):
        unit = (units or {}).get(name, "")
        if i < ws.cols:
            ws.from_list(i, data, lname=name, units=unit)
        else:
            ws.from_list(i, data, lname=name, units=unit)  # 该版本会自动扩列则最好，失败时下面兜底
    return wb, ws


__all__ = ["ensure_origin", "new_sheet", "op"]
