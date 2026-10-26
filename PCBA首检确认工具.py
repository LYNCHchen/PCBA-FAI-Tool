# -*- coding: utf-8 -*-
"""
PCBA 首检确认工具  (PCBA First Article Inspection Tool)
========================================================
Windows 桌面程序，Python 3 + Tkinter + Pillow，无需其他第三方依赖。

功能：
  1. 加载 1 张高清 PCB 图片（支持缩放 / 拖动 / 全景导航）
  2. 导入按顺序排列的元件坐标文件（制表符 / 逗号 / 空格等分隔，自动识别表头）
  3. 三点对位：在图片上点击 3 个元件的位置，自动计算整体镜像 / 旋转 / 偏移 / 比例
  4. 严格按照坐标文件顺序手动逐个确认（合格 / 不合格）
  5. ↑ / ↓ 方向键：上一步 / 下一步（跳转时当前未判定点位自动判合格 OK）
  6. 中心十字标、丝印位号、旋转指示、已检标记等均可独立显示 / 隐藏
  7. 结果一键导出 XLSX（NG 点位自动附局部截图）
  8. 项目导入 / 导出：图片 + 坐标 + 对位参数 + 检查进度打包为 .pba，可再次导入直接使用

快捷键：
  ↑ / ↓           上一步 / 下一步（当前看过的点位自动记为合格 OK）
  Enter            合格（OK）并自动下一步
  空格             不合格（NG）并自动下一步
  Esc              取消对位 / 取消选择
  Ctrl+O           打开图片
  Ctrl+I           导入坐标
  Ctrl+S           导出结果 XLSX

项目文件（.pba）：
  导出项目   把当前 图片 + 坐标 + 三点对位参数 + 已检查进度 打包为单个 .pba 文件
  导入项目   一键恢复图片、坐标、对位参数与检查进度，直接继续检查，无需重新对位
"""

import math
import os
import csv
import json
import io
import time
import zipfile
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import simpledialog
from PIL import Image, ImageTk, ImageOps, ImageEnhance

# ---------------- 高分屏适配（必须在创建窗口前调用） ----------------
try:
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        ctypes.windll.user32.SetProcessDPIAware()
except Exception:
    pass

# ---------------- 界面常量 ----------------
FONT      = ("Microsoft YaHei UI", 10)
FONT_S    = ("Microsoft YaHei UI", 9)
FONT_MONO = ("Consolas", 10)
FONT_BIG  = ("Microsoft YaHei UI", 22, "bold")

COL_CANVAS = "#20242a"
COL_CUR    = "#ff4d4f"   # 当前元件（红色）
COL_OK     = "#2ecc71"   # 合格（绿色）
COL_NG     = "#ff5757"   # 不合格（红色）

# ---------------- 操作说明文本 ----------------
HELP_TEXT = """==================== PCBA 首检确认工具 操作说明 ====================

【1. 图片加载】
· 点击「打开图片」或按 Ctrl+O，选择一张高清 PCB 图片。
· 支持格式：png / jpg / jpeg / bmp / tiff / webp（自动处理相机方向）。
· 鼠标滚轮缩放，左键拖动平移，整板与局部自由切换。

【2. 导入坐标】
· 点击「导入坐标」或按 Ctrl+I，选择元件坐标文件。
· 自动识别表头与分隔符（制表符 / 逗号 / 分号 / 空格均可）。
· 默认：第 1 行为表头，第 1 列序号、第 2 列位号、第 3 列 X、
  第 4 列 Y、第 5 列角度；支持 Q=角度、面别=镜像、点位/X转换mm 等常见别名。
· 表头无法识别时按固定列序（序号/位号/X/Y/角度）兜底，缺列不报错。

【3. 三点对位】（核心）
· 点「开始对位」进入对位模式，每个点按下面三步操作：
  1) 在 P1 行输入位号（键盘输入自动补全，或下拉选择），按【回车】确认；
  2) 移动大图到目标区域，双击图片上该元件位置，X/Y 自动录入 P1 行；
  3) 再移动大图到下一区域，P2 行输入位号+回车，双击对应点位；P3 同理。
· 三个点填齐后自动完成对位，程序计算 镜像/旋转/偏移/X-Y 比例。
· 若图片存在拉伸变形（拍照/扫描常见），自动按仿射变换精确拟合，全部点位对准。
· 对位完成后显示：镜像（是/否）、旋转角度、X 比例 × Y 比例、三点误差（px）。
· 「撤销上点」可回退最近一次双击；「清除对位」可重新开始。

【4. 逐个确认】
· 程序自动定位到第 1 个元件，选中点位自动移动到画面正中心。
· 合格：按 ↓ 方向键即可——看过的点位自动记为合格（OK）并跳到下一个；
  也可按 Enter 或点「合格 OK」按钮。
· 不合格：按空格键或点「不合格 NG」按钮，可在备注框写明原因。
· 已判定为 NG 的点位不会被方向键自动覆盖。
· ↑ 上一步 / ↓ 下一步（在图片区域、空白处均有效；焦点在备注框或表格内时方向键让位）。

【5. 显示开关】（功能设置区）
· 中心十字标 / 丝印位号 / 旋转指示 / 已检标记 / 显示全部点位：
  均可独立勾选显示或隐藏，避免遮挡元件目视检查。
· 右侧点位列表与大图区域之间的分隔条可用鼠标左右拖动，自由调整宽度。
· 功能区搜索框：输入位号可直接跳转到对应点位。

【6. 导出结果（XLSX）】
· 点「导出结果 XLSX」或按 Ctrl+S，生成 Excel 检验表。
· 列：序号 / 位号 / X(mm) / Y(mm) / 角度 / 面别 / 结果 / NG局部截图 / 备注 / 确认时间。
· 判定为 NG 的点位自动截取该点图片局部区域（约 ±450px 范围），
  并嵌入对应 NG 行，便于复核；OK 行无截图。

【7. 项目导入 / 导出（续检免对位）】
· 「导出项目」：把 图片 + 坐标 + 对位参数 + 检查进度 打包为单个 .pba 文件。
· 「导入项目」：一键恢复全部内容，画面回到导出时位置，无需重新对位即可续检。
· 适合：中途休息/换电脑继续、同板复检、把任务交给同事。

【快捷键一览】
  ↑ / ↓      上一步 / 下一步（未判定点自动记 OK）
  Enter      合格 OK 并自动下一步
  空格        不合格 NG 并自动下一步
  Esc        取消三点对位 / 取消当前选择
  Ctrl+O     打开图片
  Ctrl+I     导入坐标
  Ctrl+S     导出结果 XLSX

【联系方式】
软件开发：林奇    日期：2026-10-11
微信：chenlinwx

⚠️ 本工具仅供个人学习、工厂内部参考，禁止商用！
"""

COL_WAIT   = "#ffd60a"   # 对位提示（黄色）
COL_LABEL  = "#ffffff"   # 丝印文字
COL_PEND   = "#9fb0c3"   # 未检点位


# ============================================================
#  数学部分：2x2 SVD 与 Umeyama 相似变换（支持镜像）
# ============================================================

def svd2x2(a, b, c, d):
    """2x2 矩阵 [[a,b],[c,d]] 的 SVD：A = U·diag(s1,s2)·V^T，s1 >= s2 >= 0。
    返回 (U, s1, s2, V)，U、V 均为正交矩阵。"""
    E = (a + d) / 2.0
    F = (a - d) / 2.0
    G = (c + b) / 2.0
    H = (c - b) / 2.0
    Q = math.hypot(E, H)
    R = math.hypot(F, G)
    s1 = Q + R
    s2 = Q - R
    a1 = math.atan2(G, F)
    a2 = math.atan2(H, E)
    theta = (a1 + a2) / 2.0
    phi   = (a1 - a2) / 2.0
    ct, st = math.cos(theta), math.sin(theta)
    cp, sp = math.cos(phi),   math.sin(phi)
    U = [[ct, -st], [st, ct]]
    V = [[cp, -sp], [sp, cp]]
    if s2 < 0:                      # 奇异值为负时翻转 U 的第二列（保持重建，镜像判定随之成立）
        s2 = -s2
        U = [[U[0][0], -U[0][1]], [U[1][0], -U[1][1]]]
    return U, s1, s2, V


def umeyama(src, dst):
    """由 >=2 组对应点计算相似变换（含镜像）：
       dst = scale * R * src + t
    src / dst：[(x, y), ...] 列表（CAD 坐标 mm -> 图片像素坐标）。
    返回 dict(scale, r00,r01,r10,r11, tx,ty, mirrored, angle_deg, err_px)；
    数据退化时返回 None。"""
    n = len(src)
    msx = sum(p[0] for p in src) / n
    msy = sum(p[1] for p in src) / n
    mdx = sum(p[0] for p in dst) / n
    mdy = sum(p[1] for p in dst) / n
    a = b = c = d = 0.0
    var = 0.0
    for (sx, sy), (dx, dy) in zip(src, dst):
        ex, ey = sx - msx, sy - msy
        fx, fy = dx - mdx, dy - mdy
        a += ex * fx; b += ex * fy
        c += ey * fx; d += ey * fy
        var += ex * ex + ey * ey
    var /= n
    if var < 1e-12:
        return None
    a /= n; b /= n; c /= n; d /= n
    U, s1, s2, V = svd2x2(a, b, c, d)
    # Kabsch 解法：Sigma = U·diag(s1,s2)·V^T，最优旋转/镜像 R = V·U^T
    # （det(V·U^T) 自动取到最优的镜像符号）
    r00 = V[0][0] * U[0][0] + V[0][1] * U[0][1]
    r01 = V[0][0] * U[1][0] + V[0][1] * U[1][1]
    r10 = V[1][0] * U[0][0] + V[1][1] * U[0][1]
    r11 = V[1][0] * U[1][0] + V[1][1] * U[1][1]
    scale = (s1 + s2) / var
    tx = mdx - scale * (r00 * msx + r01 * msy)
    ty = mdy - scale * (r10 * msx + r11 * msy)
    err = 0.0
    for (sx, sy), (dx, dy) in zip(src, dst):
        px = scale * (r00 * sx + r01 * sy) + tx
        py = scale * (r10 * sx + r11 * sy) + ty
        err += (px - dx) ** 2 + (py - dy) ** 2
    err = math.sqrt(err / n)
    mirrored = (r00 * r11 - r01 * r10) < 0
    angle_deg = math.degrees(math.atan2(r10, r00))
    return dict(scale=scale, r00=r00, r01=r01, r10=r10, r11=r11,
                tx=tx, ty=ty, mirrored=mirrored,
                angle_deg=angle_deg, err_px=err)


def affine_from_3pts(src, dst):
    """由 3 组对应点精确求解仿射变换（6 参数，克莱默法则）：
       u = a*x + b*y + c ;  v = d*x + e*y + f
    可纠正图片的拉伸变形（x/y 方向不同比例、剪切、旋转、镜像）。
    src / dst：各 3 个 (x, y)；三点共线时返回 None。"""
    (x1, y1), (x2, y2), (x3, y3) = src
    (u1, v1), (u2, v2), (u3, v3) = dst
    D = x1 * (y2 - y3) + x2 * (y3 - y1) + x3 * (y1 - y2)
    if abs(D) < 1e-9:
        return None
    a = (u1 * (y2 - y3) + u2 * (y3 - y1) + u3 * (y1 - y2)) / D
    b = (u1 * (x3 - x2) + u2 * (x1 - x3) + u3 * (x2 - x1)) / D
    c = (u1 * (x2 * y3 - x3 * y2) + u2 * (x3 * y1 - x1 * y3) + u3 * (x1 * y2 - x2 * y1)) / D
    d = (v1 * (y2 - y3) + v2 * (y3 - y1) + v3 * (y1 - y2)) / D
    e = (v1 * (x3 - x2) + v2 * (x1 - x3) + v3 * (x2 - x1)) / D
    f = (v1 * (x2 * y3 - x3 * y2) + v2 * (x3 * y1 - x1 * y3) + v3 * (x1 * y2 - x2 * y1)) / D
    sx = math.hypot(a, d)              # X 方向缩放（图片 x 轴单位长度映射）
    sy = math.hypot(b, e)              # Y 方向缩放
    det = a * e - b * d
    return dict(a=a, b=b, c=c, d=d, e=e, f=f,
                scale=math.sqrt(abs(det)),        # 几何平均比例（兼容字段）
                angle_deg=math.degrees(math.atan2(d, a)),  # 旋转角（兼容字段）
                mirrored=det < 0,
                err_px=0.0, sx=sx, sy=sy)


def transform_point(t, x, y):
    """把 CAD 坐标 (x,y) 按变换 t 映射到图片像素坐标（仿射/相似通用）。"""
    if "a" in t:                       # 仿射变换
        return (t["a"] * x + t["b"] * y + t["c"],
                t["d"] * x + t["e"] * y + t["f"])
    return (t["scale"] * (t["r00"] * x + t["r01"] * y) + t["tx"],
            t["scale"] * (t["r10"] * x + t["r11"] * y) + t["ty"])


# ============================================================
#  坐标文件解析
# ============================================================

_HEADER_ALIAS = {
    "no":     {"no", "序号", "seq", "num", "number", "id", "编号", "item", "序列号", "项号", "元件编号"},
    "refdes": {"refdes", "ref", "位号", "位置名称", "designator", "name", "元件位号",
               "参考位号", "refdes1", "元件号", "元件名称", "位号名称", "ref1", "点位",
               "元件位号名称", "refdes名称"},
    "x":      {"x", "symx", "xmm", "xcad", "posx", "centerx", "x坐标", "坐标x",
               "x轴坐标", "横坐标", "xposition", "xc", "x转换mm", "x值", "坐标xmm"},
    "y":      {"y", "symy", "ymm", "ycad", "posy", "centery", "y坐标", "坐标y",
               "y轴坐标", "纵坐标", "yposition", "yc", "y转换mm", "y值", "坐标ymm"},
    "rot":    {"rot", "rotate", "rotation", "symrotate", "angle", "角度", "角度度",
               "度", "q", "theta", "deg", "旋转角度", "角度值", "方向", "orientation",
               "orient", "q坐标", "旋转", "角度坐标"},
    "mir":    {"mirror", "mir", "symmirror", "镜像", "side", "layer", "层", "面",
               "面别", "板面", "表面", "face", "surface", "boardface", "正反面",
               "topbottom", "bottomside", "topside"},
}


def _norm_header(s):
    s = s.strip().lower().replace(" ", "").replace("_", "")
    s = s.replace("(", "").replace(")", "").replace(".", "")
    return s


def _xlsx_to_rows(path):
    """读取 .xlsx → 返回 [[单元格文本/数字, ...], ...]（含表头），用于复用文本解析逻辑。"""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    out = []
    for row in ws.iter_rows(values_only=True):
        cells = []
        for v in row:
            if v is None:
                cells.append("")
            elif isinstance(v, float):
                # 整数不带 .0
                cells.append(str(int(v)) if v == int(v) else ("%.4f" % v).rstrip("0").rstrip("."))
            else:
                cells.append(str(v).strip())
        out.append(cells)
    wb.close()
    return out


def _parse_rows(rows):
    """把二维 rows（含表头）解析成坐标行；与 parse_coord_file 共用逻辑。"""
    rows = [r for r in rows if any(str(c).strip() for c in r)]
    if not rows:
        raise ValueError("坐标文件为空")
    head = rows[0]
    hdr = [str(c).strip() for c in head if str(c).strip()]

    idx = {}
    header_ok = False
    for i, c in enumerate(head):
        n = _norm_header(str(c))
        for kind, aliases in _HEADER_ALIAS.items():
            if n in aliases:
                idx[kind] = i
                header_ok = True
                break
    positional = ("refdes" not in idx) or ("x" not in idx) or ("y" not in idx)
    start = 1
    if positional:
        idx = {"no": 0, "refdes": 1, "x": 2, "y": 3, "rot": 4, "mir": 5}
        p0 = [str(c).strip() for c in rows[0]]

        def _isnum(s):
            try:
                float(str(s).replace("mm", "").strip()); return True
            except Exception:
                return False
        if len(p0) >= 4 and _isnum(p0[2]) and _isnum(p0[3]):
            start = 0

    out = []
    for parts in rows[start:]:
        parts = [str(p).strip() for p in parts]
        if not parts or all(ch in "-=~ " for ch in parts[0]):
            continue

        def num(s, _i):
            try:
                return float(str(s).replace("mm", "").strip())
            except Exception:
                return None

        def val(i):
            return parts[i] if i < len(parts) else ""

        refdes = val(idx["refdes"])
        x = num(val(idx["x"]), 0)
        y = num(val(idx["y"]), 0)
        if x is None or y is None:
            continue
        no = num(val(idx["no"]), 0)
        ri = idx.get("rot")
        rot = num(val(ri), 0) if ri is not None else 0.0
        if rot is None:
            rot = 0.0
        mi = idx.get("mir")
        mir = val(mi) if mi is not None else ""
        out.append(dict(
            no=int(no) if no is not None and no == int(no) else (len(out) + 1),
            refdes=refdes or "#%d" % (len(out) + 1),
            x=x, y=y, rot=rot, mir=mir,
            result=None, remark="", time=""))
    if not out:
        raise ValueError("未解析到任何元件数据")
    return out


def parse_coord_file(path):
    """解析坐标文件（txt/csv/dat/pos 或 xlsx），返回 [ {no,refdes,x,y,rot,mir}, ... ]（保持文件顺序）。"""
    # Excel 走 openpyxl
    if str(path).lower().endswith((".xlsx", ".xlsm", ".xls")):
        rows = _xlsx_to_rows(path)
        return _parse_rows(rows)

    text = None
    for enc in ("utf-8-sig", "gbk", "utf-8"):
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                text = f.read()
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            text = f.read()
    lines = [ln.rstrip("\r\n") for ln in text.splitlines()]
    lines = [ln for ln in lines if ln.strip()]
    if not lines:
        raise ValueError("坐标文件为空")

    head = lines[0]
    cand = {"\t": head.count("\t"), ",": head.count(","),
            ";": head.count(";"), "|": head.count("|")}
    sep = max(cand, key=cand.get)
    if cand[sep] == 0:
        sep = None
    hdr = head.split(sep) if sep else head.split()
    hdr = [c.strip() for c in hdr if c.strip()]

    # 把文本行转成二维 rows，复用 _parse_rows
    rows = []
    for ln in lines:
        parts = ln.split(sep) if sep else ln.split()
        rows.append([p.strip() for p in parts])
    return _parse_rows(rows)


# ============================================================
#  主程序
# ============================================================

class App:
    current_user = "-"
    current_role = "user"

    def __init__(self, root):
        self.root = root
        root.title("PCBA 首检确认工具")
        root.geometry("1560x920")
        root.minsize(1180, 720)

        # 状态
        self.img = None            # PIL 原始图
        self.img_path = None
        self.coords = []           # 坐标行
        self.coord_path = None
        self.transform = None      # umeyama 结果
        self.zoom = None           # 像素/图片像素
        self.ox = 0.0
        self.oy = 0.0
        self._photo = None         # 画布 PhotoImage 引用
        self.cur = -1              # 当前元件下标
        self.align_mode = False
        self.align_refs = []
        self.align_pts = []        # 图片像素坐标
        self.align_confirmed = [False, False, False]   # P1/P2/P3 位号是否已回车确认
        self._pan = None
        self._loading_tree = False
        # 全景图
        self.mm_img = None
        self.mm_photo = None
        self.mm_scale = 1.0
        self.mm_ox = 0.0
        self.mm_oy = 0.0
        # 亮度 / 对比度（-100 ~ 100，0 = 原图）；须在 _build_ui 之前定义
        self.v_bright   = tk.IntVar(value=0)
        self.v_contrast = tk.IntVar(value=0)

        self._build_menu()
        self._build_ui()
        self._bind()
        self._refresh_status()
        self._set_window_icon()

    # ---------------- 顶部菜单栏 ----------------
    def _build_menu(self):
        root = self.root
        menubar = tk.Menu(root)

        m_sys = tk.Menu(menubar, tearoff=0)
        m_sys.add_command(label="打开图片…", accelerator="Ctrl+O", command=lambda: self.load_image())
        m_sys.add_command(label="导入坐标…", accelerator="Ctrl+I", command=lambda: self.load_coords())
        m_sys.add_separator()
        m_sys.add_command(label="导出结果 XLSX…", accelerator="Ctrl+S", command=lambda: self.export_xlsx())
        m_sys.add_command(label="导出项目…", command=lambda: self.export_project())
        m_sys.add_command(label="导入项目…", command=lambda: self.import_project())
        m_sys.add_separator()
        m_sys.add_command(label="退出", command=root.quit)
        menubar.add_cascade(label="系统(S)", menu=m_sys)

        m_edit = tk.Menu(menubar, tearoff=0)
        m_edit.add_command(label="上一步", accelerator="↑", command=lambda: self.prev())
        m_edit.add_command(label="下一步", accelerator="↓", command=lambda: self.next_())
        m_edit.add_command(label="合格 OK", accelerator="Enter", command=lambda: self.mark(True))
        m_edit.add_command(label="不合格 NG", accelerator="空格", command=lambda: self.mark(False))
        m_edit.add_command(label="搜索位号…", accelerator="Ctrl+F", command=lambda: self._search_dialog())
        m_edit.add_separator()
        m_edit.add_command(label="开始对位", command=lambda: self.begin_align())
        m_edit.add_command(label="撤销上点", command=lambda: self.undo_align_pt())
        m_edit.add_command(label="清除对位", command=lambda: self.clear_align())
        menubar.add_cascade(label="编辑(U)", menu=m_edit)

        m_win = tk.Menu(menubar, tearoff=0)
        m_win.add_command(label="复位整板视图", command=lambda: self._fit_view())
        m_win.add_command(label="放大", command=lambda: self._zoom_center(1.25))
        m_win.add_command(label="缩小", command=lambda: self._zoom_center(0.8))
        menubar.add_cascade(label="窗口(W)", menu=m_win)

        m_user = tk.Menu(menubar, tearoff=0)
        m_user.add_command(label="当前用户：%s（%s）" %
                            (getattr(App, "current_user", "-"), getattr(App, "current_role", "-")),
                           state="disabled")
        m_user.add_separator()
        m_user.add_command(label="修改密码…", command=lambda: self._change_pwd_dialog())
        if getattr(App, "current_role", "user") == "admin":
            m_user.add_command(label="用户管理…", command=lambda: self._user_mgmt_dialog())
        m_user.add_separator()
        m_user.add_command(label="注销 / 切换用户", command=lambda: self._relogin())
        menubar.add_cascade(label="我的账户", menu=m_user)

        m_help = tk.Menu(menubar, tearoff=0)
        m_help.add_command(label="操作说明", command=lambda: self.show_help())
        m_help.add_command(label="关于", command=lambda: self._about_dialog())
        menubar.add_cascade(label="帮助", menu=m_help)

        root.config(menu=menubar)

    def _about_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("About PCBA 首检确认工具 (FAI)")
        win.geometry("460x360")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        # 居中
        win.update_idletasks()
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        win.geometry("+%d+%d" % ((sw - 460) // 2, (sh - 360) // 2))

        tk.Label(win, text="PCBA 首检确认工具",
                 font=("Microsoft YaHei UI", 18, "bold"), fg="#e08a00").pack(pady=(28, 2))
        tk.Label(win, text="First Article Inspection (FAI)",
                 font=("Microsoft YaHei UI", 12), fg="#1a6fb5").pack(pady=(0, 22))

        info = [
            ("APP Version:", "  1.5.0 (2026.10.14)"),
            ("DB Version:", "  Local Users JSON"),
            ("Copyright:",     " @ 2026 林奇 / 技研新阳有限公司"),
        ]
        for k, v in info:
            row = tk.Frame(win)
            row.pack(pady=3)
            tk.Label(row, text=k, font=("Consolas", 11, "bold"), width=12, anchor="e").pack(side="left")
            tk.Label(row, text=v, font=("Consolas", 11)).pack(side="left")

        ttk.Button(win, text="OK", command=win.destroy).pack(pady=18)

        tk.Label(win, text="技研新阳有限公司",
                 font=("Microsoft YaHei UI", 12, "bold"), fg="#1a6fb5").pack()
        tk.Label(win, text="SHIN TECH Co., Ltd.",
                 font=("Microsoft YaHei UI", 9), fg="#666").pack(pady=(2, 0))
        tk.Label(win, text="微信：chenlinwx  |  仅供学习与工厂内部参考，禁止商用",
                 font=("Microsoft YaHei UI", 8), fg="#999").pack(side="bottom", pady=6)

    def _relogin(self):
        root = self.root
        root.withdraw()
        if login_dialog(root):
            root.deiconify()
            messagebox.showinfo("提示", "已切换到用户：%s" % getattr(App, "current_user", "-"),
                                parent=root)
            self._build_menu()
        else:
            root.deiconify()

    def _change_pwd_dialog(self):
        users = load_users()
        u = getattr(App, "current_user", None)
        if not u or u not in users:
            return
        win = tk.Toplevel(self.root)
        win.title("修改密码")
        win.geometry("380x260")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        win.update_idletasks()
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        win.geometry("+%d+%d" % ((sw - 380) // 2, (sh - 260) // 2))
        ttk.Label(win, text="用户：" + u,
                  font=("Microsoft YaHei UI", 11)).pack(pady=(18, 10))
        ttk.Label(win, text="新密码：").pack()
        e1 = ttk.Entry(win, show="*", width=26, font=("Microsoft YaHei UI", 10))
        e1.pack(pady=4)
        ttk.Label(win, text="确认新密码：").pack(pady=(8, 0))
        e2 = ttk.Entry(win, show="*", width=26, font=("Microsoft YaHei UI", 10))
        e2.pack(pady=4)
        def ok():
            p1, p2 = e1.get(), e2.get()
            if not p1:
                messagebox.showinfo("提示", "密码不能为空", parent=win); return
            if p1 != p2:
                messagebox.showinfo("提示", "两次输入不一致", parent=win); return
            users[u]["pwd"] = p1
            save_users(users)
            messagebox.showinfo("成功", "密码已修改", parent=win)
            win.destroy()
        ttk.Button(win, text="确定", command=ok).pack(pady=16)
        e1.focus_set()

    def _user_mgmt_dialog(self):
        if getattr(App, "current_role", "user") != "admin":
            messagebox.showinfo("提示", "仅管理员可进入用户管理")
            return
        users = load_users()
        win = tk.Toplevel(self.root)
        win.title("用户管理（管理员）")
        win.geometry("420x320")
        win.transient(self.root)
        win.grab_set()
        lst = tk.Listbox(win, height=10, font=FONT_MONO)
        lst.pack(fill="both", expand=True, padx=10, pady=10)
        def refresh():
            lst.delete(0, "end")
            for name, info in users.items():
                lst.insert("end", "  %-16s  角色=%s" % (name, info.get("role", "user")))
        refresh()
        btns = tk.Frame(win); btns.pack(pady=6)
        def add():
            u = tk.simpledialog.askstring("新增用户", "用户名：", parent=win)
            if not u: return
            u = u.strip()
            if u in users:
                messagebox.showinfo("提示", "用户已存在", parent=win); return
            p = tk.simpledialog.askstring("新增用户", "初始密码：", parent=win, show="*")
            if not p: return
            users[u] = {"pwd": p, "role": "user"}
            save_users(users); refresh()
        def delete():
            sel = lst.curselection()
            if not sel: return
            name = list(users.keys())[sel[0]]
            if name == "admin":
                messagebox.showinfo("提示", "admin 不能删除", parent=win); return
            if name == getattr(App, "current_user", None):
                messagebox.showinfo("提示", "不能删除当前登录用户", parent=win); return
            if not messagebox.askyesno("确认", "删除用户 %s？" % name, parent=win): return
            del users[name]; save_users(users); refresh()
        ttk.Button(btns, text="新增", command=add).pack(side="left", padx=6)
        ttk.Button(btns, text="删除", command=delete).pack(side="left", padx=6)
        ttk.Button(btns, text="关闭", command=win.destroy).pack(side="left", padx=6)

    def _set_window_icon(self):
        """设置窗口标题栏左上角图标（开发态读脚本目录，打包态读临时目录）。"""
        import sys
        try:
            base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
            png = os.path.join(base, "app_icon.png")
            ico = os.path.join(base, "app_icon.ico")
            if os.path.exists(png):
                self._icon_img = tk.PhotoImage(file=png)
                self.root.iconphoto(True, self._icon_img)
            elif os.path.exists(ico):
                self.root.iconbitmap(default=ico)
        except Exception:
            pass

    # ---------------- 界面构建 ----------------
    def _build_ui(self):
        root = self.root
        root.grid_rowconfigure(0, weight=1)
        root.grid_rowconfigure(1, weight=0)
        root.grid_columnconfigure(0, weight=1)
        root.grid_columnconfigure(1, weight=0)

        # 主区：左 = 图片画布（大面积），右 = 点位坐标列表 + 全景图
        # 中间分隔条可用鼠标左右拖动，自由调整两侧宽度
        self.main = ttk.Panedwindow(root, orient="horizontal")
        self.main.grid(row=0, column=0, sticky="nsew")

        # 左：图片画布
        left = tk.Frame(self.main, bg=COL_CANVAS)
        left.grid_rowconfigure(0, weight=1)
        left.grid_columnconfigure(0, weight=1)
        self.canvas = tk.Canvas(left, bg=COL_CANVAS, highlightthickness=0,
                                cursor="crosshair")
        self.canvas.grid(row=0, column=0, sticky="nsew")

        # 右：上方点位坐标列表，下方全景图
        right = tk.Frame(self.main, width=560)
        right.grid_propagate(False)
        right.grid_rowconfigure(0, weight=1)
        right.grid_columnconfigure(0, weight=1)

        self.main.add(left, weight=3)
        self.main.add(right, weight=0)

        # 右-上：点位坐标列表
        ft = ttk.LabelFrame(right, text=" 点位坐标列表 ", padding=4)
        ft.grid(row=0, column=0, sticky="nsew", padx=8, pady=(8, 4))
        ft.grid_rowconfigure(0, weight=1)
        ft.grid_columnconfigure(0, weight=1)
        cols = ("no", "refdes", "x", "y", "rot", "mir", "result", "remark")
        heads = ("序号", "位号", "X(mm)", "Y(mm)", "角度", "镜像", "结果", "备注")
        widths = (65, 90, 95, 95, 60, 65, 80, 260)
        self.tree = ttk.Treeview(ft, columns=cols, show="headings", height=12)
        for c, h, w in zip(cols, heads, widths):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor="center" if c != "remark" else "w")
        vs = ttk.Scrollbar(ft, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(ft, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        self.tree.tag_configure("OK", foreground="#1a7f37")
        self.tree.tag_configure("NG", foreground="#d1242f")

        # 右-下：全景图
        f5 = ttk.LabelFrame(right, text=" 全景（点击可跳转） ", padding=6)
        f5.grid(row=1, column=0, sticky="sew", padx=8, pady=(4, 8))
        self.mm_canvas = tk.Canvas(f5, width=200, height=150, bg="#16191d",
                                   highlightthickness=0)
        self.mm_canvas.pack(fill="both", expand=True)
        self.mm_canvas.bind("<Button-1>", self._mm_click)

        # 底部：功能设置区域（横向）
        bottom = ttk.Frame(root)
        bottom.grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 8))
        bottom.grid_columnconfigure(0, weight=5)
        bottom.grid_columnconfigure(1, weight=6)
        bottom.grid_columnconfigure(2, weight=5)
        bottom.grid_columnconfigure(3, weight=4)

        # 1) 文件
        f1 = ttk.LabelFrame(bottom, text=" 文件 ", padding=8)
        f1.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        f1.grid_columnconfigure(0, weight=1)
        bt = ttk.Frame(f1)
        bt.grid(row=0, column=0, sticky="ew")
        ttk.Button(bt, text="打开图片", command=self.load_image).pack(side="left", expand=True, fill="x", padx=(0, 3))
        ttk.Button(bt, text="导入坐标", command=self.load_coords).pack(side="left", expand=True, fill="x", padx=(0, 3))
        ttk.Button(bt, text="导入项目", command=self.import_project).pack(side="left", expand=True, fill="x", padx=(0, 3))
        ttk.Button(bt, text="导出项目", command=self.export_project).pack(side="left", expand=True, fill="x")
        self.lbl_img = ttk.Label(f1, text="图片：未加载", font=FONT_S)
        self.lbl_img.grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.lbl_crd = ttk.Label(f1, text="坐标：未导入", font=FONT_S)
        self.lbl_crd.grid(row=2, column=0, sticky="w")
        sr = ttk.Frame(f1)
        sr.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        ttk.Label(sr, text="搜索位号", font=FONT_S).pack(side="left")
        self.e_search = ttk.Entry(sr, font=FONT_MONO)
        self.e_search.pack(side="left", fill="x", expand=True, padx=4)
        self.e_search.bind("<KeyRelease>", lambda e: self._search_autocomplete())
        self.e_search.bind("<Return>", lambda e: (self.search_jump(), "break")[1])
        ttk.Button(sr, text="定位", width=5, command=self.search_jump).pack(side="left")

        # 亮度 / 对比度调整（作用于整体图片显示，方便目视确认）
        adj = ttk.Frame(f1)
        adj.grid(row=4, column=0, sticky="ew", pady=(6, 0))
        ttk.Label(adj, text="亮度", font=FONT_S).pack(side="left")
        ttk.Scale(adj, from_=-100, to=100, variable=self.v_bright,
                  command=lambda e: self._render()).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Label(adj, text="对比度", font=FONT_S).pack(side="left")
        ttk.Scale(adj, from_=-100, to=100, variable=self.v_contrast,
                  command=lambda e: self._render()).pack(side="left", fill="x", expand=True, padx=4)

        # 2) 三点对位
        f2 = ttk.LabelFrame(bottom, text=" 三点对位（镜像 / 旋转 / 偏移 自动计算） ", padding=8)
        f2.grid(row=0, column=1, sticky="nsew", padx=4)
        f2.grid_columnconfigure(1, weight=1)
        self.cb1 = ttk.Combobox(f2, font=FONT_MONO, width=9)
        self.cb2 = ttk.Combobox(f2, font=FONT_MONO, width=9)
        self.cb3 = ttk.Combobox(f2, font=FONT_MONO, width=9)
        self.e_x1 = ttk.Entry(f2, width=9, font=FONT_MONO)
        self.e_y1 = ttk.Entry(f2, width=9, font=FONT_MONO)
        self.e_x2 = ttk.Entry(f2, width=9, font=FONT_MONO)
        self.e_y2 = ttk.Entry(f2, width=9, font=FONT_MONO)
        self.e_x3 = ttk.Entry(f2, width=9, font=FONT_MONO)
        self.e_y3 = ttk.Entry(f2, width=9, font=FONT_MONO)
        ttk.Label(f2, text="", width=3).grid(row=0, column=0)
        ttk.Label(f2, text="位号", font=FONT_S).grid(row=0, column=1, sticky="w")
        ttk.Label(f2, text="X(px)", font=FONT_S).grid(row=0, column=2, sticky="w", padx=(0, 2))
        ttk.Label(f2, text="Y(px)", font=FONT_S).grid(row=0, column=3, sticky="w")
        for i, (lbl, cb, ex, ey) in enumerate(
                zip(("P1", "P2", "P3"),
                    (self.cb1, self.cb2, self.cb3),
                    (self.e_x1, self.e_x2, self.e_x3),
                    (self.e_y1, self.e_y2, self.e_y3)), start=1):
            ttk.Label(f2, text=lbl, width=3).grid(row=i, column=0, sticky="w")
            cb.grid(row=i, column=1, sticky="ew", padx=(0, 4), pady=1)
            cb.bind("<KeyRelease>", lambda e, c=cb, k=i - 1: self._cb_autocomplete(c, k))
            cb.bind("<Return>", lambda e, k=i - 1: (self._confirm_align_ref(k), "break")[1])
            cb.bind("<<ComboboxSelected>>", lambda e, k=i - 1: self._confirm_align_ref(k))
            ex.grid(row=i, column=2, padx=(0, 2), pady=1)
            ey.grid(row=i, column=3, pady=1)
        bt2 = ttk.Frame(f2)
        bt2.grid(row=4, column=0, columnspan=4, sticky="ew", pady=(4, 0))
        ttk.Button(bt2, text="开始对位", command=self.begin_align).pack(side="left", expand=True, fill="x", padx=(0, 3))
        ttk.Button(bt2, text="撤销上点", command=self.undo_align_pt).pack(side="left", expand=True, fill="x", padx=(0, 3))
        ttk.Button(bt2, text="清除对位", command=self.clear_align).pack(side="left", expand=True, fill="x")
        self.lbl_align = ttk.Label(f2, text="① 位号可直接键盘输入（如 C210，自动补全）→ ② 点“开始对位” → ③ 点击图片自动填 X/Y，或直接手输像素坐标",
                                   font=FONT_S, foreground="#555555", wraplength=300, justify="left")
        self.lbl_align.grid(row=5, column=0, columnspan=4, sticky="w", pady=(4, 0))
        self.lbl_param = ttk.Label(f2, text="", font=FONT_S, foreground="#1a6fb5", wraplength=300, justify="left")
        self.lbl_param.grid(row=6, column=0, columnspan=4, sticky="w")

        # 3) 检查确认
        f3 = ttk.LabelFrame(bottom, text=" 检查确认（严格按坐标顺序） ", padding=8)
        f3.grid(row=0, column=2, sticky="nsew", padx=4)
        f3.grid_columnconfigure(0, weight=1)
        self.lbl_cur = ttk.Label(f3, text="--", font=FONT_BIG, foreground=COL_CUR)
        self.lbl_cur.grid(row=0, column=0, sticky="w")
        self.lbl_info = ttk.Label(f3, text="尚未导入坐标", font=FONT_S, foreground="#444444")
        self.lbl_info.grid(row=1, column=0, sticky="w", pady=(0, 4))
        bt3 = ttk.Frame(f3)
        bt3.grid(row=2, column=0, sticky="ew")
        self.btn_ok = ttk.Button(bt3, text="合格  OK  (Enter)", command=lambda: self.mark(True))
        self.btn_ok.pack(side="left", expand=True, fill="x", padx=(0, 3))
        self.btn_ng = ttk.Button(bt3, text="不合格 NG  (空格)", command=lambda: self.mark(False))
        self.btn_ng.pack(side="left", expand=True, fill="x")
        rk = ttk.Frame(f3)
        rk.grid(row=3, column=0, sticky="ew", pady=(4, 0))
        ttk.Label(rk, text="备注").pack(side="left")
        self.entry_remark = ttk.Entry(rk)
        self.entry_remark.pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(rk, text="保存", width=5, command=self.save_remark).pack(side="left")
        self.prog = ttk.Progressbar(f3, maximum=100)
        self.prog.grid(row=4, column=0, sticky="ew", pady=(6, 2))
        self.lbl_prog = ttk.Label(f3, text="", font=FONT_S, foreground="#333333")
        self.lbl_prog.grid(row=5, column=0, sticky="w")

        # 4) 显示选项 + 导出
        f4 = ttk.LabelFrame(bottom, text=" 显示选项 + 导出 ", padding=8)
        f4.grid(row=0, column=3, sticky="nsew", padx=(4, 0))
        self.v_cross = tk.BooleanVar(value=True)
        self.v_box   = tk.BooleanVar(value=False)
        self.v_label = tk.BooleanVar(value=True)
        self.v_rot   = tk.BooleanVar(value=False)
        self.v_done  = tk.BooleanVar(value=True)
        self.v_all   = tk.BooleanVar(value=False)
        cb4 = ttk.Frame(f4)
        cb4.grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(cb4, text="中心十字", variable=self.v_cross,
                        command=self._render).grid(row=0, column=0, sticky="w", padx=(0, 10))
        ttk.Checkbutton(cb4, text="方框", variable=self.v_box,
                        command=self._render).grid(row=0, column=1, sticky="w", padx=(0, 10))
        ttk.Checkbutton(cb4, text="丝印位号", variable=self.v_label,
                        command=self._render).grid(row=0, column=2, sticky="w", padx=(0, 10))
        ttk.Checkbutton(cb4, text="旋转指示", variable=self.v_rot,
                        command=self._render).grid(row=0, column=3, sticky="w")
        ttk.Checkbutton(cb4, text="已检标记(OK绿/NG红)", variable=self.v_done,
                        command=self._render).grid(row=1, column=0, columnspan=2, sticky="w", padx=(0, 10))
        ttk.Checkbutton(cb4, text="显示全部点位", variable=self.v_all,
                        command=self._render).grid(row=1, column=2, sticky="w")
        f6 = ttk.Frame(f4)
        f6.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        f6.grid_columnconfigure(0, weight=1)
        ttk.Button(f6, text="导出结果 XLSX（NG 附截图）", command=self.export_xlsx).grid(row=0, column=0, sticky="ew")
        ttk.Button(f6, text="操作说明", command=self.show_help).grid(row=0, column=1, sticky="ew", padx=(6, 0))
        ttk.Label(f6, text="↑/↓ 上一步/下一步（看过的点自动判 OK）   空格=NG   Enter=OK\n滚轮缩放   左键拖动平移   右键撤销对位点",
                  font=FONT_S, foreground="#777777", justify="left").grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 0))

    def _bind(self):
        root = self.root
        root.bind("<Up>", lambda e: self._key_arrow(-1))
        root.bind("<Down>", lambda e: self._key_arrow(1))
        root.bind("<Return>", self._key_return)
        root.bind("<space>", self._key_space)
        root.bind("<Escape>", self._key_escape)
        root.bind("<Control-o>", lambda e: self.load_image())
        root.bind("<Control-i>", lambda e: self.load_coords())
        root.bind("<Control-s>", lambda e: self.export_xlsx())
        self.entry_remark.bind("<Return>", lambda e: (self.save_remark(), "break")[1])

        c = self.canvas
        c.bind("<ButtonPress-1>", self._on_press)
        c.bind("<B1-Motion>", self._on_drag)
        c.bind("<ButtonRelease-1>", self._on_release)
        c.bind("<MouseWheel>", self._on_wheel)
        c.bind("<Button-3>", self._on_right)
        c.bind("<Motion>", self._on_motion)
        c.bind("<Configure>", lambda e: self._on_configure())
        self.tree.bind("<<TreeviewSelect>>", self._tree_select)

    def _key_arrow(self, delta):
        if not self._key_guard(("entry", "tree")):
            return
        if delta < 0:
            self.prev()
        else:
            self.next_()

    # ---------------- 文件加载 ----------------
    def load_image(self, path=None):
        if path is None:
            path = filedialog.askopenfilename(
                title="选择 PCB 高清图片",
                filetypes=[("图片文件", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp"), ("所有文件", "*.*")])
            if not path:
                return
        try:
            img = Image.open(path)
            img = ImageOps.exif_transpose(img)      # 处理相机方向
            img = img.convert("RGB")
        except Exception as ex:
            messagebox.showerror("错误", f"图片打开失败：\n{ex}")
            return
        self.img = img
        self.img_path = path
        self._reset_view()
        self._reset_align()
        self.lbl_img.config(text="图片：" + os.path.basename(path) +
                                 f"   ({img.width}×{img.height})")
        self._make_minimap()
        self.root.update_idletasks()
        self._fit_view()
        self.canvas.focus_set()

    def load_coords(self, path=None):
        if path is None:
            path = filedialog.askopenfilename(
                title="导入元件坐标文件",
                filetypes=[("坐标文件", "*.txt *.csv *.dat *.pos *.xlsx *.xlsm"),
                           ("Excel 文件", "*.xlsx *.xlsm"),
                           ("所有文件", "*.*")])
            if not path:
                return
        try:
            rows = parse_coord_file(path)
        except Exception as ex:
            messagebox.showerror("错误", f"坐标文件解析失败：\n{ex}")
            return
        self.coords = rows
        self.coord_path = path
        self.cur = -1
        self._reset_align()
        self._fill_combos()
        self._fill_table()
        self.lbl_crd.config(text="坐标：" + os.path.basename(path) + f"   （{len(rows)} 个元件）")
        self._refresh_status()
        self.canvas.focus_set()

    # ---------------- 对位 ----------------
    def _reset_align(self):
        self.transform = None
        self.align_mode = False
        self.align_refs = []
        self.align_pts = []
        self.align_confirmed = [False, False, False]
        for ex, ey in ((self.e_x1, self.e_y1), (self.e_x2, self.e_y2), (self.e_x3, self.e_y3)):
            ex.delete(0, "end")
            ey.delete(0, "end")
        self.lbl_param.config(text="")
        self._align_status("① 各行输入位号后按回车确认（可键盘输入，自动补全）→ ② 点“开始对位” → ③ 移动大图到目标位置，双击图片对应点位即录入 X/Y（单击也可）", "#555555")
        self._render()

    def _align_status(self, text, color="#555555"):
        self.lbl_align.config(text=text, foreground=color)

    def _read_align_entries(self):
        """从 P1/P2/P3 的 X/Y 输入框读取像素坐标；未填齐或非法返回 None。"""
        vals = []
        for ex, ey in ((self.e_x1, self.e_y1), (self.e_x2, self.e_y2), (self.e_x3, self.e_y3)):
            x = ex.get().strip()
            y = ey.get().strip()
            if not x and not y:
                vals.append(None)
                continue
            try:
                vals.append((float(x), float(y)))
            except ValueError:
                return None
        if any(v is None for v in vals):
            return None
        return vals

    def _read_align_partial(self):
        """读取 P1/P2/P3 的 X/Y：返回三元素列表，未填项为 None（部分填齐也可用）。"""
        vals = []
        for ex, ey in ((self.e_x1, self.e_y1), (self.e_x2, self.e_y2), (self.e_x3, self.e_y3)):
            x = ex.get().strip()
            y = ey.get().strip()
            if not x and not y:
                vals.append(None)
                continue
            try:
                vals.append((float(x), float(y)))
            except ValueError:
                vals.append(None)
        return vals

    def _fill_align_entry(self, i, ix, iy):
        """把图片点击坐标回显到第 i 个 P 的输入框。"""
        ex = (self.e_x1, self.e_x2, self.e_x3)[i]
        ey = (self.e_y1, self.e_y2, self.e_y3)[i]
        for w, v in ((ex, ix), (ey, iy)):
            w.delete(0, "end")
            w.insert(0, "%.1f" % v)

    def _cb_autocomplete(self, cb, k=-1):
        """位号输入框自动补全：输入前缀时自动补全为坐标文件中的完整位号。
        精确匹配优先：若坐标文件中存在与输入完全相同的位号（如 D1），保持原样不补全，
        避免被其他前缀相同但顺序靠前的位号（如 D113）覆盖。补全后需回车确认。"""
        if not self.coords:
            return
        txt = cb.get()
        if not txt:
            return
        up = txt.upper()
        # 第一遍：精确匹配存在 → 保持用户输入
        for r in self.coords:
            if r["refdes"].upper() == up:
                return
        # 第二遍：无精确匹配才做前缀补全（第一个命中）
        for r in self.coords:
            ref = r["refdes"]
            if ref.upper().startswith(up) and len(ref) > len(txt):
                cb.set(ref)
                cb.icursor("end")
                try:
                    cb.selection_range(len(txt), "end")   # 高亮补全部分，继续输入可替换
                except Exception:
                    pass
                if 0 <= k < 3:
                    self.align_confirmed[k] = False       # 文本被改动 → 需重新回车确认
                return

    def _confirm_align_ref(self, i):
        """回车 / 下拉选择后确认 P{i+1} 位号：校验存在后标记已确认，提示下一步双击图片。"""
        cb = (self.cb1, self.cb2, self.cb3)[i]
        ref = cb.get().strip()
        if not ref:
            self._align_status(f"请先在 P{i + 1} 行输入位号（回车确认）", COL_WAIT)
            return False
        if self._find_ref(ref) is None:
            messagebox.showinfo("提示", f"位号 {ref} 不在坐标文件中，请检查输入")
            self.align_confirmed[i] = False
            return False
        self.align_confirmed[i] = True
        if self.align_mode:
            self._align_status(f"P{i + 1} 位号 {ref} 已确认 ✓  请双击图片上对应位置确认坐标（单击也可，右键可撤销）", "#1a6fb5")
        return True

    def _fill_combos(self):
        refs = [r["refdes"] for r in self.coords]
        for cb in (self.cb1, self.cb2, self.cb3):
            cb["values"] = refs
            cb.set("")                     # 位号初始为空，由用户自行选择
        self.align_confirmed = [False, False, False]

    def _find_ref(self, refdes):
        up = refdes.strip().upper()
        for r in self.coords:
            if r["refdes"].upper() == up:
                return r
        return None

    def _search_autocomplete(self):
        """搜索框自动补全：输入前缀时自动补全为完整位号。
        精确匹配优先，避免输入 D1 时被 D113 等前缀位号抢先补全。"""
        if not self.coords:
            return
        txt = self.e_search.get().strip()
        if not txt:
            return
        up = txt.upper()
        for r in self.coords:
            if r["refdes"].upper() == up:
                return                       # 已完整输入，不动
        for r in self.coords:
            ref = r["refdes"]
            if ref.upper().startswith(up) and len(ref) > len(txt):
                self.e_search.delete(0, "end")
                self.e_search.insert(0, ref)
                self.e_search.icursor("end")
                try:
                    self.e_search.selection_range(len(txt), "end")
                except Exception:
                    pass
                return

    def search_jump(self):
        """按位号搜索并跳转到对应点位（画面居中 + 表格选中）。"""
        if not self.coords:
            messagebox.showinfo("提示", "请先导入坐标文件")
            return
        txt = self.e_search.get().strip()
        if not txt:
            messagebox.showinfo("提示", "请输入要搜索的位号，如 C210")
            return
        up = txt.upper()
        idx = None
        for i, r in enumerate(self.coords):
            if r["refdes"].upper() == up:
                idx = i
                break
        if idx is None:                        # 未精确命中 → 前缀匹配第一个
            for i, r in enumerate(self.coords):
                if r["refdes"].upper().startswith(up):
                    idx = i
                    break
        if idx is None:
            messagebox.showinfo("提示", f"未找到位号 {txt}")
            return
        self.select(idx)
        self.canvas.focus_set()

    def _search_dialog(self):
        """菜单「编辑 → 搜索位号…」弹出小窗输入位号并跳转。"""
        if not self.coords:
            messagebox.showinfo("提示", "请先导入坐标文件")
            return
        win = tk.Toplevel(self.root)
        win.title("搜索位号")
        win.geometry("320x130")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        win.update_idletasks()
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        win.geometry("+%d+%d" % ((sw - 320) // 2, (sh - 130) // 2))
        ttk.Label(win, text="输入位号（如 C210、R1、U5）：",
                  font=("Microsoft YaHei UI", 10)).pack(pady=(16, 6))
        e = ttk.Entry(win, font=("Consolas", 12), justify="center")
        e.pack(pady=4, padx=20, fill="x")
        e.focus_set()
        def go():
            txt = e.get().strip()
            if not txt:
                return
            self.e_search.delete(0, "end")
            self.e_search.insert(0, txt)
            win.destroy()
            self.search_jump()
        ttk.Button(win, text="定位", command=go).pack(pady=10)
        e.bind("<Return>", lambda ev: go())

    def begin_align(self):
        if self.img is None:
            messagebox.showinfo("提示", "请先打开 PCB 图片")
            return
        if not self.coords:
            messagebox.showinfo("提示", "请先导入坐标文件")
            return
        refs = [cb.get().strip() for cb in (self.cb1, self.cb2, self.cb3)]
        self.align_refs = refs
        self.align_pts = []
        self.align_mode = True
        pts = self._read_align_entries()
        if pts is not None:
            # 手输坐标已填齐：直接对位
            self.align_pts = pts
            self.compute_align()
            return
        self.canvas.focus_set()
        self._align_status("在 P1 行输入位号后按回车确认，移动大图到目标位置，双击图片对应点位确认坐标（单击也可，右键可撤销）", COL_WAIT)
        self._render()

    def undo_align_pt(self):
        if not self.align_mode:
            return
        # 清空最后一个已填写的 P 输入框
        for ex, ey in reversed(((self.e_x1, self.e_y1), (self.e_x2, self.e_y2), (self.e_x3, self.e_y3))):
            if ex.get().strip() or ey.get().strip():
                ex.delete(0, "end")
                ey.delete(0, "end")
                break
        self.align_pts = [p for p in self._read_align_partial() if p]
        k = len(self.align_pts)
        if k == 0:
            self._align_status("在 P1 行输入位号后按回车确认，移动大图到目标位置，双击图片对应点位确认坐标（单击也可，右键可撤销）", COL_WAIT)
        elif k < 3:
            self._align_status(f"接着在 P{k + 1} 行输入位号并按回车确认，再双击图片对应点位（右键可撤销）", COL_WAIT)
        self._render()

    def clear_align(self):
        self.align_mode = False
        self._reset_align()

    def compute_align(self):
        # 位号以输入框当前内容为准（可在填点过程中逐个修改）
        refs = [cb.get().strip() for cb in (self.cb1, self.cb2, self.cb3)]
        if any(not r for r in refs) or len(set(refs)) < 3:
            messagebox.showinfo("提示", "请为 P1/P2/P3 填写 3 个不同的位号（可直接键盘输入，如 C210）")
            return False
        for r in refs:
            if self._find_ref(r) is None:
                messagebox.showinfo("提示", f"位号 {r} 不在坐标文件中，请检查输入")
                return False
        self.align_refs = refs
        pts = self._read_align_entries()
        if pts is None:
            messagebox.showinfo("提示", "请为 P1/P2/P3 填齐 X/Y 坐标（双击图片可自动填入）")
            return False
        self.align_pts = pts
        src, dst = [], []
        for ref, (ix, iy) in zip(self.align_refs, self.align_pts):
            row = self._find_ref(ref)
            if row is None:
                return False
            src.append((row["x"], row["y"]))
            dst.append((ix, iy))
        t = affine_from_3pts(src, dst)
        if t is None:
            messagebox.showerror("错误", "三点过于集中或重合，无法计算变换")
            return False
        # 计算三点回投误差（仿射为精确解，误差应≈0；此处仅作展示与校验）
        err = 0.0
        for (rx, ry), (ix, iy) in zip(src, dst):
            px, py = transform_point(t, rx, ry)
            err = max(err, math.hypot(px - ix, py - iy))
        t["err_px"] = err
        self.transform = t
        self.align_mode = False
        self.align_pts = []
        self.cur = -1
        mirror = "是（镜像面）" if t["mirrored"] else "否"
        self._align_status("对位完成：镜像 %s ｜ 旋转 %.2f° ｜ X 比例 %.3f × Y 比例 %.3f px/mm ｜ 三点误差 %.2f px"
                           % (mirror, t["angle_deg"], t["sx"], t["sy"], t["err_px"]), "#1a6fb5")
        self.select(0)
        self._render()
        self.canvas.focus_set()
        return True

    # ---------------- 画布事件 ----------------
    def _canvas_to_img(self, x, y):
        return ((x - self.ox) / self.zoom, (y - self.oy) / self.zoom)

    def _on_press(self, e):
        if self.align_mode:
            now = time.time()
            if (now - getattr(self, "_last_align_ts", 0.0)) < 0.4:
                return                       # 双击的第二次按下：忽略，防止重复填点
            self._last_align_ts = now
            ix, iy = self._canvas_to_img(e.x, e.y)
            k = self._next_align_row()
            if k < 0:                        # 三行 XY 已齐 → 自动对位
                self.compute_align()
                self._render()
                return
            if not self.align_confirmed[k]:
                self._align_status(f"请先在 P{k + 1} 行输入位号并按回车确认，再双击图片对应点位（右键可撤销）", COL_WAIT)
                self._render()
                return
            self._fill_align_entry(k, ix, iy)
            self.align_pts = [p for p in self._read_align_partial() if p]
            if len(self.align_pts) >= 3:
                self.compute_align()
            else:
                self._align_status(f"P{k + 1} 已录入 ✓  接着在 P{k + 2} 行选好位号，双击图片对应点位（右键可撤销）", COL_WAIT)
            self._render()
            return
        self._pan = (e.x, e.y, self.ox, self.oy, False)

    def _next_align_row(self):
        """返回第一个尚未填 X/Y 的 P 行下标（P1→P2→P3 顺序）；全部填齐返回 -1。"""
        for i, p in enumerate(self._read_align_partial()):
            if p is None:
                return i
        return -1

    def _on_drag(self, e):
        if self._pan:
            x0, y0, ox0, oy0, _ = self._pan
            self.ox = ox0 + (e.x - x0)
            self.oy = oy0 + (e.y - y0)
            self._pan = (x0, y0, self.ox, self.oy, True)
            self._render()

    def _on_release(self, e):
        if self._pan and not self._pan[4]:
            self._click_select(e)
        self._pan = None

    def _click_select(self, e):
        if self.transform is None:
            return
        best, bd = None, 1e18
        for i, row in enumerate(self.coords):
            px, py = transform_point(self.transform, row["x"], row["y"])
            sx, sy = self.ox + px * self.zoom, self.oy + py * self.zoom
            d = math.hypot(sx - e.x, sy - e.y)
            if d < bd:
                bd, best = d, i
        if best is not None and bd <= 30:
            self.select(best)

    def _on_wheel(self, e):
        if self.img is None:
            return
        f = 1.0015 ** e.delta          # Windows 滚轮 delta=±120
        nz = max(0.005, min(self.zoom * f, 300.0))
        ix, iy = self._canvas_to_img(e.x, e.y)
        self.zoom = nz
        self.ox = e.x - ix * self.zoom
        self.oy = e.y - iy * self.zoom
        self._render()

    def _on_right(self, e):
        if self.align_mode:
            self.undo_align_pt()

    def _on_motion(self, e):
        if self.align_mode:
            self.canvas.delete("cursor")
            self.canvas.create_line(e.x - 14, e.y, e.x + 14, e.y,
                                    fill=COL_WAIT, tags=("cursor",))
            self.canvas.create_line(e.x, e.y - 14, e.x, e.y + 14,
                                    fill=COL_WAIT, tags=("cursor",))

    def _on_configure(self):
        if self.img is not None and self.zoom is None:
            self._fit_view()

    def _fit_view(self):
        if self.img is None:
            return
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        if cw <= 1 or ch <= 1:
            return
        w, h = self.img.size
        z = min(cw / w, ch / h) * 0.97
        self.zoom = max(z, 0.005)
        self.ox = (cw - w * self.zoom) / 2
        self.oy = (ch - h * self.zoom) / 2
        self._render()

    def _zoom_center(self, factor):
        """以画布中心为锚点缩放。"""
        if self.img is None or self.zoom is None:
            return
        cw = self.canvas.winfo_width() / 2
        ch = self.canvas.winfo_height() / 2
        # 屏幕中心对应的图片坐标
        ix = (cw - self.ox) / self.zoom
        iy = (ch - self.oy) / self.zoom
        self.zoom = max(0.001, self.zoom * factor)
        self.ox = cw - ix * self.zoom
        self.oy = ch - iy * self.zoom
        self._render()

    def _reset_view(self):
        self.zoom = None
        self.ox = self.oy = 0.0

    # ---------------- 导航与确认 ----------------
    def select(self, idx):
        if not self.coords:
            return
        n = len(self.coords)
        idx = max(0, min(idx, n - 1))
        if self.cur != idx:
            self._save_remark_internal()
        self.cur = idx
        row = self.coords[idx]
        self.entry_remark.delete(0, "end")
        self.entry_remark.insert(0, row["remark"] or "")
        self.lbl_cur.config(text=row["refdes"])
        self.lbl_info.config(text="第 %d / %d 个    X=%.3f  Y=%.3f  角度 %.0f°"
                                  % (idx + 1, n, row["x"], row["y"], row["rot"]))
        self._refresh_status()
        self._center_on_current()   # 画面中心移到选中点位（整板视野时自动放大到可辨级别）
        self._render()
        # 表格跟随（防重入：避免 selection_set 与 <<TreeviewSelect>> 互相触发）
        iid = self.tree.get_children()[idx]
        self._loading_tree = True
        self.tree.see(iid)
        cur_sel = self.tree.selection()
        if not cur_sel or cur_sel[0] != iid:
            self.tree.selection_set(iid)
        self._loading_tree = False

    def _center_on_current(self):
        """平移视口，使当前选中点位于画面正中心；若仍处于整板视野级别则先放大。"""
        if self.transform is None or self.cur < 0 or self.img is None:
            return
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        if cw <= 1 or ch <= 1:
            return
        if self.zoom is None:
            self._fit_view()
            return
        row = self.coords[self.cur]
        px, py = transform_point(self.transform, row["x"], row["y"])
        fit = min(cw / self.img.width, ch / self.img.height)
        if self.zoom <= fit * 1.15:           # 处于整板视野级别 → 放大 6 倍便于目视
            self.zoom = max(self.zoom, fit * 6.0)
        self.ox = cw / 2 - px * self.zoom
        self.oy = ch / 2 - py * self.zoom

    def _auto_ok_if_pending(self):
        """方向键跳转时：当前点位若尚未判定，视为已目视检查过 → 自动判合格（OK）。"""
        if 0 <= self.cur < len(self.coords):
            row = self.coords[self.cur]
            if row["result"] is None:
                row["result"] = "OK"
                row["time"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self._save_remark_internal()
                self._update_table_row(self.cur)
                self._refresh_status()

    def prev(self):
        if self.coords and self.cur > 0:
            self._auto_ok_if_pending()
            self.select(self.cur - 1)

    def next_(self):
        if self.coords and self.cur < len(self.coords) - 1:
            self._auto_ok_if_pending()
            self.select(self.cur + 1)

    def mark(self, ok):
        if self.transform is None:
            messagebox.showinfo("提示", "请先完成三点对位")
            return
        if self.cur < 0:
            return
        row = self.coords[self.cur]
        row["result"] = "OK" if ok else "NG"
        row["time"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._save_remark_internal()
        self._update_table_row(self.cur)
        self._refresh_status()
        self._render()
        self.canvas.focus_set()
        if self.cur < len(self.coords) - 1:
            self.select(self.cur + 1)

    def save_remark(self):
        self._save_remark_internal()
        self._update_table_row(self.cur)
        self.canvas.focus_set()

    def _save_remark_internal(self):
        if 0 <= self.cur < len(self.coords):
            self.coords[self.cur]["remark"] = self.entry_remark.get().strip()

    def _key_return(self, e):
        if not self._key_guard(("entry", "button")):
            return
        self.mark(True)

    def _key_space(self, e):
        if not self._key_guard(("entry", "button")):
            return
        self.mark(False)

    def _key_escape(self, e):
        if self.align_mode:
            self.clear_align()
        elif self.cur >= 0:
            self.cur = -1
            self._render()

    def _key_guard(self, kinds):
        """焦点守卫：焦点在输入控件/按钮/表格时，快捷键让位于控件本身，
        避免空格/回车/方向键被“控件默认行为 + 顶层快捷键”双重触发。"""
        w = self.root.focus_get()
        if w is None:
            return True
        entry = isinstance(w, (tk.Entry, ttk.Entry, ttk.Spinbox))
        button = isinstance(w, (tk.Button, ttk.Button))
        tree = isinstance(w, ttk.Treeview)
        if "entry" in kinds and entry:
            return False
        if "button" in kinds and button:
            return False
        if "tree" in kinds and tree:
            return False
        return True

    def _tree_select(self, e):
        if self._loading_tree:
            return
        sel = self.tree.selection()
        if sel:
            idx = int(sel[0])
            if idx == self.cur:
                return
            self.select(idx)

    # ---------------- 渲染 ----------------
    def _render(self):
        if self.img is None:
            return
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        if cw <= 1 or ch <= 1:
            return
        if self.zoom is None:
            self._fit_view()
            return
        w, h = self.img.size
        x0 = max(0, int(math.floor(-self.ox / self.zoom)))
        y0 = max(0, int(math.floor(-self.oy / self.zoom)))
        x1 = min(w, int(math.ceil((cw - self.ox) / self.zoom)))
        y1 = min(h, int(math.ceil((ch - self.oy) / self.zoom)))
        if x1 <= x0 or y1 <= y0:
            return
        disp_w = int(round((x1 - x0) * self.zoom))
        disp_h = int(round((y1 - y0) * self.zoom))
        if max(disp_w, disp_h) > 20000:            # 防止超大内存分配
            self.zoom *= 20000 / max(disp_w, disp_h)
            self._render()
            return
        if disp_w <= 0 or disp_h <= 0:
            return
        crop = self.img.crop((x0, y0, x1, y1))
        if disp_w < (x1 - x0):                      # 缩小用高质量
            crop = crop.resize((disp_w, disp_h), Image.Resampling.LANCZOS)
        else:                                       # 放大用快速
            crop = crop.resize((disp_w, disp_h), Image.Resampling.BILINEAR)
        # 亮度 / 对比度调整（仅作用于显示，不改原始图与坐标）
        _b = self.v_bright.get()
        _c = self.v_contrast.get()
        if _c != 0:
            crop = ImageEnhance.Contrast(crop).enhance(1.0 + _c / 100.0 * 0.8)
        if _b != 0:
            crop = ImageEnhance.Brightness(crop).enhance(1.0 + _b / 100.0 * 0.8)
        self._photo = ImageTk.PhotoImage(crop)
        self.canvas.delete("all")
        self.canvas.create_image(self.ox + x0 * self.zoom, self.oy + y0 * self.zoom,
                                 anchor="nw", image=self._photo, tags=("img",))
        self._draw_overlays()
        self._update_minimap()

    def _text_outline(self, cx, cy, text, color, font, tags=("ovl",)):
        for dx, dy in ((-1, 0), (1, 0), (0, -1), (0, 1)):
            self.canvas.create_text(cx + dx, cy + dy, text=text, fill="#000000",
                                    font=font, tags=tags)
        self.canvas.create_text(cx, cy, text=text, fill=color, font=font, tags=tags)

    def _draw_overlays(self):
        t = self.transform
        if t is None:
            self._draw_align_marks()
            return
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        vx0, vy0 = -self.ox / self.zoom, -self.oy / self.zoom
        vx1, vy1 = vx0 + cw / self.zoom, vy0 + ch / self.zoom
        L = min(26.0, max(6.0, self.zoom * 1.6))
        show_cross = self.v_cross.get()
        show_box   = self.v_box.get()
        show_label = self.v_label.get()
        show_rot   = self.v_rot.get()
        show_done  = self.v_done.get()
        show_all   = self.v_all.get()

        order = list(range(len(self.coords)))
        if self.cur >= 0:
            order = [i for i in order if i != self.cur] + [self.cur]   # 当前最后画

        for i in order:
            row = self.coords[i]
            px, py = transform_point(t, row["x"], row["y"])
            if px < vx0 - 100 or px > vx1 + 100 or py < vy0 - 100 or py > vy1 + 100:
                continue
            sx, sy = self.ox + px * self.zoom, self.oy + py * self.zoom
            is_cur = (i == self.cur)
            if not is_cur and not show_all:
                continue
            res = row["result"]
            if not is_cur and res is None:
                # “显示全部点位”模式下，未检元件画小圆点定位标记
                self.canvas.create_oval(sx - 3, sy - 3, sx + 3, sy + 3,
                                        fill=COL_PEND, outline="", tags=("ovl",))
            if (show_cross or show_box) and (is_cur or show_all):
                if is_cur:
                    Lc = L * 5.0
                    thin = 1.2
                    r = max(30.0, L * 1.3 * 3.0)
                    a = r * 0.55
                    wr = 1.2
                    if show_cross:
                        # 长十字线：方框打开时在方框内剪断，否则贯穿中心
                        if show_box:
                            self.canvas.create_line(sx - Lc, sy, sx - r, sy,
                                                    fill=COL_CUR, width=thin, tags=("ovl",))
                            self.canvas.create_line(sx + r, sy, sx + Lc, sy,
                                                    fill=COL_CUR, width=thin, tags=("ovl",))
                            self.canvas.create_line(sx, sy - Lc, sx, sy - r,
                                                    fill=COL_CUR, width=thin, tags=("ovl",))
                            self.canvas.create_line(sx, sy + r, sx, sy + Lc,
                                                    fill=COL_CUR, width=thin, tags=("ovl",))
                        else:
                            self.canvas.create_line(sx - Lc, sy, sx + Lc, sy,
                                                    fill=COL_CUR, width=thin, tags=("ovl",))
                            self.canvas.create_line(sx, sy - Lc, sx, sy + Lc,
                                                    fill=COL_CUR, width=thin, tags=("ovl",))
                    if show_box:
                        self.canvas.create_line(sx - r, sy - r + a, sx - r, sy - r,
                                                fill=COL_CUR, width=wr, tags=("ovl",))
                        self.canvas.create_line(sx - r, sy - r, sx - r + a, sy - r,
                                                fill=COL_CUR, width=wr, tags=("ovl",))
                        self.canvas.create_line(sx + r - a, sy - r, sx + r, sy - r,
                                                fill=COL_CUR, width=wr, tags=("ovl",))
                        self.canvas.create_line(sx + r, sy - r, sx + r, sy - r + a,
                                                fill=COL_CUR, width=wr, tags=("ovl",))
                        self.canvas.create_line(sx - r, sy + r - a, sx - r, sy + r,
                                                fill=COL_CUR, width=wr, tags=("ovl",))
                        self.canvas.create_line(sx - r, sy + r, sx - r + a, sy + r,
                                                fill=COL_CUR, width=wr, tags=("ovl",))
                        self.canvas.create_line(sx + r - a, sy + r, sx + r, sy + r,
                                                fill=COL_CUR, width=wr, tags=("ovl",))
                        self.canvas.create_line(sx + r, sy + r, sx + r, sy + r - a,
                                                fill=COL_CUR, width=wr, tags=("ovl",))
                else:
                    if show_cross:
                        self.canvas.create_line(sx - L, sy, sx + L, sy,
                                                fill=COL_PEND, width=1, tags=("ovl",))
                        self.canvas.create_line(sx, sy - L, sx, sy + L,
                                                fill=COL_PEND, width=1, tags=("ovl",))
            if show_rot and (is_cur or show_all):
                rad = math.radians(row["rot"])
                ux, uy = t["a"] * math.cos(rad) + t["b"] * math.sin(rad), \
                         t["d"] * math.cos(rad) + t["e"] * math.sin(rad)
                rl = L + 5
                ex, ey = sx + ux * rl, sy + uy * rl
                self.canvas.create_line(sx, sy, ex, ey, fill="#39c5ff", width=2, tags=("ovl",))
                for d in (0.9, 1.1):
                    ang = math.atan2(ey - sy, ex - sx)
                    hx = ex + 5 * math.cos(ang + d)
                    hy = ey + 5 * math.sin(ang + d)
                    self.canvas.create_line(ex, ey, hx, hy, fill="#39c5ff", width=2, tags=("ovl",))
            if show_label and (is_cur or show_all):
                self._text_outline(sx + L + 5, sy - L - 5, row["refdes"],
                                   COL_CUR if is_cur else COL_LABEL,
                                   ("Microsoft YaHei UI", 11 if is_cur else 9, "bold" if is_cur else "normal"))
            if show_done and res:
                m = L + 5
                if res == "OK":
                    self.canvas.create_line(sx + m - 6, sy + m, sx + m - 2, sy + m + 5, fill=COL_OK, width=2.5, tags=("ovl",))
                    self.canvas.create_line(sx + m - 2, sy + m + 5, sx + m + 6, sy + m - 5, fill=COL_OK, width=2.5, tags=("ovl",))
                else:
                    self.canvas.create_line(sx + m - 5, sy + m - 5, sx + m + 5, sy + m + 5, fill=COL_NG, width=2.5, tags=("ovl",))
                    self.canvas.create_line(sx + m - 5, sy + m + 5, sx + m + 5, sy + m - 5, fill=COL_NG, width=2.5, tags=("ovl",))
        self._draw_align_marks()

    def _draw_align_marks(self):
        """对位模式下的已选点标记。"""
        for k, (ref, (ix, iy)) in enumerate(zip(self.align_refs, self.align_pts)):
            sx, sy = self.ox + ix * self.zoom, self.oy + iy * self.zoom
            self.canvas.create_oval(sx - 10, sy - 10, sx + 10, sy + 10,
                                    outline=COL_WAIT, width=2, tags=("ovl",))
            self._text_outline(sx + 12, sy - 12, f"P{k + 1} {ref}", COL_WAIT,
                               ("Microsoft YaHei UI", 11, "bold"), ("ovl",))

    # ---------------- 全景图 ----------------
    def _make_minimap(self):
        if self.img is None:
            return
        w, h = self.img.size
        self.mm_scale = min(200.0 / w, 120.0 / h)
        self.mm_ox = (200 - w * self.mm_scale) / 2
        self.mm_oy = (120 - h * self.mm_scale) / 2
        thumb = self.img.resize((max(1, int(w * self.mm_scale)),
                                 max(1, int(h * self.mm_scale))),
                                Image.Resampling.LANCZOS)
        self.mm_img = thumb
        self.mm_photo = ImageTk.PhotoImage(thumb)

    def _update_minimap(self):
        if self.mm_photo is None:
            return
        self.mm_canvas.delete("all")
        self.mm_canvas.create_image(self.mm_ox, self.mm_oy, anchor="nw",
                                    image=self.mm_photo)
        s = self.mm_scale
        # 当前视口
        vx0, vy0 = -self.ox / self.zoom, -self.oy / self.zoom
        vx1 = vx0 + self.canvas.winfo_width() / self.zoom
        vy1 = vy0 + self.canvas.winfo_height() / self.zoom
        self.mm_canvas.create_rectangle(self.mm_ox + vx0 * s, self.mm_oy + vy0 * s,
                                        self.mm_ox + vx1 * s, self.mm_oy + vy1 * s,
                                        outline="#39c5ff", width=1)
        if self.transform is not None and self.cur >= 0:
            row = self.coords[self.cur]
            px, py = transform_point(self.transform, row["x"], row["y"])
            cx, cy = self.mm_ox + px * s, self.mm_oy + py * s
            self.mm_canvas.create_oval(cx - 4, cy - 4, cx + 4, cy + 4,
                                       fill=COL_CUR, outline="")

    def _mm_click(self, e):
        if self.img is None or self.zoom is None:
            return
        ix = (e.x - self.mm_ox) / self.mm_scale
        iy = (e.y - self.mm_oy) / self.mm_scale
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        self.ox = cw / 2 - ix * self.zoom
        self.oy = ch / 2 - iy * self.zoom
        self._render()

    # ---------------- 表格 / 状态 ----------------
    def _fill_table(self):
        self._loading_tree = True
        for it in self.tree.get_children():
            self.tree.delete(it)
        for i, row in enumerate(self.coords):
            self.tree.insert("", "end", iid=str(i), values=(
                row["no"], row["refdes"],
                "%.3f" % row["x"], "%.3f" % row["y"],
                "%.0f" % row["rot"], row["mir"],
                row["result"] or "", row["remark"] or ""),
                tags=(row["result"] or "none",))
        self._loading_tree = False

    def _update_table_row(self, i):
        if i < 0 or i >= len(self.coords):
            return
        row = self.coords[i]
        iid = self.tree.get_children()[i]
        self.tree.item(iid, values=(
            row["no"], row["refdes"], "%.3f" % row["x"], "%.3f" % row["y"],
            "%.0f" % row["rot"], row["mir"], row["result"] or "", row["remark"] or ""),
            tags=(row["result"] or "none",))

    def _refresh_status(self):
        n = len(self.coords)
        ok = sum(1 for r in self.coords if r["result"] == "OK")
        ng = sum(1 for r in self.coords if r["result"] == "NG")
        if n:
            self.lbl_prog.config(text="已确认 %d / %d   （OK %d ｜ NG %d）" % (ok + ng, n, ok, ng))
            self.prog["value"] = (ok + ng) / n * 100
        else:
            self.lbl_prog.config(text="")
            self.prog["value"] = 0

    # ---------------- 导出 ----------------
    def export_xlsx(self):
        """导出首检结果为 XLSX；判定 NG 的点位自动截取图片局部区域插入对应行。"""
        if not self.coords:
            messagebox.showinfo("提示", "尚未导入坐标")
            return
        base = os.path.splitext(os.path.basename(self.coord_path or "首检结果"))[0]
        path = filedialog.asksaveasfilename(
            title="导出首检结果（XLSX，NG 点位自动附局部截图）",
            defaultextension=".xlsx",
            initialfile=base + "_首检结果.xlsx",
            filetypes=[("Excel 文件", "*.xlsx")])
        if not path:
            return
        try:
            import tempfile as _tf
            from openpyxl import Workbook
            from openpyxl.styles import Font, Alignment, PatternFill
            from openpyxl.drawing.image import Image as XLImage

            wb = Workbook()
            ws = wb.active
            ws.title = "首检结果"
            headers = ["序号", "位号", "X(mm)", "Y(mm)", "角度(°)", "面别",
                       "结果", "NG局部截图", "备注", "确认时间"]
            ws.append(headers)
            hfill = PatternFill("solid", fgColor="D9E1F2")
            for c in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=c)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill = hfill
            widths = (8, 12, 11, 11, 9, 8, 8, 30, 22, 18)
            for i, wdt in enumerate(widths, start=1):
                ws.column_dimensions[chr(64 + i)].width = wdt

            ng_count = 0
            for i, r in enumerate(self.coords, start=2):
                ws.append([r["no"], r["refdes"], round(r["x"], 4), round(r["y"], 4),
                           "%.0f" % r["rot"], r["mir"], r["result"] or "", "",
                           r["remark"] or "", r["time"] or ""])
                if r["result"] == "NG":
                    ng_count += 1
                    res_cell = ws.cell(row=i, column=7)
                    res_cell.font = Font(bold=True, color="D1242F")
                    # 截取 NG 点位图片局部区域
                    if self.transform is not None and self.img is not None:
                        px, py = transform_point(self.transform, r["x"], r["y"])
                        half = 450   # 截图半径（225×2，再扩大1倍以容纳大器件）
                        x0 = max(0, int(px - half))
                        y0 = max(0, int(py - half))
                        x1 = min(self.img.width, int(px + half))
                        y1 = min(self.img.height, int(py + half))
                        if x1 > x0 and y1 > y0:
                            crop = self.img.crop((x0, y0, x1, y1))
                            wdt = 280   # 显示宽度同步加大，避免大图被缩得太小看不清
                            hgt = max(1, int(crop.height * wdt / crop.width))
                            crop = crop.resize((wdt, hgt), Image.Resampling.LANCZOS)
                            tmp = os.path.join(_tf.gettempdir(), "ng_shot_%d_%s.png"
                                               % (i, r["refdes"].replace("/", "_")))
                            crop.save(tmp, "PNG")
                            xi = XLImage(tmp)
                            xi.width = wdt
                            xi.height = hgt
                            ws.add_image(xi, "H%d" % i)
                            ws.row_dimensions[i].height = hgt * 0.78
                    else:
                        ws.cell(row=i, column=8, value="(未对位，无截图)")
            wb.save(path)
        except Exception as ex:
            messagebox.showerror("错误", f"导出失败：\n{ex}")
            return
        msg = "已保存：\n%s" % path
        if ng_count:
            msg += "\n\n共 %d 个 NG 点位，局部截图已插入对应行。" % ng_count
        messagebox.showinfo("导出成功", msg)

    # ---------------- 操作说明 ----------------
    def show_help(self):
        """弹出操作说明窗口（参考案例样式：标题 + 滚动文本区）。"""
        win = tk.Toplevel(self.root)
        win.title("操作说明 - PCBA 首检确认工具")
        win.geometry("780x640")
        txt = tk.Text(win, wrap="word", font=("Microsoft YaHei UI", 11), padx=14, pady=12)
        sb = ttk.Scrollbar(win, command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        txt.pack(side="left", fill="both", expand=True)
        content = HELP_TEXT
        txt.insert("1.0", content)
        txt.configure(state="disabled")

    # ---------------- 项目导入 / 导出（图片+坐标+对位参数+进度） ----------------
    def export_project(self):
        """把图片、坐标、三点对位参数与检查进度打包为单个 .pba 项目文件，
        之后用「导入项目」可一键恢复，无需重新对位。"""
        if self.img is None:
            messagebox.showinfo("提示", "请先打开 PCB 图片")
            return
        if not self.coords:
            messagebox.showinfo("提示", "请先导入坐标文件")
            return
        base = os.path.splitext(os.path.basename(self.coord_path or "首检项目"))[0]
        path = filedialog.asksaveasfilename(
            title="导出首检项目（含图片与对位参数）",
            defaultextension=".pba",
            initialfile=base + "_首检项目.pba",
            filetypes=[("PCBA 首检项目", "*.pba")])
        if not path:
            return
        try:
            payload = dict(
                version=2,
                img_name=os.path.basename(self.img_path or "image.png"),
                coord_name=os.path.basename(self.coord_path or ""),
                cur=self.cur if self.cur >= 0 else 0,
                coords=[dict(r) for r in self.coords],
                transform=self.transform,
                align_refs=list(self.align_refs),
                align_pts=[list(p) for p in self.align_pts],
            )
            ext = os.path.splitext(payload["img_name"])[1].lower().lstrip(".")
            fmt_map = {"jpg": "JPEG", "jpeg": "JPEG", "png": "PNG",
                       "webp": "WEBP", "bmp": "PNG", "tif": "PNG",
                       "tiff": "PNG", "": "PNG"}
            fmt = fmt_map.get(ext, "PNG")
            save_ext = "png" if fmt == "PNG" else ext
            buf = io.BytesIO()
            self.img.save(buf, format=fmt)
            data = buf.getvalue()
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
                z.writestr("project.json",
                           json.dumps(payload, ensure_ascii=False, default=str))
                z.writestr("image." + save_ext, data)
        except Exception as ex:
            messagebox.showerror("错误", f"导出项目失败：\n{ex}")
            return
        messagebox.showinfo("导出成功",
                            "项目已导出（含图片、坐标、对位参数与检查进度）：\n%s\n\n"
                            "之后用「导入项目」即可直接恢复对位继续检查。" % path)

    def import_project(self):
        """导入 .pba 项目文件：图片、坐标、三点对位参数与检查进度一键恢复。"""
        path = filedialog.askopenfilename(
            title="导入 PCBA 首检项目",
            filetypes=[("PCBA 首检项目", "*.pba"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            with zipfile.ZipFile(path, "r") as z:
                names = set(z.namelist())
                if "project.json" not in names:
                    raise ValueError("项目文件缺少 project.json")
                payload = json.loads(z.read("project.json").decode("utf-8"))
                img_name = payload.get("img_name", "image.png")
                img_bytes = None
                for cand in (img_name, "image.png", "image.jpg", "image.jpeg"):
                    if cand in names:
                        img_bytes = z.read(cand)
                        img_name = cand
                        break
                if img_bytes is None:
                    raise ValueError("项目文件中未找到图片数据")
            img = Image.open(io.BytesIO(img_bytes))
            img = ImageOps.exif_transpose(img)
            img = img.convert("RGB")
            rows = [dict(r) for r in payload["coords"]]
            if not rows:
                raise ValueError("项目文件中无坐标数据")
            t = payload.get("transform")
            cur = payload.get("cur", 0)
        except Exception as ex:
            messagebox.showerror("错误", f"导入项目失败：\n{ex}")
            return

        self.img = img
        self.img_path = os.path.join(os.path.dirname(path), img_name)
        self.coords = rows
        self.coord_path = path
        self.transform = t
        self.cur = -1
        self.align_mode = False
        self.align_refs = list(payload.get("align_refs") or [])
        self.align_pts = [tuple(p) for p in (payload.get("align_pts") or [])]
        for ex, ey in ((self.e_x1, self.e_y1), (self.e_x2, self.e_y2), (self.e_x3, self.e_y3)):
            ex.delete(0, "end")
            ey.delete(0, "end")
        self._fill_combos()
        self._fill_table()
        self.lbl_img.config(text="图片：" + img_name + f"   ({img.width}×{img.height})")
        self.lbl_crd.config(text="坐标：%s   （%d 个元件）" % (os.path.basename(path), len(rows)))
        self._make_minimap()
        self._reset_view()
        self._refresh_status()
        if t is not None:
            mirror = "是（镜像面）" if t.get("mirrored") else "否"
            self.lbl_param.config(text="对位参数已恢复：镜像 %s ｜ 旋转 %.2f° ｜ X 比例 %.3f × Y 比例 %.3f px/mm"
                                       % (mirror, t.get("angle_deg", 0), t.get("sx", 0), t.get("sy", 0)))
            self._align_status("项目已导入：图片 + 坐标 + 对位参数 + 检查进度已恢复，可直接开始检查（↑/↓ 或 Enter/空格）", "#1a6fb5")
            self.select(cur if 0 <= cur < len(rows) else 0)
        else:
            self._align_status("项目已导入（未包含对位参数），请重新三点对位", "#555555")
            self._fit_view()
        self.canvas.focus_set()


def main():
    root = tk.Tk()
    # 把根窗口位置/大小对齐登录窗（460x420 居中），让空白"tk"小窗正好被登录窗压住
    root.update_idletasks()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry("460x340+%d+%d" % ((sw - 460) // 2, (sh - 340) // 2))
    # 启动登录：未登录直接退出
    if not login_dialog(root):
        root.destroy()
        return
    App(root)
    root.state("zoomed")   # 主程序窗口直接最大化
    root.mainloop()


# ---------------- 用户登录 / 权限 ----------------
DEFAULT_USERS = {
    "admin": {"pwd": "123456", "role": "admin"},
    "guest": {"pwd": "0000",   "role": "user"},
}

def _users_path():
    import sys
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    # 写 exe 同目录（开发态写脚本目录）
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    return os.path.join(base, "pcba_users.json")

def load_users():
    p = _users_path()
    if not os.path.exists(p):
        return dict(DEFAULT_USERS)
    try:
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return dict(DEFAULT_USERS)

def save_users(users):
    try:
        with open(_users_path(), "w", encoding="utf-8") as f:
            json.dump(users, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def login_dialog(root):
    """模态登录窗；返回 True 登录成功，False 取消/退出。"""
    import sys
    users = load_users()
    win = tk.Toplevel(root)
    win.title("用户登录")
    win.geometry("460x340")
    win.resizable(False, False)
    win.transient(root)
    win.grab_set()
    # 登录窗图标与主程序一致
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    icon_png = os.path.join(base, "app_icon.png")
    if os.path.exists(icon_png):
        _lg_icon = tk.PhotoImage(file=icon_png)
        win.iconphoto(True, _lg_icon)
    # 居中
    win.update_idletasks()
    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    win.geometry("+%d+%d" % ((sw - 460) // 2, (sh - 340) // 2))

    tk.Label(win, text="PCBA 首检确认工具", font=("Microsoft YaHei UI", 16, "bold"),
             fg="#1a6fb5").pack(pady=(24, 4))
    tk.Label(win, text="First Article Inspection (FAI)", font=("Microsoft YaHei UI", 10),
             fg="#888").pack(pady=(0, 14))

    box = tk.Frame(win)
    box.pack(pady=4)
    tk.Label(box, text="用户:", font=FONT).grid(row=0, column=0, padx=6, pady=8, sticky="e")
    cb_user = ttk.Combobox(box, width=22, font=FONT, values=list(users.keys()))
    cb_user.set("admin")
    cb_user.grid(row=0, column=1, padx=6, pady=8)
    tk.Label(box, text="密码:", font=FONT).grid(row=1, column=0, padx=6, pady=8, sticky="e")
    e_pwd = ttk.Entry(box, width=24, font=FONT, show="*")
    e_pwd.grid(row=1, column=1, padx=6, pady=8)

    result = {"ok": False}
    def do_login():
        u = cb_user.get().strip()
        p = e_pwd.get()
        if u in users and users[u].get("pwd") == p:
            result["ok"] = True
            result["user"] = u
            result["role"] = users[u].get("role", "user")
            win.destroy()
        else:
            messagebox.showerror("登录失败", "用户名或密码错误", parent=win)
            e_pwd.delete(0, "end")
            e_pwd.focus_set()
    def do_cancel():
        win.destroy()

    btns = tk.Frame(win)
    btns.pack(pady=24)
    ttk.Button(btns, text="确认(O)", command=do_login).pack(side="left", padx=8)
    ttk.Button(btns, text="取消(C)", command=do_cancel).pack(side="left", padx=8)

    e_pwd.bind("<Return>", lambda e: do_login())
    cb_user.bind("<Return>", lambda e: e_pwd.focus_set())
    cb_user.focus_set()
    root.wait_window(win)
    if result["ok"]:
        App.current_user = result["user"]
        App.current_role = result["role"]
    return result["ok"]


if __name__ == "__main__":
    main()
