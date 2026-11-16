# -*- coding: utf-8 -*-
"""
实用工具 v1.4.3  (单文件 GUI)
================================
整合自多个独立小程序，统一在一个带图形界面的软件里：
  * FolderImg2Word   —— 文件夹图片 -> Word（A4 / 2x2 / 7x11cm）
  * Excel图片导出Word —— Excel 一行多图 + M列备注 -> Word
  * Word 批量替换     —— 来自「E:\\docx批量替换」批量查找/替换工具
                         （保留原格式 / 正则 / 备份 / 页眉页脚表格 / 子文件夹 / 可停止）
  * 批量打印（PDF）   —— 内置，调用系统自带 PDF 打印能力
  * PDF工具           —— 跳转 PDF24 在线工具

内置「版本更新」：
  - 顶部常驻版本号 + 「检查更新」按钮；每次启动后台静默检查新版本
  - 多镜像并发取最高版本、断点续传分段下载、MZ/大小/SHA256 三重校验
  - 发现新版本用非模态横幅提示（立即更新 / 跳过此版本 / 稍后再说）
  - 原地替换运行中的 exe（失败必还原、绝不退出），装完问「立即重启 / 稍后重启」

布局（v1.2 新增）：
  - 左侧导航：功能列表；中间主区：功能名 + 功能介绍 + 配置控件
  - 右侧栏：「📦 成品」与「🕘 历史记录」两个列表，双击即可打开对应成品
  - 底部：进度条 / 日志 / 状态栏

设计目标（用户要求）：
  - 单个 .py 文件，带图形界面，软件名「实用工具」
  - 进度与最终执行结果都显示在这个软件主界面上
  - 可扩展：后续加功能只需新增一个 Tool 子类，并在其中填好 name/desc/intro，
    框架会在每个功能页顶部自动放一个「功能介绍」按钮，点开弹窗显示 intro，
    无需改动其它代码

运行：
  py 实用工具.py               启动图形界面
  py 实用工具.py --selftest    跑核心逻辑自测（不弹界面）
依赖：tkinter(自带) / Pillow / python-docx / openpyxl(Excel版兜底用)
"""
import os
import re
import io
import sys
import html
import shutil
import webbrowser
import zipfile
import tempfile
import threading
import traceback
import subprocess
import hashlib
import json
import time
import urllib.request
import urllib.error
from datetime import datetime

from PIL import Image
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ============================================================
# 通用常量
# ============================================================
IMG_W_CM = 7.0        # 导出图片固定宽度 cm
IMG_H_CM = 11.0       # 导出图片固定高度 cm
DEFAULT_COLS = 2      # 网格列数（2 列 -> 2x2 = 4 张/A4）
IMG_EXTS = ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp', '.tif', '.tiff')

# 打印机列表缓存：None=尚未获取；[]/list=已获取结果（避免每次切换功能都卡一下）
_PRINTERS = None

# ============================================================
# 内置「版本更新」模块（用户规范：每个新电脑软件默认必须带）
# 设计要点：
#  - 清单直链优先（仓库框只作兜底），多镜像并发请求、按版本号取最高、每条带 ?_t= 破 CDN 缓存
#  - 不使用 jsDelivr（@main 会锁 ~12h 前的 commit，导致误报已是最新）
#  - 分段并发下载 + 断点续传 + 镜像轮换 + MZ/size/SHA256 三重校验
#  - 原地替换运行中的 exe（Windows 允许重命名运行中的 exe），失败必还原、绝不退出
#  - PyInstaller 单文件重启必剥离 _MEIPASS2/_PYI_APPLICATION_HOME_DIR/_PYI_ARCHIVE_FILE 并设
#    PYINSTALLER_RESET_ENVIRONMENT=1，避免两进程互删 base_library.zip
# ============================================================
APP_VERSION = '1.4.5'
GITHUB_OWNER = 'H-ZHOWEI'
GITHUB_REPO = 'shiyong-gongju'
UPDATE_MANIFEST_PRIMARY = ('https://raw.githubusercontent.com/%s/%s/main/version.json'
                           % (GITHUB_OWNER, GITHUB_REPO))
# exe 下载走镜像轮换（raw / ghproxy.net / gh-proxy.com / github.com/raw）
_EXE_MIRROR_PREFIX = ['https://ghproxy.net/', 'https://gh-proxy.com/']
_UPDATE_SKIP_NETWORK = False   # 自测可置 True 跳过网络

# ============================================================
# 全局字体管理（命名字体 -> 切换工具栏“字体大小”可实时缩放全部界面）
#  - 用 Tk 命名字体（named font），所有控件以名字引用；
#    重新配置字体后，引用它的控件会即时更新，无需重建界面。
#  - 新增功能若需固定字号，直接用对应的字体名字符串即可，已自动纳入缩放。
# ============================================================
FONT_ROLES = {
    # 名字        (字体族,       基准字号, 字重,      下划线)
    'AppTitle':    ('微软雅黑', 18, 'bold',   0),     # 顶部软件标题
    'AppBig':      ('微软雅黑', 14, 'bold',   0),     # 功能/章节标题
    'AppNavTitle': ('微软雅黑', 11, 'bold',   0),     # 侧栏“功能列表”等小标题
    'AppIntroTitle': ('微软雅黑', 13, 'bold', 0),     # 功能介绍弹窗标题
    'AppBold':     ('微软雅黑', 10, 'bold',   0),     # 强调文字 / 主按钮
    'AppBold12':   ('微软雅黑', 12, 'bold',   0),     # 更新完成提示
    'AppBase':     ('微软雅黑', 10, 'normal', 0),     # 默认正文字体
    'AppSmall':    ('微软雅黑', 9,  'normal', 0),     # 提示 / 次要文字
    'AppIntro':    ('微软雅黑', 11, 'normal', 0),     # 介绍正文
    'AppMono':     ('Consolas', 9,  'normal', 0),     # 日志 / 代码示例
    'AppLink':     ('微软雅黑', 9,  'normal', 1),     # 超链接
}
_FONT_SCALE = 1.0   # 当前缩放倍率（菜单可改）


def setup_fonts(scale=None):
    """创建或更新全部命名字体。scale 为相对于基准的倍率。

    用 Tk 原生 `font create`/`font configure` 直接注册命名字体（比 tkinter.font
    的 Font 对象更确定，且所有控件以名字字符串引用后，重新配置即全局即时生效）。
    """
    global _FONT_SCALE
    if scale is not None:
        _FONT_SCALE = scale
    import tkinter as _tk
    root = _tk._default_root
    if root is None:
        return
    tkcall = root.tk.call
    try:
        names = set(tkcall('font', 'names'))
    except Exception:
        names = set()
    for name, (fam, base, weight, under) in FONT_ROLES.items():
        size = max(7, round(base * _FONT_SCALE))
        args = ('-family', fam, '-size', size, '-weight', weight, '-underline', under)
        if name in names:
            tkcall('font', 'configure', name, *args)
        else:
            try:
                tkcall('font', 'create', name, *args)
            except _tk.TclError:
                tkcall('font', 'configure', name, *args)


def _font(name):
    """返回命名字体的名字字符串（供需要传对象而非名字字符串的场合）。"""
    return name


def _app_dir():
    """运行时所在目录：打包后取 exe 目录，脚本模式取 .py 目录。"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _icon_path():
    """图标文件路径：打包后从 PyInstaller 解包目录取，脚本模式从源码目录取。"""
    if getattr(sys, 'frozen', False):
        return os.path.join(getattr(sys, '_MEIPASS', _app_dir()), 'app_icon.ico')
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'app_icon.ico')


def _mirror_urls(primary):
    """清单/下载地址 -> 多镜像列表（直链 + 代理镜像）。不使用 jsDelivr。"""
    urls = [primary]
    for p in _EXE_MIRROR_PREFIX:
        urls.append(p.rstrip('/') + '/' + primary)
    if 'raw.githubusercontent.com' in primary:
        alt = primary.replace('raw.githubusercontent.com', 'github.com', 1)
        alt = re.sub(r'github\.com/([^/]+)/([^/]+)/main/',
                     r'github.com/\1/\2/raw/main/', alt, count=1)
        if alt != primary:
            urls.append(alt)
    return urls


def parse_version(v):
    """'1.10' -> (1,10,0)；'v1.9.2' -> (1,9,2)。"""
    v = (v or '').strip().lstrip('vV')
    parts = []
    for seg in v.split('.'):
        m = re.findall(r'\d+', seg)
        parts.append(int(m[0]) if m else 0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:3])


def cmp_version(a, b):
    pa, pb = parse_version(a), parse_version(b)
    return (pa > pb) - (pa < pb)


def update_config_path():
    return os.path.join(_app_dir(), 'update_config.json')


def load_update_config():
    cfg = {'source': None, 'skip_version': None}
    p = update_config_path()
    if os.path.isfile(p):
        try:
            with open(p, 'r', encoding='utf-8') as f:
                cfg.update(json.load(f))
        except Exception:
            pass
    return cfg


def save_update_config(cfg):
    try:
        with open(update_config_path(), 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _http_get_bytes(url, start=None, end=None, timeout=30):
    headers = {'User-Agent': 'Mozilla/5.0'}
    if start is not None:
        headers['Range'] = 'bytes=%d-%s' % (start, end if end is not None else '')
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def _fetch_manifest(urls):
    """并发请求多个清单地址，返回 (best_dict, hit_source) 或 (None, None)。"""
    results = {}

    def worker(u):
        try:
            raw = _http_get_bytes(u, timeout=10)
            txt = raw.decode('utf-8', 'replace')
            if not txt.lstrip().startswith('{'):
                return
            d = json.loads(txt)
            if isinstance(d, dict) and d.get('version'):
                # 校验 download_url 存在
                if d.get('download_url'):
                    results[u] = d
        except Exception:
            pass
    ths = [threading.Thread(target=worker, args=(u,), daemon=True) for u in urls]
    for t in ths:
        t.start()
    for t in ths:
        t.join(timeout=12)
    if not results:
        return None, None
    best = max(results, key=lambda u: parse_version(results[u].get('version', '0')))
    return results[best], best


def _probe_size(mirrors):
    for base in mirrors:
        try:
            req = urllib.request.Request(base, headers={'User-Agent': 'Mozilla/5.0',
                                                         'Range': 'bytes=0-0'})
            with urllib.request.urlopen(req, timeout=10) as r:
                cr = r.headers.get('Content-Range', '')
                m = re.search(r'/\s*(\d+)', cr)
                if m:
                    return int(m.group(1))
                cl = r.headers.get('Content-Length')
                if cl:
                    return int(cl)
        except Exception:
            continue
    return 0


def verify_exe(path, sha256=None, size=None):
    """三重校验：存在 / MZ 头 / 大小 / SHA256。"""
    if not os.path.isfile(path):
        return False, '文件不存在'
    sz = os.path.getsize(path)
    if size and sz != size:
        return False, '文件大小不符（%d != %d）' % (sz, size)
    with open(path, 'rb') as f:
        head = f.read(2)
    if head != b'MZ':
        return False, '文件头不是 MZ（不是有效 exe）'
    if sha256:
        h = hashlib.sha256()
        with open(path, 'rb') as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b''):
                h.update(chunk)
        if h.hexdigest().lower() != sha256.lower():
            return False, 'SHA256 校验失败'
    return True, 'ok'


def download_update(download_url, dest, sha256=None, size=None,
                   on_progress=None, on_stage=None):
    """分段并发下载 + 断点续传 + 镜像轮换 + 校验。返回 (ok, msg)。"""
    mirrors = _mirror_urls(download_url)
    total = size or _probe_size(mirrors)
    if not total:
        return False, '无法获取文件大小（网络或地址错误）'
    if on_stage:
        on_stage('总大小 %.1f MB，开始下载…' % (total / 1024.0 / 1024.0))
    tmpdir = dest + '.parts'
    try:
        os.makedirs(tmpdir, exist_ok=True)
        seg = max(1, min(5, total // (1024 * 1024)))   # 每片 >=1MB，最多 5 片
        seg_size = total // seg
        segs = []
        for i in range(seg):
            s = i * seg_size
            e = total - 1 if i == seg - 1 else s + seg_size - 1
            segs.append((i, s, e))
        done = [0]
        lock = threading.Lock()
        errs = []
        part_paths = [os.path.join(tmpdir, 'seg_%d' % i) for i in range(seg)]

        def try_seg(idx, s, e):
            part = part_paths[idx]
            existing = os.path.getsize(part) if os.path.isfile(part) else 0
            cur_s = s + existing
            if cur_s > e:
                return
            for mi, base in enumerate(mirrors):
                for attempt in range(3):
                    try:
                        data = _http_get_bytes(base, cur_s, e)
                        if not data:
                            continue
                        with open(part, 'ab') as f:
                            f.write(data)
                        with lock:
                            done[0] += len(data)
                        return
                    except Exception as ex:
                        if mi == len(mirrors) - 1 and attempt == 2:
                            errs.append(str(ex)[:120])

        ths = [threading.Thread(target=try_seg, args=(i, s, e), daemon=True)
               for (i, s, e) in segs]
        for t in ths:
            t.start()
        while any(t.is_alive() for t in ths):
            if on_progress:
                on_progress(done[0], total)
            time.sleep(0.2)
        for t in ths:
            t.join(timeout=1)
        if on_progress:
            on_progress(done[0], total)
        # 检查每片都完整
        for i in range(seg):
            p = part_paths[i]
            if not os.path.isfile(p) or os.path.getsize(p) == 0:
                return False, '分片 %d 下载失败：%s' % (i, errs[-1] if errs else '未知')
        with open(dest, 'wb') as out:
            for i in range(seg):
                with open(part_paths[i], 'rb') as pf:
                    shutil.copyfileobj(pf, out)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
    ok, msg = verify_exe(dest, sha256, size)
    if not ok:
        return False, msg
    return True, 'ok'


def install_update(new_exe):
    """把下载好的 new_exe 就地替换为当前 exe。失败必还原，绝不退出。"""
    current = os.path.abspath(sys.executable)
    bak = current + '.bak'
    try:
        if os.path.isfile(bak):
            os.remove(bak)
    except Exception:
        pass
    # 额外放一份到隔壁，方便手动双击
    try:
        base = os.path.splitext(os.path.basename(current))[0]
        neighbor = os.path.join(os.path.dirname(current), base + '_新版.exe')
        shutil.copy2(new_exe, neighbor)
    except Exception:
        pass
    try:
        os.rename(current, bak)           # Windows 允许重命名运行中的 exe
        shutil.copy2(new_exe, current)
        ok, msg = verify_exe(current)
        if not ok:
            try:
                os.remove(current)
            except Exception:
                pass
            os.rename(bak, current)
            return False, '安装校验失败，已还原原程序：' + msg
        return True, 'ok'
    except Exception as ex:
        try:
            if os.path.isfile(bak):
                os.rename(bak, current)
        except Exception:
            pass
        return False, '安装失败，已尽量还原：' + str(ex)[:120]


def restart_app(root=None):
    """PyInstaller 单文件重启：剥离临时环境变量，拉起新实例后退出旧实例。"""
    env = dict(os.environ)
    for k in ('_MEIPASS2', '_PYI_APPLICATION_HOME_DIR', '_PYI_ARCHIVE_FILE'):
        env.pop(k, None)
    env['PYINSTALLER_RESET_ENVIRONMENT'] = '1'
    try:
        subprocess.Popen([sys.executable], env=env)
    except Exception:
        pass
    try:
        if root is not None:
            root.quit()
    except Exception:
        pass
    threading.Timer(3, lambda: os._exit(0)).start()
    try:
        if root is not None:
            root.destroy()
    except Exception:
        pass
    os._exit(0)


def cleanup_mei():
    """启动时清理历史残留的 _MEI 临时目录（单文件打包会堆到 GB 级）。

    注意：必须跳过当前进程自己的解压目录 sys._MEIPASS——
    否则会把本进程刚解压出来的 tcl/tk 脚本数据删掉，
    下一行 tk.Tk() 就会报「Can't find a usable init.tcl」。
    （被占用的 dll 删不掉、纯数据 .tcl/.enc 却删得掉，正好造成这个现象。）
    """
    try:
        tmp = tempfile.gettempdir()
        me = getattr(sys, '_MEIPASS', None)          # 打包运行时 = 当前进程解压目录；源码运行时无此属性
        me_abs = os.path.abspath(me).lower() if me else None
        for name in os.listdir(tmp):
            if name.startswith('_MEI') and os.path.isdir(os.path.join(tmp, name)):
                p = os.path.abspath(os.path.join(tmp, name))
                if me_abs and p.lower() == me_abs:
                    continue   # 自己的解压目录，绝不能删
                shutil.rmtree(p, ignore_errors=True)
    except Exception:
        pass


# ---- Excel 版专用布局参数 ----
REMARK_COL = 12       # 备注列 = M（0基: A=0, M=12）
IMG_FIRST = 13        # 第一张图片列 = N
IMG_LAST = 21         # 最后一张图片列 = V
SKIP_FIRST_ROW = True
DEFAULT_UPSCALE = 3.0
MAX_SCAN_COL = 26


# ============================================================
# 通用小工具
# ============================================================
def natural_key(s):
    """文件名自然排序：pic1, pic2, pic10 而不是 pic1, pic10, pic2。"""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', s)]


def col_to_idx(col):
    """列字母 -> 0 基索引：'A'=0, 'N'=13, 'M'=12"""
    idx = 0
    for ch in str(col).upper():
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1


def parse_col(s):
    """列字母/数字 -> 0 基索引；空返回 None。'M'->12, '1'->0, 'AA'->26"""
    s = (s or '').strip().upper()
    if not s:
        return None
    if s.isdigit():
        return int(s) - 1
    try:
        return col_to_idx(s)
    except Exception:
        return None


def idx_to_col(idx):
    """0 基索引 -> 列字母：0='A', 13='N'"""
    s = ''
    idx += 1
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(ord('A') + r) + s
    return s


def upscale_image(blob, factor):
    """用 LANCZOS 放大，返回新图片字节（PNG）。factor<=1 原样返回。"""
    if factor <= 1:
        return blob
    try:
        im = Image.open(io.BytesIO(blob)).convert('RGB')
        w, h = im.size
        im = im.resize((int(w * factor), int(h * factor)), Image.LANCZOS)
        buf = io.BytesIO()
        im.save(buf, format='PNG')
        return buf.getvalue()
    except Exception:
        return blob


def list_images(folder):
    """返回文件夹内所有图片的完整路径，按文件名自然排序。"""
    files = [os.path.join(folder, f) for f in os.listdir(folder)
             if f.lower().endswith(IMG_EXTS)]
    files.sort(key=lambda p: natural_key(os.path.basename(p)))
    return files


# ============================================================
# Excel 解析（内嵌 DISPIMG / 兜底浮动图）
# ============================================================
def parse_cellimages(xlsx_path):
    """解析 xl/cellimages.xml，返回 (id2blob, has_cellimages)。"""
    try:
        z = zipfile.ZipFile(xlsx_path)
    except Exception:
        return {}, False
    try:
        cxml = z.read('xl/cellimages.xml').decode('utf-8')
    except KeyError:
        return {}, False

    id2rid = {}
    for b in re.findall(r'<etc:cellImage>.*?</etc:cellImage>', cxml, re.S):
        name = re.search(r'<xdr:cNvPr[^>]*\bname="([^"]*)"', b)
        rid = re.search(r'r:embed="([^"]*)"', b)
        if name and rid:
            id2rid[name.group(1)] = rid.group(1)

    rid2media = {}
    try:
        rels = z.read('xl/_rels/cellimages.xml.rels').decode('utf-8')
        for rid, tgt in re.findall(r'Id="([^"]+)"[^>]*?Target="([^"]*)"', rels):
            rid2media[rid] = tgt
    except KeyError:
        pass

    id2blob = {}
    for name, rid in id2rid.items():
        tgt = rid2media.get(rid)
        if not tgt:
            continue
        path = tgt if tgt.startswith('xl/') else ('xl/' + tgt.lstrip('/'))
        try:
            id2blob[name] = z.read(path)
        except KeyError:
            pass
    return id2blob, True


def read_shared_strings(z):
    try:
        sxml = z.read('xl/sharedStrings.xml').decode('utf-8')
    except KeyError:
        return []
    out = []
    for si in re.findall(r'<si>(.*?)</si>', sxml, re.S):
        parts = re.findall(r'<t[^>]*>(.*?)</t>', si, re.S)
        out.append(html.unescape(''.join(parts)))
    return out


def read_cell_value(sheet_xml, col_idx, row, shared_strings):
    """读取某工作表 xml 中 (row, col_idx) 单元格的显示值。"""
    ref = f'{idx_to_col(col_idx)}{row}'
    m = re.search(r'<c r="%s"([^>]*)>(.*?)</c>' % re.escape(ref), sheet_xml, re.S)
    if not m:
        return ''
    attrs, body = m.group(1), m.group(2)
    t = re.search(r't="([^"]+)"', attrs)
    t = t.group(1) if t else 'n'
    if t == 's':
        v = re.search(r'<v>(.*?)</v>', body)
        if v and v.group(1).isdigit():
            i = int(v.group(1))
            return html.unescape(shared_strings[i]) if 0 <= i < len(shared_strings) else ''
        return ''
    if t in ('str',) or 'inlineStr' in attrs or '<is>' in body:
        return html.unescape(''.join(re.findall(r'<t[^>]*>(.*?)</t>', body, re.S)))
    v = re.search(r'<v>(.*?)</v>', body)
    return v.group(1) if v else ''


def get_sheet_order(z):
    """返回有序的 [(sheet_name, worksheet_path), ...]；失败返回 None。"""
    try:
        wb = z.read('xl/workbook.xml').decode('utf-8')
        rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    except Exception:
        return None
    sheets = re.findall(r'<sheet\b[^>]*\bname="([^"]+)"[^>]*\br:id="([^"]+)"', wb)
    rid2tgt = dict(re.findall(r'Id="([^"]+)"[^>]*?Target="([^"]*)"', rels))
    ordered = []
    for name, rid in sheets:
        tgt = rid2tgt.get(rid, '')
        if not tgt:
            continue
        path = tgt if tgt.startswith('xl/') else ('xl/' + tgt.lstrip('/'))
        ordered.append((name, path))
    return ordered or None


def scan_dispiimg(sheet_xml):
    """逐格扫描 DISPIMG，返回 [(col_idx, row, id_name), ...]。"""
    out = []
    cell_pat = re.compile(r'<c r="([A-Z]+\d+)"([^>]*)>(.*?)</c>', re.S)
    id_pat = re.compile(r'DISPIMG\((?:&quot;|")(ID_[A-Za-z0-9]+)(?:&quot;|")')
    for m in cell_pat.finditer(sheet_xml):
        ref, _attrs, body = m.group(1), m.group(2), m.group(3)
        if 'DISPIMG' not in body:
            continue
        idm = id_pat.search(body)
        if not idm:
            continue
        cm = re.match(r'([A-Z]+)(\d+)', ref)
        if not cm:
            continue
        col = col_to_idx(cm.group(1))
        row = int(cm.group(2))
        out.append((col, row, idm.group(1)))
    return out


def _fallback_drawing(xlsx_path, remark_col, img_first, img_last,
                      shared_strings, skip_first_row, sidx_xml):
    """兜底：openpyxl 读浮动图片，按锚点行分组。备注由调用方统一读 XML。"""
    try:
        import openpyxl
    except Exception:
        return {}
    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    except Exception:
        return {}
    rows_map = {}
    for sidx, ws in enumerate(wb.worksheets):
        for image in getattr(ws, '_images', []):
            anchor = getattr(image, 'anchor', None)
            if anchor is None:
                continue
            frm = getattr(anchor, '_from', None) or anchor
            r = getattr(frm, 'row', None)
            c = getattr(frm, 'col', None)
            if r is None or c is None:
                continue
            r, c = int(r) + 1, int(c) + 1
            if c < img_first or c > img_last:
                continue
            if skip_first_row and r == 1:
                continue
            blob = None
            ref = getattr(image, 'ref', None)
            if hasattr(ref, 'getvalue'):
                try:
                    blob = ref.getvalue()
                except Exception:
                    blob = None
            if not blob:
                blob = getattr(image, 'blob', None)
                if hasattr(blob, 'getvalue'):
                    blob = blob.getvalue()
            if not blob:
                continue
            key = (sidx, r)
            d = rows_map.setdefault(key, {'cols': []})
            d['cols'].append((c, blob))
    return rows_map


def load_image_rows(xlsx_path, remark_col=REMARK_COL, img_first=IMG_FIRST,
                    img_last=IMG_LAST, skip_first_row=SKIP_FIRST_ROW,
                    mode='auto'):
    """返回 [(remark_text, [img_blob, ...]), ...]，按行序、行内按列序。
       只读取 备注列 + 图片列；标题行忽略。备注始终从工作表 XML 读，因此即便走兜底路径也不会丢。
       mode: 'auto' 自动（先嵌入后浮动兜底）/ 'embedded' 仅 DISPIMG 嵌入 / 'floating' 仅浮动粘贴。
       无图片返回 []。"""
    z = zipfile.ZipFile(xlsx_path)
    ss = read_shared_strings(z)
    id2blob, has_ci = parse_cellimages(xlsx_path)

    order = get_sheet_order(z)
    if order:
        sheets = order
    else:
        sheets = [(None, n) for n in sorted(
            n for n in z.namelist() if re.match(r'xl/worksheets/sheet\d+\.xml$', n))]

    if mode == 'floating':
        use_embedded, use_floating = False, True
    elif mode == 'embedded':
        use_embedded, use_floating = True, False
    else:  # auto
        use_embedded, use_floating = bool(has_ci and id2blob), True

    sidx_xml = {}
    rows_map = {}
    for sidx, (sname, spath) in enumerate(sheets):
        try:
            xml = z.read(spath).decode('utf-8')
        except Exception:
            continue
        sidx_xml[sidx] = xml
        cells = scan_dispiimg(xml) if use_embedded else []
        for col, row, idn in cells:
            if col < img_first or col > img_last:
                continue
            blob = id2blob.get(idn)
            if not blob:
                continue
            if skip_first_row and row == 1:
                continue
            key = (sidx, row)
            d = rows_map.setdefault(key, {'cols': []})
            d['cols'].append((col, blob))

    if not rows_map and use_floating:
        rows_map = _fallback_drawing(xlsx_path, remark_col, img_first, img_last,
                                     ss, skip_first_row, sidx_xml)
        if not rows_map:
            return []

    groups = []
    for (sidx, row), d in rows_map.items():
        d['cols'].sort(key=lambda x: x[0])
        blobs = [b for _, b in d['cols']]
        remark = ''
        xml = sidx_xml.get(sidx, '')
        if 0 <= remark_col < MAX_SCAN_COL and xml:
            remark = (read_cell_value(xml, remark_col, row, ss) or '').strip()
        groups.append((sidx, row, remark, blobs))

    groups.sort(key=lambda g: (g[0], g[1]))
    return [(g[2], g[3]) for g in groups]


# ============================================================
# 通用 Word 构建（item: bytes 二进制 或 str 文件路径）
# ============================================================
def build_docx(rows, out_path, img_w_cm=IMG_W_CM, img_h_cm=IMG_H_CM,
               cols=DEFAULT_COLS, upscale=1.0, show_remark=False,
               on_progress=None):
    """把 rows=(remark, [items]) 拼进 Word。
       item 可以是 图片字节(bytes) 或 文件路径(str)；后者不放大、按原样放入。
       固定尺寸 宽7cm×高11cm，A4 纸 + 1.5cm 页边距（2x2=4张/页）。"""
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(1.5)
    sec.bottom_margin = Cm(1.5)
    sec.left_margin = Cm(1.5)
    sec.right_margin = Cm(1.5)
    style = doc.styles['Normal']
    style.font.name = '宋体'
    style.font.size = Pt(11)
    try:
        from docx.oxml.ns import qn
        rpr = style.element.get_or_add_rPr()
        rfonts = rpr.find(qn('w:rFonts'))
        if rfonts is None:
            rfonts = rpr.makeelement(qn('w:rFonts'), {})
            rpr.append(rfonts)
        rfonts.set(qn('w:eastAsia'), '宋体')
    except Exception:
        pass

    total = sum(len(b) for _, b in rows)
    done = 0

    # 取消备注：全部图片按 (行->列) 顺序拼成一条连续序列
    if not show_remark:
        flat = []
        for _, blobs in rows:
            flat.extend(blobs)
        rows = [('', flat)]

    for remark, blobs in rows:
        if show_remark and remark:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(remark)
            run.bold = True
            run.font.size = Pt(12)

        if not blobs:
            continue
        per = min(cols, len(blobs)) or 1
        for i in range(0, len(blobs), per):
            chunk = blobs[i:i + per]
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            for item in chunk:
                run = p.add_run()
                name = ''
                if isinstance(item, (bytes, bytearray)):
                    up = upscale_image(item, upscale)
                    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
                    tmp.write(up)
                    tmp.close()
                    run.add_picture(tmp.name, width=Cm(img_w_cm), height=Cm(img_h_cm))
                    os.unlink(tmp.name)
                else:
                    run.add_picture(item, width=Cm(img_w_cm), height=Cm(img_h_cm))
                    name = os.path.basename(item)
                done += 1
                if on_progress:
                    try:
                        on_progress(done, total, name)
                    except Exception:
                        pass
    doc.save(out_path)
    return total


# ============================================================
# 工具基类 + 两个具体工具
# ============================================================
class Tool:
    """功能模块基类。后续加功能只需继承此类并实现 build/prepare/run。"""
    name = ''
    desc = ''
    intro = ''   # 功能介绍（小白引导）；新增功能请同步填写，框架会自动渲染

    def build(self, parent, app):
        """在 parent 中创建配置控件。"""
        raise NotImplementedError

    def prepare(self, app):
        """主线程调用：把 UI 值读到 self.params（供子线程安全读取）。"""
        self.params = {}

    def run(self, app):
        """子线程中执行；通过 app.log / set_progress / finish 汇报。"""
        raise NotImplementedError


class FolderTool(Tool):
    name = '文件夹图片 → Word'
    desc = '选一个图片文件夹，按文件名顺序导出到 Word（A4 / 2×2 / 7×11cm，无备注）'
    intro = ('功能介绍：把一个文件夹里的所有图片，按顺序拼进一个 Word 文档，方便打印或存档。\n'
             '· 每张图片固定 7×11cm，A4 纸每页放 4 张（2×2 排列）；一页放不下会自动翻页。\n'
             '· 支持 png / jpg / bmp / gif / webp / tif 等常见格式，按文件名顺序排（pic1、pic2…）。\n'
             '· 默认不显示备注、不做放大，原图直出。\n'
             '用法：选好图片文件夹 → 设置每行列数（默认 2）→ 点「开始导出」。\n'
             '成品「图片导出.docx」会生成在原文件夹里。')

    def build(self, parent, app):
        self.app = app
        f = ttk.LabelFrame(parent, text='文件夹图片导出', padding=10)
        f.pack(fill='both', expand=True, padx=8, pady=8)

        ttk.Label(f, text=self.desc, wraplength=520,
                  foreground='#555').pack(anchor='w', pady=(0, 8))

        row = ttk.Frame(f)
        row.pack(fill='x', pady=4)
        ttk.Label(row, text='图片文件夹：', width=12).pack(side='left')
        self.folder_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.folder_var).pack(side='left', fill='x', expand=True, padx=(0, 6))
        ttk.Button(row, text='浏览…', command=self._pick).pack(side='left')

        row2 = ttk.Frame(f)
        row2.pack(fill='x', pady=4)
        ttk.Label(row2, text='每行列数：', width=12).pack(side='left')
        self.cols_var = tk.StringVar(value=str(DEFAULT_COLS))
        ttk.Spinbox(row2, from_=1, to=4, width=8, textvariable=self.cols_var).pack(side='left')
        ttk.Label(row2, text='   （A4 纸一行最多 2 张，2 列即 4 张/页）',
                  foreground='#888').pack(side='left')

        ttk.Label(f, text='支持：png / jpg / jpeg / bmp / gif / webp / tif / tiff',
                  foreground='#888').pack(anchor='w', pady=(6, 0))

        ttk.Button(f, text='▶ 开始导出', command=self._on_run,
                   style='Accent.TButton').pack(anchor='e', pady=(12, 0))

    def _pick(self):
        p = filedialog.askdirectory(title='选择包含图片的文件夹')
        if p:
            self.folder_var.set(p)

    def _on_run(self):
        self.app.execute(self)

    def prepare(self, app):
        self.params = {
            'folder': self.folder_var.get().strip(),
            'cols': max(1, int(self.cols_var.get() or DEFAULT_COLS)),
        }

    def run(self, app):
        folder = self.params['folder']
        if not folder or not os.path.isdir(folder):
            app.log('⚠ 请先选择有效的图片文件夹。')
            return
        images = list_images(folder)
        if not images:
            app.log('⚠ 该文件夹中没有找到图片。')
            return
        app.log(f'找到 {len(images)} 张图片，开始导出（按文件名顺序）…')
        out = os.path.join(folder, '图片导出.docx')
        total = build_docx([('', images)], out, cols=self.params['cols'],
                            upscale=1.0, show_remark=False,
                            on_progress=app.make_progress_cb(len(images)))
        app.finish(out)
        app.log(f'共处理 {total} 张图片。')


class ExcelTool(Tool):
    name = 'Excel 图片 → Word'
    desc = '在右侧配置表格布局（备注列 / 图片起止列 / 图片放置方式），软件据此实时解析并导出 Word'
    intro = ('功能介绍：把 Excel 里「一行 = 一组图」的图片导出成 Word，常用于把填好图的报表整理成可打印文档。\n'
             '· 默认读 M 列作备注、N~V 列作图片；若你的表格列不同，可在界面改「备注列 / 图片从 / 图片到」。\n'
             '· 图片放置支持三种：自动识别 / 嵌入单元格(DISPIMG 公式) / 浮动粘贴（直接粘的截图）。\n'
             '· 先点「解析预览」确认识别到几行、共几张图，没问题再点「开始导出」。\n'
             '· 默认把全部图片按 行→列 顺序拼合（不显示备注）；需要保留备注文字再勾选「显示备注」。\n'
             '用法：选 Excel → 按需调列配置 → 预览 → 导出，成品为「原文件名_导出.docx」。')

    def build(self, parent, app):
        self.app = app
        f = ttk.LabelFrame(parent, text='Excel 图片导出', padding=10)
        f.pack(fill='both', expand=True, padx=8, pady=8)

        ttk.Label(f, text=self.desc, wraplength=520,
                  foreground='#555').pack(anchor='w', pady=(0, 8))

        row = ttk.Frame(f)
        row.pack(fill='x', pady=4)
        ttk.Label(row, text='Excel 文件：', width=12).pack(side='left')
        self.file_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.file_var).pack(side='left', fill='x', expand=True, padx=(0, 6))
        ttk.Button(row, text='浏览…', command=self._pick).pack(side='left')

        # 列配置：备注列 / 图片起 / 图片止
        rowc = ttk.Frame(f)
        rowc.pack(fill='x', pady=4)
        ttk.Label(rowc, text='备注列：', width=12).pack(side='left')
        self.remark_col_var = tk.StringVar(value='M')
        ttk.Entry(rowc, textvariable=self.remark_col_var, width=6).pack(side='left', padx=(0, 12))
        ttk.Label(rowc, text='图片从：').pack(side='left')
        self.imgf_var = tk.StringVar(value='N')
        ttk.Entry(rowc, textvariable=self.imgf_var, width=6).pack(side='left', padx=(0, 4))
        ttk.Label(rowc, text='到：').pack(side='left')
        self.imgl_var = tk.StringVar(value='V')
        ttk.Entry(rowc, textvariable=self.imgl_var, width=6).pack(side='left')

        # 图片放置方式
        rowm = ttk.Frame(f)
        rowm.pack(fill='x', pady=4)
        ttk.Label(rowm, text='图片放置：', width=12).pack(side='left')
        self.mode_var = tk.StringVar(value='自动识别')
        ttk.Combobox(rowm, textvariable=self.mode_var, width=20,
                     values=['自动识别', '嵌入单元格', '浮动粘贴'],
                     state='readonly').pack(side='left')
        ttk.Label(rowm, text='（嵌入=DISPIMG公式；浮动=直接粘贴截图）',
                  foreground='#888').pack(side='left', padx=(6, 0))

        row2 = ttk.Frame(f)
        row2.pack(fill='x', pady=4)
        ttk.Label(row2, text='每行列数：', width=12).pack(side='left')
        self.cols_var = tk.StringVar(value=str(DEFAULT_COLS))
        ttk.Spinbox(row2, from_=1, to=4, width=8, textvariable=self.cols_var).pack(side='left')
        ttk.Label(row2, text='   放大倍数：', width=10).pack(side='left')
        self.up_var = tk.StringVar(value=str(DEFAULT_UPSCALE))
        ttk.Spinbox(row2, from_=1, to=6, increment=0.5, width=8, textvariable=self.up_var).pack(side='left')

        self.remark_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text='显示备注（备注列文字，不拼合）',
                        variable=self.remark_var).pack(anchor='w', pady=(6, 0))
        ttk.Label(f, text='默认：不显示备注，全部图片按行→列顺序拼合',
                  foreground='#888').pack(anchor='w', pady=(2, 0))

        # 操作按钮
        rowb = ttk.Frame(f)
        rowb.pack(fill='x', pady=(12, 0))
        ttk.Button(rowb, text='▶ 解析预览', command=self._on_preview,
                   style='Accent.TButton').pack(side='left')
        ttk.Button(rowb, text='▶ 开始导出', command=self._on_run,
                   style='Accent.TButton').pack(side='right')

    def _pick(self):
        p = filedialog.askopenfilename(title='选择 Excel 文件',
                                       filetypes=[('Excel 文件', '*.xlsx'), ('所有文件', '*.*')])
        if p:
            self.file_var.set(p)

    def _resolve_params(self):
        """从界面读取并解析参数。"""
        rc = parse_col(self.remark_col_var.get()) if self.remark_col_var.get().strip() else REMARK_COL
        ifc = parse_col(self.imgf_var.get()) if self.imgf_var.get().strip() else IMG_FIRST
        ilc = parse_col(self.imgl_var.get()) if self.imgl_var.get().strip() else IMG_LAST
        mode = {'自动识别': 'auto', '嵌入单元格': 'embedded', '浮动粘贴': 'floating'}.get(
            self.mode_var.get(), 'auto')
        return {
            'path': self.file_var.get().strip(),
            'remark_col': rc, 'img_first': ifc, 'img_last': ilc,
            'mode': mode,
            'cols': max(1, int(self.cols_var.get() or DEFAULT_COLS)),
            'upscale': max(1.0, float(self.up_var.get() or DEFAULT_UPSCALE)),
            'remark': bool(self.remark_var.get()),
        }

    def _on_run(self):
        self.params = self._resolve_params()
        if not self.params['path'] or not os.path.isfile(self.params['path']):
            self.app.log('⚠ 请先选择有效的 Excel 文件。')
            return
        self.app.execute(self)

    def _on_preview(self):
        """解析预览：按当前配置统计行/图片数，帮助实时确认列配置是否正确。"""
        p = self._resolve_params()
        if not p['path'] or not os.path.isfile(p['path']):
            self.app.log('⚠ 请先选择有效的 Excel 文件。')
            return
        try:
            rows = load_image_rows(p['path'], remark_col=p['remark_col'],
                                   img_first=p['img_first'], img_last=p['img_last'],
                                   mode=p['mode'])
        except Exception as e:
            self.app.log('解析失败：' + str(e))
            return
        if not rows:
            self.app.log('⚠ 未识别到图片。请检查：备注列/图片起止列是否正确，图片放置方式是否匹配。')
            return
        n_img = sum(len(b) for _, b in rows)
        self.app.log(f'【解析预览】备注列={idx_to_col(p["remark_col"])}  图片列='
                     f'{idx_to_col(p["img_first"])}~{idx_to_col(p["img_last"])}  方式={p["mode"]}')
        self.app.log(f'识别到 {len(rows)} 行分组、共 {n_img} 张图片：')
        for i, (remark, blobs) in enumerate(rows):
            self.app.log(f'   第{i + 1}组  备注={remark!r}  图片数={len(blobs)}')

    def prepare(self, app):
        # _on_run 已解析好 self.params，此处无需再读 UI
        pass

    def run(self, app):
        p = self.params
        try:
            rows = load_image_rows(p['path'], remark_col=p['remark_col'],
                                   img_first=p['img_first'], img_last=p['img_last'],
                                   mode=p['mode'])
        except Exception as e:
            app.log('读取 Excel 失败：' + str(e))
            return
        if not rows:
            app.log('⚠ 该 Excel 中未识别到图片（请检查列配置与图片放置方式）。')
            return
        n_img = sum(len(b) for _, b in rows)
        app.log(f'识别到 {len(rows)} 行分组、共 {n_img} 张图片，开始导出…')
        out = os.path.splitext(p['path'])[0] + '_导出.docx'
        build_docx(rows, out, cols=p['cols'], upscale=p['upscale'],
                   show_remark=p['remark'],
                   on_progress=app.make_progress_cb(n_img))
        app.finish(out)
        app.log(f'共处理 {n_img} 张图片。')


# ============================================================
# 批量打印 PDF（内置：调用系统自带 PDF 打印能力，无需第三方 exe）
# ============================================================
def _list_printers():
    """返回 [(name, is_default), ...]；失败返回 []。依赖 Windows 自带 PowerShell。"""
    try:
        out = subprocess.run(
            ['powershell', '-NoProfile', '-Command',
             'Get-Printer | ForEach-Object { "$($_.Name)|$($_.IsDefault)" }'],
            capture_output=True, text=True, timeout=25)
        printers = []
        for line in out.stdout.splitlines():
            line = line.strip()
            if not line:
                continue
            parts = line.split('|')
            name = parts[0].strip()
            is_def = len(parts) > 1 and parts[1].strip().lower() == 'true'
            if name:
                printers.append((name, is_def))
        return printers
    except Exception:
        return []


def _find_acro():
    """返回 Adobe Reader/Acrobat 的 AcroRd32.exe 路径或 None。"""
    candidates = [
        r'C:/Program Files (x86)/Adobe/Acrobat Reader DC/Reader/AcroRd32.exe',
        r'C:/Program Files/Adobe/Acrobat Reader DC/Reader/AcroRd32.exe',
        r'C:/Program Files (x86)/Adobe/Reader/Reader/AcroRd32.exe',
        r'C:/Program Files (x86)/Adobe/Acrobat/Acrobat/AcroRd32.exe',
        r'C:/Program Files/Adobe/Acrobat/Acrobat/AcroRd32.exe',
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def print_pdf(path, printer=None):
    """打印单个 PDF，返回 (ok, msg)。
       优先用 Adobe Reader 的 /t 静默打印（可指定打印机）；
       否则用系统默认 PDF 应用（通常是 Edge）的 print 动词（仅默认打印机）。
       均依赖 Windows 自带能力，不调用第三方 exe。"""
    if not os.path.isfile(path):
        return False, '文件不存在'
    acro = _find_acro()
    if acro:
        cmd = [acro, '/t', path]
        if printer:
            cmd.append(printer)
        try:
            proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            try:
                proc.wait(timeout=60)
            except subprocess.TimeoutExpired:
                proc.kill()  # 打印任务通常已提交，超时仅清理进程
            return True, '已通过 Adobe Reader 发送' + (f'（{printer}）' if printer else '（默认打印机）')
        except Exception as e:
            return False, 'Adobe 打印失败：' + str(e)
    # 兜底：系统默认应用 print 动词（Edge 等），仅默认打印机
    try:
        ps = f'Start-Process -FilePath "{path}" -Verb Print'
        r = subprocess.run(['powershell', '-NoProfile', '-Command', ps],
                           capture_output=True, text=True, timeout=60)
        if r.returncode != 0:
            return False, '系统打印失败：' + (r.stderr.strip()[:200] or '返回码非 0')
        note = '（默认打印机）' if not printer else '（系统默认应用仅支持默认打印机）'
        return True, '已通过系统默认应用发送打印' + note
    except Exception as e:
        return False, '打印失败：' + str(e)


class PrintTool(Tool):
    name = '批量打印（PDF）'
    desc = '选择 PDF 文件或文件夹，逐个发送到打印机（用系统自带 PDF 打印能力，无需额外安装）'
    intro = ('功能介绍：把选中的 PDF 文件逐个发送到打印机，无需安装「打印助手」等第三方软件。\n'
             '· 支持多选 PDF，或选一整个文件夹批量打印。\n'
             '· 自动列出本机打印机，可指定某台；不选则用系统默认打印机。\n'
             '· 优先调用 Adobe Reader 静默打印；若没装就在系统默认应用（如 Edge）里打印（仅限默认打印机）。\n'
             '用法：选 PDF / 文件夹 → 选打印机 → 点「开始打印」。\n'
             '注意：真实打印需要本机已连好打印机；本功能不生成文件成品，只登记历史记录。')

    def build(self, parent, app):
        self.app = app
        self.selected = []
        f = ttk.LabelFrame(parent, text='批量打印 PDF', padding=10)
        f.pack(fill='both', expand=True, padx=8, pady=8)

        ttk.Label(f, text=self.desc, wraplength=520, foreground='#555').pack(anchor='w', pady=(0, 8))

        row = ttk.Frame(f)
        row.pack(fill='x', pady=4)
        ttk.Button(row, text='＋ 选择 PDF 文件…', command=self._pick_files).pack(side='left', padx=(0, 6))
        ttk.Button(row, text='＋ 选择文件夹…', command=self._pick_folder).pack(side='left')

        self.info_var = tk.StringVar(value='尚未选择文件')
        ttk.Label(f, textvariable=self.info_var, foreground='#2c3e50').pack(anchor='w', pady=(6, 0))
        ttk.Button(f, text='清空选择', command=self._clear).pack(anchor='w', pady=(2, 0))

        # 打印机选择（后台异步获取，避免切换功能时界面卡顿）
        self.printer_var = tk.StringVar(value='')
        self.printer_container = ttk.Frame(f)
        self.printer_container.pack(fill='x', pady=(12, 0))
        self._fill_printer_container()

        ttk.Label(f, text='说明：打印调用系统自带 PDF 打印能力（Edge / 默认应用 / Adobe Reader），'
                          '不依赖「打印助手V7」等第三方 exe。',
                  foreground='#888', wraplength=540).pack(anchor='w', pady=(14, 0))

        ttk.Button(f, text='▶ 开始打印', command=self._on_run,
                   style='Accent.TButton').pack(anchor='e', pady=(12, 0))

    def _pick_files(self):
        paths = filedialog.askopenfilenames(title='选择 PDF 文件',
                                            filetypes=[('PDF 文件', '*.pdf'), ('所有文件', '*.*')])
        if paths:
            self.selected = list(dict.fromkeys(self.selected + list(paths)))  # 去重保序
            self._refresh_info()

    def _pick_folder(self):
        d = filedialog.askdirectory(title='选择包含 PDF 的文件夹')
        if d:
            for fn in sorted(os.listdir(d), key=natural_key):
                if fn.lower().endswith('.pdf'):
                    self.selected.append(os.path.join(d, fn))
            self.selected = list(dict.fromkeys(self.selected))
            self._refresh_info()

    def _clear(self):
        self.selected = []
        self._refresh_info()

    def _refresh_info(self):
        n = len(self.selected)
        self.info_var.set('尚未选择文件' if n == 0 else f'已选择 {n} 个 PDF（点「开始打印」后逐个发送）')

    def _fill_printer_container(self):
        """填充打印机选择区：优先用缓存；未获取则后台拉取，拿到后刷新。"""
        global _PRINTERS
        for w in self.printer_container.winfo_children():
            w.destroy()
        printers = _PRINTERS if _PRINTERS is not None else []
        if _PRINTERS is None:
            ttk.Label(self.printer_container, text='（正在检测本机打印机…）',
                      foreground='#a55').pack(anchor='w')
            # 首次进入：后台获取，避免阻塞界面
            threading.Thread(target=self._fetch_printers, daemon=True).start()
            return
        if printers:
            rowp = ttk.Frame(self.printer_container)
            rowp.pack(fill='x')
            ttk.Label(rowp, text='打印机：', width=10).pack(side='left')
            default_name = next((n for n, d in printers if d), '')
            if not self.printer_var.get():
                self.printer_var.set(default_name)
            ttk.Combobox(rowp, textvariable=self.printer_var, width=42,
                         values=[n for n, _ in printers], state='readonly').pack(side='left')
        else:
            ttk.Label(self.printer_container,
                      text='（未检测到打印机列表，将使用系统默认打印机）',
                      foreground='#a55').pack(anchor='w')

    def _fetch_printers(self):
        global _PRINTERS
        try:
            _PRINTERS = _list_printers()
        except Exception:
            _PRINTERS = []
        self.app.root.after(0, self._fill_printer_container)

    def _on_run(self):
        if not self.selected:
            self.app.log('⚠ 请先选择要打印的 PDF 文件或文件夹。')
            return
        self.app.execute(self)

    def prepare(self, app):
        self.params = {
            'paths': list(self.selected),
            'printer': self.printer_var.get().strip() or None,
        }

    def run(self, app):
        paths = self.params['paths']
        printer = self.params['printer']
        app.log('目标打印机：' + (printer or '系统默认'))
        total = len(paths)
        ok = fail = 0
        for i, p in enumerate(paths, 1):
            app.log(f'[{i}/{total}] 发送打印：{os.path.basename(p)}')
            good, msg = print_pdf(p, printer)
            if good:
                ok += 1
                if printer and 'Adobe' not in msg:
                    app.log('    ✓ ' + msg + '（系统默认应用仅打默认打印机）')
                else:
                    app.log('    ✓ ' + msg)
            else:
                fail += 1
                app.log('    ✗ ' + msg)
            app.set_progress(i, total)
        app.log(f'批量打印完成：成功 {ok} 个，失败 {fail} 个。')
        if fail:
            app.log('失败项请检查：文件是否损坏 / 是否有权限 / 打印机是否在线。')
        app.add_history('批量打印（PDF）', note=f'打印 {total} 个，成功 {ok} / 失败 {fail}')


class WebLinkTool(Tool):
    """通用「打开网站」功能：新增网址只需继承此类，定义 name/desc/intro/url 即可。"""
    frame_title = '在线工具'
    tip = '（点击链接或按钮，将在默认浏览器中打开网站）'

    def build(self, parent, app):
        self.app = app
        f = ttk.LabelFrame(parent, text=self.frame_title, padding=10)
        f.pack(fill='both', expand=True, padx=8, pady=8)

        ttk.Label(f, text=self.desc, wraplength=520,
                  foreground='#555').pack(anchor='w', pady=(0, 8))

        ttk.Label(f, text='链接：', foreground='#555').pack(anchor='w')
        link_lbl = tk.Label(f, text=self.url, fg='#1a73e8', cursor='hand2',
                            font='AppLink')
        link_lbl.pack(anchor='w', pady=(0, 10))
        link_lbl.bind('<Button-1>', lambda e: self._open())

        ttk.Button(f, text='🌐 打开网站', command=self._open,
                   style='Accent.TButton').pack(anchor='w', pady=(4, 0))
        ttk.Label(f, text=self.tip,
                  foreground='#888').pack(anchor='w', pady=(8, 0))

    def _open(self):
        try:
            webbrowser.open(self.url)
            self.app.log('已打开网站：' + self.url)
        except Exception as e:
            self.app.log('打开失败：' + str(e))


class LinkTool(WebLinkTool):
    name = 'PDF工具'
    desc = '在线 PDF 工具合集（PDF24）：合并、分割、压缩、转换、加水印等'
    intro = ('功能介绍：这里是 PDF24 在线工具箱的快捷入口，提供合并、分割、压缩、格式转换、加水印等常见 PDF 操作。\n'
             '· 点击链接文字或「打开网站」按钮，会在你的默认浏览器里打开 PDF24 中文站点。\n'
             '· 所有处理都在浏览器中完成，本软件不保存、不上传你的文件。\n'
             '用法：点「打开网站」即可使用，无需任何额外配置。')
    url = 'https://tools.pdf24.org/zh/'
    frame_title = 'PDF 在线工具'
    tip = '（点击链接或按钮，将在默认浏览器中打开 PDF24 在线工具）'


class TranslateTool(WebLinkTool):
    name = '沉浸式翻译'
    desc = '沉浸式翻译（Immersive Translate）：网页双语对照翻译、PDF / 电子书翻译等在线服务'
    intro = ('功能介绍：这里是「沉浸式翻译」官网的快捷入口，一个浏览器翻译工具。\n'
             '· 支持网页双语对照翻译，也能翻译 PDF、EPUB 电子书和视频字幕。\n'
             '· 点击链接文字或「打开网站」按钮，会在你的默认浏览器里打开官网。\n'
             '· 具体安装浏览器插件与账号设置，请在官网页面按提示操作；本软件不保存、不上传你的文件。\n'
             '用法：点「打开网站」即可使用，无需任何额外配置。')
    url = 'https://app.immersivetranslate.com/'
    frame_title = '沉浸式翻译'
    tip = '（点击链接或按钮，将在默认浏览器中打开沉浸式翻译官网）'


# ============================================================
# Word 批量替换（整合自 E:\docx批量替换\reminder_app.py）
#   保留原格式（按 run 替换）/ 支持正则 / 备份 / 页眉页脚表格 / 子文件夹 / 可停止
# ============================================================
class DocxReplaceTool(Tool):
    name = 'Word 批量替换'
    desc = '选择 Word 文档文件夹，批量查找并替换文字（保留原格式，可正则、可备份、可停止）'
    intro = ('功能介绍：批量在多个 Word 文档里查找并替换文字，并且保留原来的字体、颜色、排版等格式。\n'
             '· 可整文件夹处理，也能递归处理子文件夹；动手前可先「测试查找」看哪些文件命中。\n'
             '· 支持正则表达式（如 故障\\d+个），可只处理页眉 / 页脚 / 表格。\n'
             '· 强烈建议勾选「创建备份」，万一改错可找回原文件；旧版 .doc 会被自动跳过（请另存为 .docx）。\n'
             '· 处理过程中可随时点「停止」。\n'
             '用法：选文件夹 → 填「查找 / 替换为」→（可选）高级选项 → 点「开始替换」。')

    # ---------- 核心替换逻辑（按 run 操作，保留字体/颜色等格式）----------
    @staticmethod
    def _match(text, pattern, use_regex, case_sensitive):
        if not text:
            return False
        if use_regex:
            try:
                return bool(re.search(pattern, text, 0 if case_sensitive else re.IGNORECASE))
            except Exception:
                return False
        if case_sensitive:
            return pattern in text
        return pattern.lower() in text.lower()

    @staticmethod
    def _replace_in_runs(runs, old, new, use_regex, case_sensitive):
        for run in runs:
            t = run.text
            if not t:
                continue
            if use_regex:
                try:
                    if re.search(old, t, 0 if case_sensitive else re.IGNORECASE):
                        run.text = re.sub(old, new, t, flags=0 if case_sensitive else re.IGNORECASE)
                except Exception:
                    pass
            else:
                if (case_sensitive and old in t) or (not case_sensitive and old.lower() in t.lower()):
                    if case_sensitive:
                        run.text = t.replace(old, new)
                    else:
                        pat = old.lower()
                        tl = t.lower()
                        idx = tl.find(pat)
                        if idx != -1:
                            run.text = t[:idx] + new + t[idx + len(old):]

    @staticmethod
    def _process_file(path, old, new, backup_dir, use_regex, case_sensitive,
                      headers, footers, tables):
        """处理单个文件，返回是否修改。.doc 旧版会抛异常，由调用方跳过。"""
        if not os.path.exists(path) or os.path.getsize(path) == 0:
            return False
        if backup_dir:
            try:
                shutil.copy2(path, os.path.join(backup_dir, os.path.basename(path)))
            except Exception:
                pass
        modified = False
        doc = Document(path)
        for para in doc.paragraphs:
            if DocxReplaceTool._match(para.text, old, use_regex, case_sensitive):
                modified = True
                DocxReplaceTool._replace_in_runs(para.runs, old, new, use_regex, case_sensitive)
        if tables:
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            if DocxReplaceTool._match(para.text, old, use_regex, case_sensitive):
                                modified = True
                                DocxReplaceTool._replace_in_runs(para.runs, old, new, use_regex, case_sensitive)
        if headers:
            for sec in doc.sections:
                for para in sec.header.paragraphs:
                    if DocxReplaceTool._match(para.text, old, use_regex, case_sensitive):
                        modified = True
                        DocxReplaceTool._replace_in_runs(para.runs, old, new, use_regex, case_sensitive)
        if footers:
            for sec in doc.sections:
                for para in sec.footer.paragraphs:
                    if DocxReplaceTool._match(para.text, old, use_regex, case_sensitive):
                        modified = True
                        DocxReplaceTool._replace_in_runs(para.runs, old, new, use_regex, case_sensitive)
        if modified:
            doc.save(path)
        return modified

    # ---------- 界面 ----------
    def build(self, parent, app):
        self.app = app
        self.stop_req = False
        # 直接构建到 parent（_select 已提供可滚动画布 + 尺寸自适应，
        # 不再额外套一层画布，避免内容横向无法铺满的问题）。
        inner = parent

        def LF(title):
            f = ttk.LabelFrame(inner, text=title, padding=8)
            f.pack(fill='x', padx=8, pady=6)
            return f

        # 文件夹
        f = LF('Word 文档文件夹')
        row = ttk.Frame(f)
        row.pack(fill='x', pady=2)
        ttk.Label(row, text='文件夹：', width=10).pack(side='left')
        self.folder_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.folder_var).pack(side='left', fill='x', expand=True, padx=(0, 6))
        ttk.Button(row, text='浏览…', command=self._pick).pack(side='left')

        # 查找 / 替换
        f = LF('查找 / 替换')
        r1 = ttk.Frame(f)
        r1.pack(fill='x', pady=2)
        ttk.Label(r1, text='查找：', width=10).pack(side='left')
        self.find_var = tk.StringVar()
        ttk.Entry(r1, textvariable=self.find_var).pack(side='left', fill='x', expand=True)
        r2 = ttk.Frame(f)
        r2.pack(fill='x', pady=2)
        ttk.Label(r2, text='替换为：', width=10).pack(side='left')
        self.replace_var = tk.StringVar()
        ttk.Entry(r2, textvariable=self.replace_var).pack(side='left', fill='x', expand=True)

        # 选项
        f = LF('选项')
        self.backup_var = tk.BooleanVar(value=True)
        self.recursive_var = tk.BooleanVar(value=True)
        self.case_var = tk.BooleanVar(value=False)
        self.both_var = tk.BooleanVar(value=False)
        c1 = ttk.Frame(f)
        c1.pack(fill='x', pady=2)
        ttk.Checkbutton(c1, text='创建备份', variable=self.backup_var).pack(side='left', padx=4)
        ttk.Checkbutton(c1, text='包含子文件夹', variable=self.recursive_var).pack(side='left', padx=4)
        ttk.Checkbutton(c1, text='区分大小写', variable=self.case_var).pack(side='left', padx=4)
        c2 = ttk.Frame(f)
        c2.pack(fill='x', pady=2)
        ttk.Label(c2, text='文件类型：').pack(side='left', padx=4)
        ttk.Radiobutton(c2, text='仅 .docx', variable=self.both_var, value=False).pack(side='left', padx=4)
        ttk.Radiobutton(c2, text='含 .doc（旧版仅跳过）', variable=self.both_var, value=True).pack(side='left', padx=4)

        # 高级选项
        f = LF('高级选项')
        self.regex_var = tk.BooleanVar(value=False)
        self.headers_var = tk.BooleanVar(value=True)
        self.footers_var = tk.BooleanVar(value=True)
        self.tables_var = tk.BooleanVar(value=True)
        self.backup_fmt_var = tk.StringVar(value='backup_%Y%m%d_%H%M%S')
        ttk.Checkbutton(f, text='使用正则表达式', variable=self.regex_var,
                        command=self._toggle_regex).pack(anchor='w', pady=2)
        rg = ttk.Frame(f)
        rg.pack(fill='x', pady=2)
        ttk.Label(rg, text='正则式：', width=10).pack(side='left')
        self.regex_pat_var = tk.StringVar(value=r'')
        ttk.Entry(rg, textvariable=self.regex_pat_var).pack(side='left', fill='x', expand=True)
        ttk.Button(rg, text='验证正则', command=self._validate_regex,
                   style='Link.TButton').pack(side='left', padx=(6, 0))
        # 正则示例
        ex = ttk.Frame(f)
        ex.pack(fill='x', pady=(2, 4))
        ttk.Label(ex, text='示例：', width=10, anchor='n').pack(side='left')
        ex_box = tk.Text(ex, wrap='word', height=3, font='AppMono',
                         relief='sunken', borderwidth=1, padx=4, pady=2)
        ex_box.insert('1.0',
                      '故障\\d+个        匹配“故障”后跟数字\n'
                      '第[一二三]周      匹配中文周数\n'
                      '张[三省]          匹配张三/张省')
        ex_box.config(state='disabled')
        ex_box.pack(side='left', fill='x', expand=True)
        sc = ttk.Frame(f)
        sc.pack(fill='x', pady=2)
        ttk.Checkbutton(sc, text='处理页眉', variable=self.headers_var).pack(side='left', padx=4)
        ttk.Checkbutton(sc, text='处理页脚', variable=self.footers_var).pack(side='left', padx=4)
        ttk.Checkbutton(sc, text='处理表格', variable=self.tables_var).pack(side='left', padx=4)
        bn = ttk.Frame(f)
        bn.pack(fill='x', pady=2)
        ttk.Label(bn, text='备份名：', width=10).pack(side='left')
        ttk.Entry(bn, textvariable=self.backup_fmt_var, width=30).pack(side='left', padx=(0, 6))
        ttk.Label(bn, text='（如 backup_%Y%m%d_%H%M%S）', foreground='#888').pack(side='left')

        # 处理统计
        f = LF('处理统计')
        self.stat_vars = {
            'total': tk.StringVar(value='0'),
            'processed': tk.StringVar(value='0'),
            'modified': tk.StringVar(value='0'),
            'errors': tk.StringVar(value='0'),
        }
        sc = ttk.Frame(f)
        sc.pack(fill='x', pady=2)
        for key, label in [('total', '总文件'), ('processed', '已处理'),
                           ('modified', '已修改'), ('errors', '错误')]:
            ttk.Label(sc, text=label + '：', width=8).pack(side='left', padx=2)
            ttk.Label(sc, textvariable=self.stat_vars[key], width=6,
                      font='AppBold').pack(side='left', padx=(0, 14))

        # 按钮行
        bf = ttk.Frame(inner)
        bf.pack(fill='x', padx=8, pady=10)
        self.start_btn = ttk.Button(bf, text='▶ 开始替换', command=self._on_run, style='Accent.TButton')
        self.start_btn.pack(side='left', padx=(0, 6))
        self.stop_btn = ttk.Button(bf, text='⏹ 停止', command=self._on_stop, state=tk.DISABLED)
        self.stop_btn.pack(side='left')
        ttk.Button(bf, text='🔍 测试查找', command=self._on_test).pack(side='left', padx=6)

        # 弹性占位：窗口放大时把多余空白收拢到最底部，避免中间突兀留白
        spacer = tk.Frame(inner, height=1)
        spacer.pack(fill='both', expand=True)

    def _pick(self):
        p = filedialog.askdirectory(title='选择包含 Word 文档的文件夹')
        if p:
            self.folder_var.set(p)

    def _toggle_regex(self):
        if self.regex_var.get():
            self.find_var.set(self.regex_pat_var.get())

    def _validate_regex(self):
        pat = self.regex_pat_var.get()
        if not pat:
            messagebox.showwarning('提示', '请先填写正则表达式。')
            return
        try:
            re.compile(pat)
            messagebox.showinfo('验证通过', '正则表达式语法正确：\n' + pat)
        except re.error as e:
            messagebox.showerror('正则错误', '表达式有误：\n' + pat + '\n\n' + str(e))

    def _set_stat(self, key, value):
        try:
            self.stat_vars[key].set(str(value))
        except Exception:
            pass

    def _resolve(self):
        return {
            'folder': self.folder_var.get().strip(),
            'find': self.find_var.get(),
            'replace': self.replace_var.get(),
            'use_regex': bool(self.regex_var.get()),
            'regex': self.regex_pat_var.get(),
            'case_sensitive': bool(self.case_var.get()),
            'recursive': bool(self.recursive_var.get()),
            'both': bool(self.both_var.get()),
            'backup': bool(self.backup_var.get()),
            'backup_fmt': self.backup_fmt_var.get().strip() or 'backup_%Y%m%d_%H%M%S',
            'headers': bool(self.headers_var.get()),
            'footers': bool(self.footers_var.get()),
            'tables': bool(self.tables_var.get()),
        }

    def _on_run(self):
        p = self._resolve()
        if not p['folder'] or not os.path.isdir(p['folder']):
            self.app.log('⚠ 请先选择有效的文件夹。')
            return
        old = p['regex'] if p['use_regex'] else p['find']
        if not old:
            self.app.log('⚠ 查找内容不能为空。')
            return
        self.app.execute(self)

    def _on_stop(self):
        self.stop_req = True
        self.app.log('⏹ 已请求停止，处理完当前文件后结束…')

    def _set_buttons(self, running):
        self.start_btn.config(state=tk.DISABLED if running else tk.NORMAL)
        self.stop_btn.config(state=tk.NORMAL if running else tk.DISABLED)

    def _on_test(self):
        p = self._resolve()
        if not p['folder'] or not os.path.isdir(p['folder']):
            self.app.log('⚠ 请先选择有效的文件夹。')
            return
        old = p['regex'] if p['use_regex'] else p['find']
        if not old:
            self.app.log('⚠ 查找内容不能为空。')
            return
        threading.Thread(target=self._test_thread, args=(
            p['folder'], old, p['both'], p['use_regex'], p['case_sensitive']), daemon=True).start()

    def _test_thread(self, folder, old, both, use_regex, case_sensitive):
        exts = ('.docx', '.doc') if both else ('.docx',)
        found = []
        for root, dirs, files in os.walk(folder):
            if 'backup_' in root:
                continue
            for fn in files:
                if fn.lower().endswith(exts):
                    fp = os.path.join(root, fn)
                    try:
                        doc = Document(fp)
                        text = '\n'.join(pg.text for pg in doc.paragraphs[:50])
                        if self._match(text, old, use_regex, case_sensitive):
                            found.append(fp)
                    except Exception:
                        pass
        self.app.log(f'测试查找完成：{len(found)} 个文件包含目标文本。')
        if found:
            self.app.root.after(0, self._show_results, found)

    def _show_results(self, files):
        win = tk.Toplevel(self.app.root)
        win.title('查找结果')
        win.geometry('560x380')
        ttk.Label(win, text=f'找到 {len(files)} 个文件包含目标文本：',
                  font=('微软雅黑', 10, 'bold')).pack(pady=10)
        frame = ttk.Frame(win)
        frame.pack(fill='both', expand=True, padx=10, pady=10)
        tree = ttk.Treeview(frame, columns=('文件名', '路径'), show='headings', height=14)
        tree.heading('文件名', text='文件名')
        tree.heading('路径', text='路径')
        tree.column('文件名', width=160)
        tree.column('路径', width=340)
        for fp in files:
            tree.insert('', tk.END, values=(os.path.basename(fp), os.path.dirname(fp)))
        tree.pack(side='left', fill='both', expand=True)
        ttk.Scrollbar(frame, orient='vertical', command=tree.yview).pack(side='right', fill='y')
        tree.configure(yscrollcommand=lambda *a: None)
        ttk.Button(win, text='关闭', command=win.destroy).pack(pady=10)

    def prepare(self, app):
        self.params = self._resolve()
        self.stop_req = False

    def run(self, app):
        p = self.params
        self.app.root.after(0, self._set_buttons, True)
        for k in ('total', 'processed', 'modified', 'errors'):
            self.app.root.after(0, self._set_stat, k, 0)
        folder = p['folder']
        old = p['regex'] if p['use_regex'] else p['find']
        new = p['replace']
        if not folder or not os.path.isdir(folder):
            app.log('⚠ 请先选择有效的文件夹。')
            self.app.root.after(0, self._set_buttons, False)
            return
        if not old:
            app.log('⚠ 查找内容不能为空。')
            self.app.root.after(0, self._set_buttons, False)
            return

        exts = ('.docx', '.doc') if p['both'] else ('.docx',)
        files = []
        if p['recursive']:
            for root, dirs, fs in os.walk(folder):
                if 'backup_' in root:
                    continue
                for fn in fs:
                    if fn.lower().endswith(exts):
                        files.append(os.path.join(root, fn))
        else:
            files = [os.path.join(folder, fn) for fn in os.listdir(folder)
                     if fn.lower().endswith(exts)]
        if not files:
            app.log('⚠ 该文件夹（及子文件夹）中未找到 Word 文档。')
            self.app.root.after(0, self._set_buttons, False)
            return

        self.app.root.after(0, self._set_stat, 'total', len(files))
        app.log(f'找到 {len(files)} 个 Word 文档，开始替换…')
        backup_dir = None
        if p['backup']:
            fmt = p['backup_fmt']
            dname = datetime.now().strftime(fmt) if '%' in fmt else fmt
            backup_dir = os.path.join(folder, dname)
            os.makedirs(backup_dir, exist_ok=True)
            app.log(f'已创建备份文件夹：{backup_dir}')

        total = len(files)
        ok = err = 0
        for i, fp in enumerate(files, 1):
            if self.stop_req:
                app.log('⏹ 用户已请求停止，处理已中断。')
                break
            name = os.path.basename(fp)
            try:
                modified = self._process_file(
                    fp, old, new, backup_dir, p['use_regex'], p['case_sensitive'],
                    p['headers'], p['footers'], p['tables'])
                if modified:
                    ok += 1
                    app.log(f'    ✓ 已修改：{name}')
                else:
                    app.log(f'    · 未匹配：{name}')
            except Exception as e:
                if fp.lower().endswith('.doc'):
                    app.log(f'    ⚠ 跳过旧版 .doc：{name}（python-docx 不支持，请另存为 .docx）')
                else:
                    err += 1
                    app.log(f'    ✗ 错误：{name} - {str(e)[:150]}')
            app.set_progress(i, total)
            self.app.root.after(0, self._set_stat, 'processed', i)
            self.app.root.after(0, self._set_stat, 'modified', ok)
            self.app.root.after(0, self._set_stat, 'errors', err)

        app.log(f'替换完成：共 {total} 个文件，已修改 {ok} 个，错误 {err} 个。')
        app.finish_folder(folder)
        self.app.root.after(0, self._set_buttons, False)


# ============================================================
# 主程序：实用工具
# ============================================================
class App:
    def __init__(self):
        cleanup_mei()   # 启动清理历史残留 _MEI 临时目录
        self.root = tk.Tk()
        self.root.title('实用工具')
        try:
            self.root.iconbitmap(_icon_path())   # 窗口标题栏 / 任务栏图标
        except Exception:
            pass
        self.root.geometry('1220x680')
        self.root.minsize(1040, 600)
        self._running = False
        self.last_output = None
        self.last_output_is_folder = False
        self._apply_styles()
        self._build_ui()
        self.tools = [FolderTool(), ExcelTool(), DocxReplaceTool(), PrintTool(),
                      LinkTool(), TranslateTool()]
        self._current = None
        self._build_nav()
        self._select(0)
        self.log('欢迎使用「实用工具」v' + APP_VERSION + ' —— 选择左侧功能，配置后点「开始」。')
        self.log('已整合：文件夹图片→Word / Excel图片→Word / Word批量替换 / 批量打印 / PDF工具 / 沉浸式翻译。')
        self.log('右侧栏实时显示「成品」与「历史记录」，双击列表项即可打开对应成品。')
        self.log('每个功能顶部都有「❔ 功能介绍」按钮，点开即可看详细说明。')
        self.log('顶部「检查更新」可手动检查；每次启动也会后台静默检查新版本。')
        # 启动后后台静默检查更新
        self.root.after(1500, lambda: self._check_update(manual=False))

    # ---------- 样式 ----------
    def _apply_styles(self):
        setup_fonts(_FONT_SCALE)
        try:
            s = ttk.Style()
            s.theme_use('clam')
            # 全局默认字体（命名字体，菜单可实时缩放）
            FONT = 'AppBase'
            s.configure('.', font=FONT)
            s.configure('TLabel', font=FONT)
            s.configure('TButton', font=FONT)
            s.configure('TCheckbutton', font=FONT)
            s.configure('TRadiobutton', font=FONT)
            s.configure('TCombobox', font=FONT)
            s.configure('TSpinbox', font=FONT)
            s.configure('TEntry', font=FONT)
            s.configure('TLabelframe.Label', font='AppBold',
                        foreground='#34495e')
            s.configure('Accent.TButton', background='#4CAF50', foreground='white',
                        font='AppBold')
            s.map('Accent.TButton', background=[('active', '#43a047'),
                                                ('disabled', '#a5d6a7')])
            s.configure('Nav.TButton', font='AppBase')
            s.configure('Link.TButton', background='#eef3f8', foreground='#1a73e8',
                        font='AppSmall')
            s.map('Link.TButton', background=[('active', '#dbe7f3')])
            s.configure('Top.TButton', background='#34495e', foreground='white',
                        font='AppSmall')
            s.map('Top.TButton', background=[('active', '#3d566e'),
                                             ('disabled', '#6b7a8a')])
        except Exception:
            pass

    # ---------- 布局 ----------
    def _build_ui(self):
        # 顶部标题栏
        top = tk.Frame(self.root, bg='#2c3e50', height=56)
        top.pack(side='top', fill='x')
        top.pack_propagate(False)
        tk.Label(top, text='实用工具', bg='#2c3e50', fg='white',
                 font='AppTitle').pack(side='left', padx=16)
        tk.Label(top, text='v1.4.3 · 集成多个小工具', bg='#2c3e50', fg='#aab7c4',
                 font='AppSmall').pack(side='left', padx=(4, 0))
        self.top_bar = top
        self._build_update_ui()

        # 版本更新横幅（默认隐藏，发现新版本时显示）
        self._build_banner()

        # 主体：左导航 + 右内容
        body = tk.Frame(self.root)
        body.pack(side='top', fill='both', expand=True)

        self.nav_frame = tk.Frame(body, width=210, bg='#f4f6f8')
        self.nav_frame.pack(side='left', fill='y')
        self.nav_frame.pack_propagate(False)
        tk.Label(self.nav_frame, text='功能列表', bg='#f4f6f8', fg='#34495e',
                 font='AppNavTitle').pack(anchor='w', padx=14, pady=(12, 6))

        self.panel = tk.Frame(body, bg='white')
        self.panel.pack(side='left', fill='both', expand=True)

        # 右侧栏：成品 + 历史记录
        self.products = []   # [{name, path, is_folder}]
        self.history = []    # [{time, tool, desc, path, is_folder}]
        self.side_frame = tk.Frame(body, width=266, bg='#f4f6f8')
        self.side_frame.pack(side='right', fill='y')
        self.side_frame.pack_propagate(False)

        # 成品区
        lf_p = tk.LabelFrame(self.side_frame, text='📦 成品', bg='#f4f6f8',
                             fg='#34495e', font='AppBold')
        lf_p.pack(side='top', fill='both', expand=True, padx=6, pady=(8, 4))
        pf = tk.Frame(lf_p)
        pf.pack(fill='both', expand=True, padx=4, pady=4)
        self.prod_list = tk.Listbox(pf, font='AppSmall', bg='white',
                                    relief='sunken', borderwidth=1, activestyle='none')
        psb = ttk.Scrollbar(pf, command=self.prod_list.yview)
        psb.pack(side='right', fill='y')
        self.prod_list.config(yscrollcommand=psb.set)
        self.prod_list.pack(side='left', fill='both', expand=True)
        self.prod_list.bind('<Double-Button-1>',
                            lambda e: self._open_list(self.prod_list, self.products))
        self.prod_list.insert(tk.END, '（暂无成品，运行功能后会显示在这里）')

        # 历史记录区
        lf_h = tk.LabelFrame(self.side_frame, text='🕘 历史记录', bg='#f4f6f8',
                             fg='#34495e', font='AppBold')
        lf_h.pack(side='top', fill='both', expand=True, padx=6, pady=(4, 8))
        hf = tk.Frame(lf_h)
        hf.pack(fill='both', expand=True, padx=4, pady=4)
        self.hist_list = tk.Listbox(hf, font='AppSmall', bg='white',
                                    relief='sunken', borderwidth=1, activestyle='none')
        hsb = ttk.Scrollbar(hf, command=self.hist_list.yview)
        hsb.pack(side='right', fill='y')
        self.hist_list.config(yscrollcommand=hsb.set)
        self.hist_list.pack(side='left', fill='both', expand=True)
        self.hist_list.bind('<Double-Button-1>',
                            lambda e: self._open_list(self.hist_list, self.history))
        self.hist_list.insert(tk.END, '（暂无记录）')

        # 底部：进度 + 日志 + 状态
        bottom = tk.Frame(self.root, bg='#fafafa', height=210)
        bottom.pack(side='bottom', fill='x')
        bottom.pack_propagate(False)

        bar_row = tk.Frame(bottom, bg='#fafafa')
        bar_row.pack(side='top', fill='x', padx=12, pady=(8, 0))
        self.progress = ttk.Progressbar(bar_row, orient='horizontal',
                                        mode='determinate', maximum=1)
        self.progress.pack(side='left', fill='x', expand=True, padx=(0, 8))
        self.open_btn = ttk.Button(bar_row, text='打开生成的 Word',
                                   command=self._open_result, state=tk.DISABLED)
        self.open_btn.pack(side='left')
        self.open_folder_btn = ttk.Button(bar_row, text='打开Word所在文件夹',
                                          command=self._open_folder, state=tk.DISABLED)
        self.open_folder_btn.pack(side='left', padx=(6, 0))
        self.clear_btn = ttk.Button(bar_row, text='清空日志', command=self._clear_log)
        self.clear_btn.pack(side='left', padx=(6, 0))

        log_row = tk.Frame(bottom, bg='#fafafa')
        log_row.pack(side='top', fill='both', expand=True, padx=12, pady=(6, 8))
        self.log_text = tk.Text(log_row, height=8, wrap='word',
                                font='AppMono', bg='white',
                                relief='sunken', borderwidth=1, state=tk.DISABLED)
        self.log_text.pack(side='left', fill='both', expand=True)
        sb = ttk.Scrollbar(log_row, command=self.log_text.yview)
        sb.pack(side='right', fill='y')
        self.log_text.config(yscrollcommand=sb.set)

        self.status_var = tk.StringVar(value='就绪')
        st = tk.Label(self.root, textvariable=self.status_var, anchor='w',
                      relief='sunken', bg='#ecf0f1', font='AppSmall')
        st.pack(side='bottom', fill='x')

        # 顶部菜单栏（设置 -> 字体大小：实时缩放全局界面）
        self._build_menu()

    # ---------- 菜单 ----------
    def _build_menu(self):
        self.font_scale_var = tk.DoubleVar(value=_FONT_SCALE)
        menubar = tk.Menu(self.root)
        m_set = tk.Menu(menubar, tearoff=0)
        m_font = tk.Menu(m_set, tearoff=0)
        for label, sc in [('小（约 9 号）', 0.9), ('标准（约 10 号）', 1.0),
                          ('大（约 12 号）', 1.2), ('特大（约 14 号）', 1.4),
                          ('超大（约 16 号）', 1.6)]:
            m_font.add_radiobutton(label=label, value=sc, variable=self.font_scale_var,
                                  command=lambda s=sc: self.set_font_scale(s))
        m_set.add_cascade(label='字体大小', menu=m_font)
        m_set.add_separator()
        m_set.add_command(label='恢复默认字体', command=lambda: self.set_font_scale(1.0))
        m_set.add_command(label='关于「实用工具」', command=self._show_about)
        menubar.add_cascade(label='设置', menu=m_set)
        self.root.config(menu=menubar)

    def set_font_scale(self, scale):
        """实时缩放全局字体：重新配置命名字体，所有引用它的控件即时更新。"""
        setup_fonts(scale)
        try:
            self.font_scale_var.set(scale)
        except Exception:
            pass
        tag = '默认' if abs(scale - 1.0) < 1e-6 else ('%.1fx' % scale)
        self.log('字体大小已调整：%s。' % tag)

    def _show_about(self):
        win = tk.Toplevel(self.root)
        win.title('关于「实用工具」')
        win.geometry('420x200')
        win.transient(self.root)
        win.grab_set()
        tk.Label(win, text='实用工具', font='AppTitle', fg='#2c3e50').pack(pady=(18, 4))
        tk.Label(win, text='版本 v' + APP_VERSION, font='AppBase', fg='#555').pack()
        tk.Label(win, text='集成多项小工具：文件夹图片→Word / Excel图片→Word /\n'
                           'Word批量替换 / 批量打印(PDF) / PDF工具',
                 font='AppSmall', fg='#777', justify='center').pack(pady=(10, 4))
        ttk.Button(win, text='关闭', command=win.destroy,
                   style='Accent.TButton').pack(pady=(8, 0))

    # ---------- 导航 ----------
    def _build_nav(self):
        self.nav_buttons = []
        for i, t in enumerate(self.tools):
            b = tk.Button(self.nav_frame, text=t.name, anchor='w',
                          relief='flat', bg='#f4f6f8', fg='#2c3e50',
                          font='AppBase', padx=14, pady=8,
                          command=lambda idx=i: self._select(idx))
            b.pack(side='top', fill='x', padx=6, pady=2)
            self.nav_buttons.append(b)

    def _select(self, idx):
        self._current = self.tools[idx]
        for i, b in enumerate(self.nav_buttons):
            b.config(bg='#4CAF50' if i == idx else '#f4f6f8',
                     fg='white' if i == idx else '#2c3e50')
        for w in self.panel.winfo_children():
            w.destroy()
        # 顶部：功能名 + 功能介绍按钮（介绍改为点击弹窗，不默认展开）
        head = tk.Frame(self.panel, bg='#eef3f8')
        head.pack(side='top', fill='x')
        title_row = tk.Frame(head, bg='#eef3f8')
        title_row.pack(fill='x', padx=12, pady=(8, 6))
        tk.Label(title_row, text=self._current.name, bg='#eef3f8', fg='#1f3a5f',
                 font='AppBig').pack(side='left')
        ttk.Button(title_row, text='❔ 功能介绍', style='Link.TButton',
                   command=lambda: self.show_intro(self._current)).pack(side='right')
        # 功能区（各功能自己的控件）—— 放在可滚动画布里，
        # 这样选项很多时（如 Word 批量替换）底部按钮也不会被遮挡。
        # 同时让内容宽度/高度跟随画布：放大窗口时自动铺满，不再留下大片空地。
        canvas = tk.Canvas(self.panel, bg='white', highlightthickness=0)
        vsb = ttk.Scrollbar(self.panel, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        area = tk.Frame(canvas, bg='white')

        def _on_configure(_e=None):
            # 内容比画布小时，把内框撑满画布（铺满、不留空地）；
            # 内容比画布大时，按实际高度滚动。
            cw, ch = canvas.winfo_width(), canvas.winfo_height()
            iw, ih = area.winfo_reqwidth(), area.winfo_reqheight()
            w = max(cw, iw)
            h = max(ch, ih)
            canvas.itemconfig(area_id, width=w, height=h)
            canvas.configure(scrollregion=(0, 0, w, h))

        area_id = canvas.create_window((0, 0), window=area, anchor='nw')
        area.bind('<Configure>', _on_configure)
        canvas.bind('<Configure>', _on_configure)
        # 鼠标滚轮滚动（切换功能时先解绑旧的，避免重复绑定）
        old = getattr(self, '_wheel_id', None)
        if old:
            try:
                self.root.unbind('<MouseWheel>', old)
            except Exception:
                pass
        self._wheel_id = canvas.bind_all(
            '<MouseWheel>',
            lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units'))
        self._current.build(area, self)

    def show_intro(self, tool):
        """点「功能介绍」按钮后弹出的独立窗口。内容来自 Tool.intro。"""
        intro = getattr(tool, 'intro', '')
        win = tk.Toplevel(self.root)
        win.title(tool.name + ' — 功能介绍')
        win.geometry('580x440')
        win.transient(self.root)
        try:
            win.tk.call('tk', 'scaling', 1.0)  # 与父窗口一致的清晰缩放
        except Exception:
            pass
        # 居中显示在应用主窗口中央（而非屏幕左上角）
        self.root.update_idletasks()
        win.update_idletasks()
        rw, rh = self.root.winfo_width(), self.root.winfo_height()
        rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
        ww, wh = 580, 440
        x = max(0, rx + (rw - ww) // 2)
        y = max(0, ry + (rh - wh) // 2)
        win.geometry('+%d+%d' % (x, y))
        # 标题条
        bar = tk.Frame(win, bg='#2c3e50')
        bar.pack(fill='x')
        tk.Label(bar, text=tool.name, bg='#2c3e50', fg='white',
                 font='AppIntroTitle', anchor='w', padx=14, pady=10).pack(fill='x')
        # 内容（可滚动）
        f = tk.Frame(win, bg='white')
        f.pack(fill='both', expand=True, padx=12, pady=10)
        txt = tk.Text(f, wrap='word', font='AppIntro', relief='flat',
                      padx=8, pady=8, bg='white', state='disabled',
                      spacing1=3, spacing2=2)
        sb = ttk.Scrollbar(f, command=txt.yview)
        txt.config(yscrollcommand=sb.set)
        txt.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')
        txt.config(state='normal')
        txt.insert('1.0', intro or '（暂无介绍）')
        txt.config(state='disabled')
        ttk.Button(win, text='关闭', command=win.destroy).pack(pady=(0, 12))

    # ---------- 执行控制 ----------
    def execute(self, tool):
        if self._running:
            self.log('⚠ 有任务正在运行，请稍候…')
            return
        tool.prepare(self)   # 主线程读取 UI
        self._running = True
        self.open_btn.config(state=tk.DISABLED)
        self.open_folder_btn.config(state=tk.DISABLED)
        self.last_output_is_folder = False
        self.set_progress(0)
        self.set_status('正在执行：' + tool.name)
        threading.Thread(target=self._worker, args=(tool,), daemon=True).start()

    def _worker(self, tool):
        try:
            tool.run(self)
            self.set_status('完成')
        except Exception:
            self.log('❌ 执行出错：\n' + traceback.format_exc())
            self.set_status('执行出错')
        finally:
            self._running = False

    # ---------- 对外回调（供 Tool.run 调用）----------
    def log(self, msg):
        self.root.after(0, self._append_log, msg)

    def _append_log(self, msg):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + '\n')
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete('1.0', tk.END)
        self.log_text.config(state=tk.DISABLED)

    def set_progress(self, value, maximum=None):
        self.root.after(0, self._set_progress, value, maximum)

    def _set_progress(self, value, maximum):
        if maximum is not None:
            self.progress['maximum'] = maximum
        self.progress['value'] = value
        if self.progress['maximum']:
            pct = int(value / self.progress['maximum'] * 100)
            self.status_var.set(f'进度 {pct}%')
        else:
            self.status_var.set('就绪')

    def set_status(self, msg):
        self.root.after(0, lambda: self.status_var.set(msg))

    def make_progress_cb(self, total):
        def cb(done, total=total, name=''):
            self.set_progress(done, total)
            if name:
                self.log(f'  [{done}/{total}] {name}')
        return cb

    def finish(self, path):
        self.last_output = path
        self.last_output_is_folder = False
        self.root.after(0, lambda: [self.open_btn.config(state=tk.NORMAL),
                                     self.open_folder_btn.config(state=tk.NORMAL)])
        self.log('✅ 已生成：' + path)
        self.add_history(self._current.name, path=path, is_folder=False)

    def finish_folder(self, folder):
        """结果是一个文件夹（如批量替换的多份文档）：仅启用「打开文件夹」按钮。"""
        self.last_output = folder
        self.last_output_is_folder = True
        self.root.after(0, lambda: self.open_folder_btn.config(state=tk.NORMAL))
        self.log('✅ 结果文件夹：' + folder)
        self.add_history(self._current.name, path=folder, is_folder=True)

    def add_history(self, tool_name, path=None, is_folder=False, note=None):
        """登记一条历史记录；若 path 存在，同时进入「成品」列表。
        由子线程调用也安全：数据先更新，UI 刷新走主线程 after(0)。"""
        now = datetime.now().strftime('%m-%d %H:%M')
        if note:
            desc = note
        elif path:
            desc = os.path.basename(path.rstrip('/\\')) or path
        else:
            desc = '完成'
        self.history.insert(0, {'time': now, 'tool': tool_name, 'desc': desc,
                                'path': path, 'is_folder': is_folder})
        if path and os.path.exists(path):
            self.products.insert(0, {'name': desc, 'path': path, 'is_folder': is_folder})
        self.root.after(0, self._refresh_side_lists)

    def _refresh_side_lists(self):
        self.prod_list.delete(0, tk.END)
        if not self.products:
            self.prod_list.insert(tk.END, '（暂无成品，运行功能后会显示在这里）')
        else:
            for p in self.products:
                icon = '📁' if p['is_folder'] else '📄'
                self.prod_list.insert(tk.END, f"{icon} {p['name']}")
        self.hist_list.delete(0, tk.END)
        if not self.history:
            self.hist_list.insert(tk.END, '（暂无记录）')
        else:
            for h in self.history:
                icon = '📁' if h['is_folder'] else ('📄' if h['path'] else '⚙')
                self.hist_list.insert(tk.END, f"{h['time']} {icon}[{h['tool']}] {h['desc'][:16]}")

    def _open_list(self, listbox, store):
        if not store:
            return
        sel = listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx >= len(store):
            return
        item = store[idx]
        if item.get('path') and os.path.exists(item['path']):
            try:
                os.startfile(item['path'])
                self.log('已打开：' + item['path'])
            except Exception:
                self.log('无法自动打开：' + item['path'])
        else:
            self.log('该记录没有可打开的成品文件（如批量打印只记录结果，不生成文件）。')

    def _open_result(self):
        if self.last_output and os.path.exists(self.last_output):
            try:
                os.startfile(self.last_output)
            except Exception:
                self.log('（无法自动打开，请手动打开文件）')

    def _open_folder(self):
        if self.last_output and os.path.exists(self.last_output):
            d = (self.last_output if getattr(self, 'last_output_is_folder', False)
                 else os.path.dirname(os.path.abspath(self.last_output)))
            try:
                os.startfile(d)
            except Exception:
                self.log('（无法自动打开文件夹，路径：' + d + '）')


    # ---------- 版本更新 ----------
    def _build_update_ui(self):
        # 顶部抬头右侧：版本号(可点) + 检查更新 + 状态 + 恢复默认更新地址
        f = tk.Frame(self.top_bar, bg='#2c3e50')
        f.pack(side='right', padx=14)
        self.ver_label = tk.Label(f, text='v' + APP_VERSION, bg='#2c3e50', fg='#cfe3ff',
                                  font='AppBase', cursor='hand2')
        self.ver_label.pack(side='left', padx=(0, 8))
        self.ver_label.bind('<Button-1>', lambda e: self._check_update(manual=True))
        self.update_btn = ttk.Button(f, text='检查更新', style='Top.TButton',
                                     command=lambda: self._check_update(manual=True))
        self.update_btn.pack(side='left', padx=(0, 8))
        self.update_status = tk.Label(f, text='', bg='#2c3e50', fg='#aab7c4',
                                      font='AppSmall')
        self.update_status.pack(side='left', padx=(0, 8))
        self.restore_src_btn = ttk.Button(f, text='恢复默认更新地址', style='Top.TButton',
                                          command=self._restore_default_source)
        # 仅当设置了自定义地址才显示
        cfg = load_update_config()
        if cfg.get('source'):
            self.restore_src_btn.pack(side='left', padx=(0, 4))

    def _build_banner(self):
        b = tk.Frame(self.root, bg='#fff3cd', relief='solid', borderwidth=1)
        b.pack(side='top', fill='x')
        b.pack_forget()   # 默认隐藏
        self.banner = b
        inner = tk.Frame(b, bg='#fff3cd')
        inner.pack(fill='x', padx=12, pady=6)
        self.banner_text = tk.Label(inner, text='', bg='#fff3cd', fg='#7a5b00',
                                    font='AppBold')
        self.banner_text.pack(side='left')
        self.banner_notes = tk.Label(inner, text='', bg='#fff3cd', fg='#8a6d00',
                                     font='AppSmall')
        self.banner_notes.pack(side='left', padx=10)
        self.banner_stage = tk.Label(inner, text='', bg='#fff3cd', fg='#8a6d00',
                                     font='AppSmall')
        self.banner_stage.pack(side='left', padx=10)
        self.banner_pb = ttk.Progressbar(inner, orient='horizontal', mode='determinate',
                                         maximum=1)
        # 进度条默认不显示，下载时才 pack
        bf = tk.Frame(inner, bg='#fff3cd')
        bf.pack(side='right')
        self.btn_now = ttk.Button(bf, text='立即更新', style='Accent.TButton',
                                  command=self._start_download)
        self.btn_now.pack(side='left', padx=4)
        self.btn_skip = ttk.Button(bf, text='跳过此版本', command=self._on_skip)
        self.btn_skip.pack(side='left', padx=4)
        self.btn_later = ttk.Button(bf, text='稍后再说', command=self._on_later)
        self.btn_later.pack(side='left', padx=4)

    def _check_update(self, manual=False):
        cfg = load_update_config()
        source = cfg.get('source') or UPDATE_MANIFEST_PRIMARY
        urls = _mirror_urls(source)
        self.update_status.config(text='检查更新中…', fg='#aab7c4')
        self.update_btn.config(state=tk.DISABLED)
        if manual:
            self.log('手动检查更新…')
        else:
            self.log('启动后台检查更新…')

        def worker():
            if _UPDATE_SKIP_NETWORK:
                self.root.after(0, self._on_manifest, None, None, manual)
                return
            d, hit = _fetch_manifest(urls)
            self.root.after(0, self._on_manifest, d, hit, manual)
        threading.Thread(target=worker, daemon=True).start()

    def _on_manifest(self, manifest, hit_source, manual):
        self.update_btn.config(state=tk.NORMAL)
        if not manifest:
            self.update_status.config(text='检查失败（网络/地址）', fg='#aab7c4')
            if manual:
                messagebox.showwarning('检查更新', '暂时无法连接更新服务器，请稍后再试。')
            return
        latest = manifest.get('version', '0')
        self.update_status.config(text='线上 v%s · %s' % (latest, self._short_source(hit_source)),
                                  fg='#aab7c4')
        if cmp_version(latest, APP_VERSION) > 0:
            self.update_status.config(fg='#ffd27f')   # 比本机旧版本号 → 橙色 stale 提示
            cfg = load_update_config()
            # 跳过此版本：仅启动检查生效，且只对该版本
            if (not manual) and cfg.get('skip_version') == latest:
                return
            self._show_banner(manifest, hit_source, manual)
        else:
            if manual:
                messagebox.showinfo('检查更新', '当前已是最新版本（v%s）。' % APP_VERSION)

    @staticmethod
    def _short_source(url):
        if not url:
            return ''
        if 'ghproxy.net' in url:
            return 'ghproxy'
        if 'gh-proxy.com' in url:
            return 'gh-proxy'
        if 'github.com' in url and '/raw/' in url:
            return 'github/raw'
        return 'raw'

    def _show_banner(self, manifest, hit_source, manual):
        self.banner_version = manifest.get('version')
        self.banner_manifest = manifest
        self.banner_hit = hit_source
        notes = manifest.get('notes', '')
        self.banner_text.config(text='发现新版本 v%s%s' % (
            self.banner_version, '（重要更新）' if manifest.get('mandatory') else ''))
        self.banner_notes.config(text=(notes[:120] + ('…' if len(notes) > 120 else '')) if notes else '')
        self.banner_stage.config(text='')
        self.banner_pb.pack_forget()
        self.banner.pack(fill='x')
        # 「跳过此版本」仅对启动检查生效
        self.btn_skip.config(state=tk.DISABLED if manual else tk.NORMAL)

    def _start_download(self):
        m = getattr(self, 'banner_manifest', None)
        if not m:
            return
        self.banner_pb.pack(fill='x', padx=12, pady=(2, 6))
        self.banner_pb['value'] = 0
        self.banner_pb['maximum'] = 1
        self.banner_stage.config(text='准备下载…')
        self.btn_now.config(state=tk.DISABLED)
        self.btn_skip.config(state=tk.DISABLED)
        self.btn_later.config(state=tk.DISABLED)
        self.log('开始下载更新包…')

        def on_progress(done, total):
            self.root.after(0, lambda: [
                self.banner_pb.config(maximum=total or 1, value=done),
                self.banner_stage.config(text='下载中 %.0f%%' % (done / (total or 1) * 100))])

        def on_stage(t):
            self.root.after(0, lambda: self.banner_stage.config(text=t))

        def worker():
            dest = os.path.join(_app_dir(), '_update_tmp.exe')
            try:
                if os.path.isfile(dest):
                    os.remove(dest)
            except Exception:
                pass
            ok, msg = download_update(
                m.get('download_url'), dest,
                sha256=m.get('sha256'), size=m.get('size'),
                on_progress=on_progress, on_stage=on_stage)
            self.root.after(0, self._on_download_done, ok, msg, dest)
        threading.Thread(target=worker, daemon=True).start()

    def _on_download_done(self, ok, msg, dest):
        if not ok:
            self.banner_stage.config(text='下载失败：' + msg)
            self.btn_now.config(state=tk.NORMAL)
            self.btn_later.config(state=tk.NORMAL)
            self.btn_skip.config(state=tk.NORMAL)
            messagebox.showerror('更新失败', '下载未成功：' + msg + '\n可点「立即更新」重试。')
            return
        self.banner_stage.config(text='下载完成，正在安装替换…')
        self.log('更新包下载完成，开始安装替换…')
        ok2, msg2 = install_update(dest)
        if not ok2:
            self.banner_stage.config(text='安装失败：' + msg2)
            self.btn_now.config(state=tk.NORMAL)
            messagebox.showerror('安装失败', msg2)
            return
        self.banner.pack_forget()
        self._ask_restart()

    def _ask_restart(self):
        win = tk.Toplevel(self.root)
        win.title('更新完成')
        win.geometry('360x150')
        win.transient(self.root)
        tk.Label(win, text='新版本已安装完成！', font='AppBold12').pack(pady=(16, 6))
        tk.Label(win, text='是否立即重启以使用新版本？', font='AppBase').pack(pady=4)
        bf = tk.Frame(win)
        bf.pack(pady=10)
        ttk.Button(bf, text='立即重启', style='Accent.TButton',
                   command=lambda: (win.destroy(), restart_app(self.root))).pack(side='left', padx=8)
        ttk.Button(bf, text='稍后重启', command=win.destroy).pack(side='left', padx=8)

    def _on_skip(self):
        v = getattr(self, 'banner_version', None)
        if v:
            cfg = load_update_config()
            cfg['skip_version'] = v
            save_update_config(cfg)
        self.banner.pack_forget()

    def _on_later(self):
        self.banner.pack_forget()

    def _restore_default_source(self):
        cfg = load_update_config()
        cfg['source'] = None
        save_update_config(cfg)
        self.restore_src_btn.pack_forget()
        self.update_status.config(text='已恢复默认更新地址', fg='#aab7c4')
        self.log('已恢复默认更新地址。')


# ============================================================
# 入口
# ============================================================
def _selftest():
    """核心逻辑自测，不弹界面。"""
    import tempfile
    print('== selftest ==')
    # 1) folder
    d = tempfile.mkdtemp(prefix='uitest_')
    for n in ['a.png', 'b.jpg', 'c.png', 'd.png']:
        Image.new('RGB', (300, 500), (120, 160, 200)).save(os.path.join(d, n))
    imgs = list_images(d)
    out1 = os.path.join(d, '图片导出.docx')
    build_docx([('', imgs)], out1, upscale=1.0, show_remark=False)
    print(f'folder: {len(imgs)} 张 -> {out1}')

    # 2) excel
    xlsx = r'G:/待删除/测试.xlsx'
    if os.path.isfile(xlsx):
        rows = load_image_rows(xlsx)
        out2 = r'G:/待删除/_selftest_导出.docx'
        build_docx(rows, out2, upscale=3.0, show_remark=False)
        print(f'excel: {len(rows)} 行 -> {out2}')
    else:
        print('(跳过 excel 自测：未找到测试.xlsx)')

    # 校验 folder 产物
    import zipfile as zf
    z = zf.ZipFile(out1)
    doc = z.read('word/document.xml').decode('utf-8')
    exts = re.findall(r'<a:ext cx="(\d+)" cy="(\d+)"', doc)
    print('folder 图片尺寸固定 7x11cm:',
          all(c == '2520000' and y == '3960000' for c, y in exts), f'({len(exts)} 张)')

    # 3) docx 批量替换
    from docx import Document as _D
    repdir = tempfile.mkdtemp(prefix='reptest_')
    src = os.path.join(repdir, 'a.docx')
    _doc = _D()
    _doc.add_paragraph('每周处理故障80个，表现良好。')
    _doc.add_paragraph('另一个段落也有 故障80个 字样。')
    tbl = _doc.add_table(rows=1, cols=1)
    tbl.rows[0].cells[0].text = '表内：每周处理故障80个需要关注'
    _doc.save(src)
    ok = DocxReplaceTool._process_file(
        src, '每周处理故障80个', '每周处理故障100个', None, False, False,
        True, True, True)
    rd = _D(src)
    para_txt = '\n'.join(p.text for p in rd.paragraphs)
    cell_txt = rd.tables[0].rows[0].cells[0].text
    print('docx 替换：段落命中 =', ok,
          '| 段落含新值 =', '每周处理故障100个' in para_txt,
          '| 表格命中 =', '故障100个' in cell_txt)
    # 正则
    src2 = os.path.join(repdir, 'b.docx')
    _doc2 = _D(); _doc2.add_paragraph('故障12个，故障3个。'); _doc2.save(src2)
    DocxReplaceTool._process_file(src2, r'故障(\d+)个', r'故障[\1]个', None, True, False, True, True, True)
    rt = '\n'.join(p.text for p in _D(src2).paragraphs)
    print('docx 正则替换：', rt)

    # 4) 版本更新模块逻辑自测（不联网：用本地 http 服务模拟清单与 exe）
    import http.server, socketserver, threading as _th
    print('\n-- 版本更新模块自测 --')
    assert parse_version('1.10') > parse_version('1.9'), '版本比较错误'
    assert parse_version('v1.4.0') == parse_version('1.4.0')
    assert cmp_version('1.4.0', '1.3.9') > 0
    print('版本比较：OK')

    # 造一个假 exe（MZ 头 + 内容），计算 sha256/size
    fake = b'MZ' + bytes(range(256)) * 40
    sha = hashlib.sha256(fake).hexdigest()
    sz = len(fake)
    Handler = http.server.SimpleHTTPRequestHandler

    class _Srv(socketserver.TCPServer):
        allow_reuse_address = True

    # 用临时目录起服务（先起服务拿到端口，再写清单）
    srv_dir = tempfile.mkdtemp(prefix='updsrv_')
    with open(os.path.join(srv_dir, 'upd.exe'), 'wb') as f:
        f.write(fake)
    os.chdir(srv_dir)
    import functools
    h = functools.partial(Handler)
    srv = _Srv(('127.0.0.1', 0), h)
    port = srv.server_address[1]
    _th.Thread(target=srv.serve_forever, daemon=True).start()
    manifest = {'version': '9.9.9', 'download_url': 'http://127.0.0.1:%d/upd.exe' % port,
                'notes': '测试', 'sha256': sha, 'size': sz, 'pub_date': '2026-11-14',
                'mandatory': False}
    with open(os.path.join(srv_dir, 'version.json'), 'w', encoding='utf-8') as f:
        json.dump(manifest, f)
    try:
        murl = 'http://127.0.0.1:%d/version.json' % port
        d, hit = _fetch_manifest([murl])
        assert d and d.get('version') == '9.9.9', '清单获取失败'
        print('清单获取：OK (命中 %s)' % hit)
        dest = os.path.join(srv_dir, '_dl.exe')
        ok, msg = download_update(d['download_url'], dest, sha256=sha, size=sz)
        assert ok, '下载校验失败: ' + msg
        with open(dest, 'rb') as f:
            assert f.read(2) == b'MZ'
        print('下载+校验(MZ/sha256/size)：OK')
    finally:
        srv.shutdown()
        os.chdir(os.path.dirname(os.path.abspath(__file__)))

    print('selftest OK')


if __name__ == '__main__':
    if '--selftest' in sys.argv:
        _selftest()
    elif '--tkprobe' in sys.argv:
        # 打包后验证 Tcl/Tk 能否真正初始化（selftest 不建 Tk，测不出 init.tcl 问题）
        try:
            _r = tk.Tk()
            _r.withdraw()
            print('TK OK', _r.tk.eval('info patchlevel'),
                  '| tcl_library =', _r.tk.eval('info library'))
            _r.destroy()
            sys.exit(0)
        except Exception:
            traceback.print_exc()
            sys.exit(1)
    else:
        App().root.mainloop()
