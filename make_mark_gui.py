# -*- coding: utf-8 -*-
"""唛头生成器 · 图形界面版

给不装 Python 的同事用：双击 exe → 把「唛头模板表.xlsx」拖进窗口 → 点生成。
输出**固定写到模板表所在的文件夹**，和他们的三单/四单放在一起。

设计原则（都是被同事的机器逼出来的）
1. **宽容**：能忽略的绝不报错（多余的列、空行、票名不连续、PL 对不上 → 提示/询问）。
   只有「不修就没法生成」的问题才拦下来。
2. **说清楚**：每条问题都带**行号 + 单元格**，并给出**具体怎么改**（advice）。
3. **不装环境**：单文件 exe，配置/母本都有打包兜底，旁边没有也能跑。

命令行（给自检用，不影响 GUI）：
    make_mark_gui.py                     打开图形界面
    make_mark_gui.py --check 模板表.xlsx  不开窗口，只做检查并打印问题（测试用）
    make_mark_gui.py --selftest-gui      建一遍全部界面与图示后退出（回归用）
"""

import os
import re
import sys
import json
import traceback

# 同目录的核心逻辑（打包时一并收进 exe）
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import make_mark as core          # noqa: E402

# ★ tkinter 用「软导入」：`--check` 是纯检查、不开窗口，没有 tkinter 的机器也要能跑
#   （测试机上的精简 Python 就没有 tkinter）。真正建窗口时再要求它存在。
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    _HAS_TK = True
except Exception:                  # pragma: no cover
    tk = ttk = filedialog = messagebox = None
    _HAS_TK = False

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _HAS_DND = True
except Exception:                  # 拖拽是加分项，装不上就退化成「点按钮选文件」
    TkinterDnD, DND_FILES, _HAS_DND = None, None, False
if not _HAS_TK:
    _HAS_DND = False

APP_TITLE = "唛头生成器"

# ---------------------------------------------------------------- 外观

C_BG      = "#f2f4f7"   # 窗口底
C_CARD    = "#ffffff"   # 卡片
C_BORDER  = "#d5dae2"
C_TEXT    = "#1f2329"
C_MUTED   = "#6b7280"
C_ACCENT  = "#1668dc"
C_ACCENT_D= "#0d4fb0"
C_OK      = "#12813f"
C_WARN    = "#a8620a"
C_ERR     = "#c22b2b"
C_DROP_BG = "#f7faff"
C_DROP_HI = "#e8f1ff"

FONT_UI   = ("Microsoft YaHei UI", 10)
FONT_UI_B = ("Microsoft YaHei UI", 10, "bold")
FONT_H1   = ("Microsoft YaHei UI", 15, "bold")
FONT_H2   = ("Microsoft YaHei UI", 11, "bold")
FONT_MONO = ("Consolas", 9)
FONT_BIG  = ("Microsoft YaHei UI", 12, "bold")
FONT_HINT = ("Microsoft YaHei UI", 9)


def _pick_fonts(root):
    """挑一个本机有的中文字体，避免在某些机器上掉成方框。"""
    global FONT_UI, FONT_UI_B, FONT_H1, FONT_H2, FONT_BIG, FONT_HINT, FONT_MONO
    try:
        fams = set(root.tk.call("font", "families"))
    except Exception:
        return
    for cand in ("Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", "SimHei", "Segoe UI"):
        if cand in fams:
            FONT_UI = (cand, 10)
            FONT_UI_B = (cand, 10, "bold")
            FONT_H1 = (cand, 15, "bold")
            FONT_H2 = (cand, 11, "bold")
            FONT_BIG = (cand, 12, "bold")
            FONT_HINT = (cand, 9)
            break
    for cand in ("Consolas", "Cascadia Mono", "Courier New"):
        if cand in fams:
            FONT_MONO = (cand, 9)
            break


# ---------------------------------------------------------------- 小工具

def round_rect(cv, x1, y1, x2, y2, r=8, **kw):
    """Canvas 圆角矩形（用平滑多边形近似）。"""
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
           x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return cv.create_polygon(pts, smooth=True, **kw)


def arrow(cv, x1, y1, x2, y2, color=C_MUTED, w=2, head=7):
    """带箭头的直线。"""
    cv.create_line(x1, y1, x2, y2, fill=color, width=w)
    import math
    ang = math.atan2(y2 - y1, x2 - x1)
    a1 = (x2 - head * math.cos(ang - 0.42), y2 - head * math.sin(ang - 0.42))
    a2 = (x2 - head * math.cos(ang + 0.42), y2 - head * math.sin(ang + 0.42))
    cv.create_polygon(x2, y2, a1[0], a1[1], a2[0], a2[1], fill=color, outline=color)


def open_in_explorer(path, select=False):
    """在资源管理器里打开文件夹 / 选中文件。"""
    try:
        if select and os.path.exists(path):
            os.system('explorer /select,"%s"' % os.path.normpath(path))
        else:
            os.startfile(os.path.dirname(os.path.abspath(path)) if os.path.isfile(path) else path)
        return True
    except Exception:
        return False


def open_file(path):
    try:
        os.startfile(path)
        return True
    except Exception as e:
        messagebox.showwarning(APP_TITLE, "打不开这个文件：\n%s\n\n%s" % (path, e))
        return False


XLSX_EXT = (".xlsx", ".xlsm", ".xltx", ".xltm")


def is_template_file(p):
    return bool(p) and os.path.isfile(p) and p.lower().endswith(XLSX_EXT)


def output_path_for(template_path):
    """★ 输出规则：和模板表同一个文件夹，名字 = 唛头 + 合同号.docx

    合同号从「模板表里的票名」推不出来（要读表），所以这里先用文件名兜底，
    真正生成时以表里的第一个票名为准（见 derive_output）。
    """
    d = os.path.dirname(os.path.abspath(template_path))
    stem = os.path.splitext(os.path.basename(template_path))[0]
    return os.path.join(d, "唛头%s.docx" % contract_stem(stem))


def contract_stem(name):
    """DEMO-20260101-1A → DEMO-20260101-1（去掉末尾的分票字母）。"""
    return re.sub(r"(?<=[0-9])[A-Za-z]+$", "", name) or name


def derive_output(template_path, first_ticket):
    d = os.path.dirname(os.path.abspath(template_path))
    stem = contract_stem(first_ticket) if first_ticket else \
        contract_stem(os.path.splitext(os.path.basename(template_path))[0])
    return os.path.join(d, "唛头%s.docx" % stem)


# ---------------------------------------------------------------- 生成流程（GUI 与命令行共用）

class GenResult(object):
    """一次「检查 / 生成」的结果。"""

    def __init__(self):
        self.issues = core.IssueSink()   # 结构化问题
        self.log = []                    # 给日志区用的文本行 (level, text)
        self.groups = None
        self.master = None               # load_master 的返回值
        self.singular = True
        self.out_path = None
        self.n_pages = 0
        self.written = False
        self.fatal = None                # 致命错误文本

    # -- 便捷判断
    @property
    def errors(self):
        return self.issues.of("error")

    @property
    def confirms(self):
        return self.issues.of("confirm")

    @property
    def warns(self):
        return self.issues.of("warn")

    @property
    def needs_confirm(self):
        """要在生成前问用户一声的问题：PL 对不上、以及「内容太长会挤到第二页」。"""
        return self.confirms + [w for w in self.warns if w.get("code") == "overflow"]


class _Tee(object):
    """把 print/stderr 同时收进日志列表（GUI 用来显示解析明细）。"""

    def __init__(self, sink, level="info"):
        self.sink, self.level, self.buf = sink, level, ""

    def write(self, s):
        self.buf += s
        while "\n" in self.buf:
            line, self.buf = self.buf.split("\n", 1)
            if line.strip():
                self.sink.append((self.level, line.rstrip()))

    def flush(self):
        if self.buf.strip():
            self.sink.append((self.level, self.buf.rstrip()))
        self.buf = ""


def load_config(exe_dir=None):
    """读配置：优先 exe 旁边那份（用户改得动），没有就用打包进 exe 的默认份。"""
    exe_dir = exe_dir or core.base_dir()
    path = os.path.join(exe_dir, "mark_config.json")
    if not os.path.exists(path):
        path = core.find_asset("mark_config.json", exe_dir) or path
    if not os.path.exists(path):
        raise IOError("找不到配置文件 mark_config.json。\n"
                      "请把「唛头生成器」整个文件夹一起复制，不要只拷 exe。")
    with open(path, encoding="utf-8") as f:
        cfg = json.load(f)
    # 相对路径一律相对 exe 所在目录（母本就放在 exe 旁边）
    cfg["_config_dir"] = exe_dir
    return cfg


def make_effective_cfg(cfg, template_path):
    """把配置改成「这一次生成」要用的样子。

    ★ 三处覆盖，全是 GUI 的硬要求：
      ① 输入 = 用户拖进来的那张表
      ② 输出 = **模板表所在文件夹**（不是 exe 所在文件夹）
      ③ PL 核对目录 = 模板表所在文件夹（三单/四单就和模板表放一起）
    """
    c = json.loads(json.dumps(cfg, ensure_ascii=False))     # 深拷贝，别污染原配置
    tpl_dir = os.path.dirname(os.path.abspath(template_path))
    c["input_xlsx"] = os.path.abspath(template_path)
    c["output_docx"] = ""                                   # 留空 → 由下面算出绝对路径
    pc = c.setdefault("pl_check", {})
    pc["dir"] = [tpl_dir]
    return c


def find_master(cfg):
    name = cfg.get("master_docx", "唛头母本.docx")
    p = os.path.join(cfg["_config_dir"], name)
    if os.path.exists(p):
        return p
    return core.find_asset(os.path.basename(name), cfg["_config_dir"])


def run_check(template_path, cfg=None, log=None):
    """阶段一：解析 + 校验 + PL 核对 + 溢出预估。不写任何文件。"""
    res = GenResult()
    log = res.log if log is None else log
    cfg = cfg or load_config()
    eff = make_effective_cfg(cfg, template_path)

    master = find_master(cfg)
    if not master:
        res.fatal = ("找不到「唛头母本.docx」。\n"
                     "它必须和本程序放在同一个文件夹里 —— 请把整个「唛头生成器」文件夹一起复制。")
        return res
    log.append(("info", "模板表：%s" % os.path.abspath(template_path)))
    log.append(("info", "母本  ：%s" % master))

    old_out, old_err = sys.stdout, sys.stderr
    tee = _Tee(log)
    sys.stdout = sys.stderr = tee
    try:
        groups, _ = core.parse_template(eff["input_xlsx"], eff)
        # ★ 先把输出路径算出来（哪怕后面校验不过）—— 界面上要能告诉用户「会写到哪」
        res.out_path = derive_output(template_path,
                                     groups[0]["name"] if groups else None)
        core.validate(groups, eff, sink=res.issues)
        if res.errors:                       # 必填项都没齐，后面的检查没意义
            return res
        core.check_pl(groups, eff, sink=res.issues)
        singular = bool(eff.get("total_singular", True))
        res.master = core.load_master(master)
        labels, geom = res.master[6], res.master[7]
        core.check_overflow(groups, labels, geom, eff, singular, sink=res.issues)
        res.groups, res.singular = groups, singular
        res.out_path = derive_output(template_path, groups[0]["name"])
        total = sum(g["total_pkg"] for g in groups)
        log.append(("h", "解析结果：%d 票 / %d 行明细 → 共 %d 个唛头（%d 页）"
                    % (len(groups), sum(len(g["rows"]) for g in groups), total, total)))
        for g in groups:
            log.append(("info", "  ● %s   件数合计 %d" % (g["name"], g["total_pkg"])))
            idx = 1
            for it in g["rows"]:
                v = it["values"]
                log.append(("dim", "      第 %-3d 行  件数 %-3d  → 编号 %d-%d … %d-%d   %s / %s"
                            % (it["row"], it["n_pkg"], g["total_pkg"], idx,
                               g["total_pkg"], idx + it["n_pkg"] - 1,
                               v.get("品名英文", ""), v.get("厂家英文", ""))))
                idx += it["n_pkg"]
        log.append(("info", "输出将写到：%s" % res.out_path))
    except SystemExit as e:
        res.fatal = "检查没通过（退出码 %s）。" % e.code
    except Exception as e:
        res.fatal = "读取模板表出错：%s\n%s" % (e, traceback.format_exc(limit=3))
    finally:
        sys.stdout, sys.stderr = old_out, old_err
        tee.flush()
    return res


def run_build(res, template_path):
    """阶段二：真正写 docx。用阶段一的结果，不再重复解析。"""
    if not res.groups or not res.master:
        res.fatal = res.fatal or "还没有通过检查，无法生成。"
        return res
    out = res.out_path
    try:
        os.makedirs(os.path.dirname(out), exist_ok=True)
    except Exception:
        pass
    if os.path.exists(out):
        try:
            os.remove(out)
        except Exception as e:
            res.fatal = ("目标文件正被 Word 打开，写不进去：\n%s\n\n"
                         "请先关掉 Word 里打开的这个文件，再点一次生成。" % out)
            return res
    mz, head, block, sep, tail, sectpr, labels = res.master[:7]
    try:
        n = core.build_docx(mz, head, block, sep, tail, sectpr, labels,
                            res.groups, out, singular=res.singular)
    except Exception as e:
        res.fatal = "写文件失败：%s\n%s" % (e, traceback.format_exc(limit=3))
        return res
    res.written, res.n_pages = True, n
    res.log.append(("ok", "已生成：%s" % out))
    res.log.append(("ok", "  %d 个唛头 / %d 页（一件一个、块间自动分页）" % (n, n)))
    return res


# ---------------------------------------------------------------- 界面

BaseTk = object
if _HAS_TK:
    BaseTk = TkinterDnD.Tk if _HAS_DND else tk.Tk


class App(BaseTk):

    def __init__(self):
        super().__init__()
        _pick_fonts(self)
        self.title("%s v%s" % (APP_TITLE, core.APP_VERSION))
        self.configure(bg=C_BG)
        self.geometry("980x690")
        self.minsize(860, 560)
        try:
            self.iconbitmap(default="")
        except Exception:
            pass

        self.template = None        # 当前模板表绝对路径
        self.last_out = None        # 最近一次产出
        self.res = None             # 最近一次检查结果
        self.busy = False

        self._build_style()
        self._build_header()
        self._build_tabs()
        self._build_statusbar()

        if _HAS_DND:
            try:
                self.drop_target_register(DND_FILES)
                self.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass
        self.after(120, self._greet)

    # ------------------------------------------------------------ 样式

    def _build_style(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except Exception:
            pass
        st.configure("TNotebook", background=C_BG, borderwidth=0)
        st.configure("TNotebook.Tab", font=FONT_UI_B, padding=(18, 9),
                     background="#e4e8ee", foreground=C_MUTED, borderwidth=0)
        st.map("TNotebook.Tab",
               background=[("selected", C_CARD)],
               foreground=[("selected", C_ACCENT)])
        st.configure("TFrame", background=C_BG)
        st.configure("Card.TFrame", background=C_CARD)
        st.configure("TLabel", background=C_BG, foreground=C_TEXT, font=FONT_UI)
        st.configure("Card.TLabel", background=C_CARD, foreground=C_TEXT, font=FONT_UI)
        st.configure("H2.TLabel", background=C_CARD, foreground=C_TEXT, font=FONT_H2)
        st.configure("Hint.TLabel", background=C_CARD, foreground=C_MUTED, font=FONT_HINT)
        st.configure("TButton", font=FONT_UI, padding=(12, 6))
        st.configure("Big.TButton", font=FONT_BIG, padding=(26, 10),
                     background=C_ACCENT, foreground="#ffffff", borderwidth=0)
        st.map("Big.TButton",
               background=[("disabled", "#b9c4d2"), ("pressed", C_ACCENT_D),
                           ("active", C_ACCENT_D)],
               foreground=[("disabled", "#ffffff")])
        st.configure("Link.TButton", font=FONT_HINT, padding=(4, 2),
                     background=C_CARD, foreground=C_ACCENT, borderwidth=0,
                     relief="flat")
        st.map("Link.TButton", background=[("active", "#eaf2ff")],
               foreground=[("active", C_ACCENT_D)])
        st.configure("Vertical.TScrollbar", background="#c8cfd8")

    # ------------------------------------------------------------ 顶部

    def _build_header(self):
        bar = tk.Frame(self, bg=C_ACCENT, height=56)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text="唛头生成器", bg=C_ACCENT, fg="#ffffff",
                 font=("Microsoft YaHei UI", 15, "bold")).pack(side="left", padx=(18, 12))
        tk.Label(bar, text="把 Excel 模板表变成 Word 唛头 · 一件一张 · 自动分页",
                 bg=C_ACCENT, fg="#d6e6ff", font=FONT_UI).pack(side="left", pady=(3, 0))
        tk.Label(bar, text="v%s" % core.APP_VERSION, bg=C_ACCENT, fg="#d6e6ff",
                 font=FONT_HINT).pack(side="right", padx=16)

    # ------------------------------------------------------------ 主体

    def _build_tabs(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=14, pady=(12, 6))
        self.tab_gen = ttk.Frame(self.nb, style="TFrame")
        self.tab_tut = ttk.Frame(self.nb, style="TFrame")
        self.tab_faq = ttk.Frame(self.nb, style="TFrame")
        self.nb.add(self.tab_gen, text="  开始生成  ")
        self.nb.add(self.tab_tut, text="  使用教程  ")
        self.nb.add(self.tab_faq, text="  常见问题  ")
        self._build_generate_tab()
        self._build_tutorial_tab()
        self._build_faq_tab()

    def _card(self, parent, pady=(0, 8)):
        outer = tk.Frame(parent, bg=C_BORDER)
        outer.pack(fill="x", pady=pady)
        inner = tk.Frame(outer, bg=C_CARD)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        return inner

    def _step_head(self, card, n, text):
        head = tk.Frame(card, bg=C_CARD)
        head.pack(fill="x", padx=14, pady=(10, 4))
        tk.Label(head, text="第 %d 步" % n, bg=C_ACCENT, fg="#ffffff", font=FONT_HINT,
                 padx=8, pady=2).pack(side="left")
        tk.Label(head, text="  " + text, bg=C_CARD, fg=C_TEXT,
                 font=FONT_H2).pack(side="left")
        return head

    # ------------------------------------------------------------ 生成页

    def _build_generate_tab(self):
        root = self.tab_gen

        # ---- 第 1 步：拖入
        card = self._card(root)
        self._step_head(card, 1, "把「唛头模板表.xlsx」拖进下面的框里")

        self.drop = tk.Canvas(card, height=104, bg=C_CARD, highlightthickness=0)
        self.drop.pack(fill="x", padx=14, pady=(2, 6))
        self.drop.bind("<Configure>", lambda e: self._draw_drop())
        self.drop.bind("<Button-1>", lambda e: self.choose_file())
        self.drop.bind("<Enter>", lambda e: self._drop_hover(True))
        self.drop.bind("<Leave>", lambda e: self._drop_hover(False))
        self.drop.configure(cursor="hand2")
        self._drop_state = "idle"

        row = tk.Frame(card, bg=C_CARD)
        row.pack(fill="x", padx=14, pady=(0, 10))
        self.btn_choose = ttk.Button(row, text="选择模板表…", command=self.choose_file)
        self.btn_choose.pack(side="left")
        self.btn_open_tpl = ttk.Button(row, text="打开模板表", command=self.open_template,
                                       state="disabled")
        self.btn_open_tpl.pack(side="left", padx=6)
        self.btn_open_dir = ttk.Button(row, text="打开所在文件夹", command=self.open_template_dir,
                                       state="disabled")
        self.btn_open_dir.pack(side="left")
        # ⚠ 这里原来用的是「链接样式」的 tk.Label —— 实测在部分机器上根本不绘制
        #   （winfo_ismapped=1 但屏幕上一个像素都没有）。统一用 ttk.Button，稳。
        self.btn_sample = ttk.Button(row, text="载入示例模板", command=self.use_sample)
        self.btn_sample.pack(side="right")

        # ---- 第 2 步：输出位置
        card2 = self._card(root)
        self._step_head(card2, 2, "确认输出位置")
        self.lbl_out = tk.Label(card2, text="（拖入模板表后这里会显示输出到哪）",
                                bg=C_CARD, fg=C_MUTED, font=FONT_UI,
                                anchor="w", justify="left", wraplength=880)
        self.lbl_out.pack(fill="x", padx=14, pady=(0, 10))

        # ---- 第 3 步：生成
        card3 = self._card(root, pady=(0, 6))
        self._step_head(card3, 3, "先检查、再生成")

        btns = tk.Frame(card3, bg=C_CARD)
        btns.pack(fill="x", padx=14, pady=(2, 4))
        self.btn_check = ttk.Button(btns, text="先检查不生成", command=lambda: self.do_work("check"))
        self.btn_check.pack(side="left")
        self.btn_go = ttk.Button(btns, text="开始生成", style="Big.TButton",
                                 command=lambda: self.do_work("build"))
        self.btn_go.pack(side="left", padx=10)
        self.btn_res = ttk.Button(btns, text="打开生成的文件", command=self.open_result,
                                  state="disabled")
        self.btn_res.pack(side="left", padx=6)
        self.btn_resdir = ttk.Button(btns, text="打开所在文件夹", command=self.open_result_dir,
                                     state="disabled")
        self.btn_resdir.pack(side="left")
        self.btn_clear = ttk.Button(btns, text="清空信息", command=self.clear_log)
        self.btn_clear.pack(side="right")

        # ---- 日志
        logwrap = tk.Frame(card3, bg=C_CARD)
        logwrap.pack(fill="both", expand=True, padx=14, pady=(6, 10))
        self.log = tk.Text(logwrap, height=5, wrap="word", font=FONT_UI,
                           bg="#fbfcfd", fg=C_TEXT, relief="flat",
                           highlightthickness=1, highlightbackground=C_BORDER,
                           padx=10, pady=6, state="disabled")
        sb = ttk.Scrollbar(logwrap, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log.pack(side="left", fill="both", expand=True)
        self.log.tag_configure("h", font=FONT_UI_B, foreground=C_TEXT,
                               spacing1=6, spacing3=3)
        self.log.tag_configure("ok", foreground=C_OK)
        self.log.tag_configure("err", foreground=C_ERR)
        self.log.tag_configure("warn", foreground=C_WARN)
        self.log.tag_configure("confirm", foreground=C_WARN)
        self.log.tag_configure("info", foreground=C_TEXT)
        self.log.tag_configure("dim", foreground=C_MUTED, font=FONT_MONO)
        self.log.tag_configure("path", foreground=C_ACCENT)
        self.log.tag_configure("advice", foreground="#3a4250", lmargin1=22, lmargin2=22)

    # ---- 拖拽框绘制

    def _draw_drop(self):
        cv = self.drop
        cv.delete("all")
        w = cv.winfo_width() or 900
        h = 104
        hover = self._drop_state == "hover"
        loaded = self._drop_state == "loaded"
        bg = C_DROP_HI if hover else C_DROP_BG
        border = C_ACCENT if (hover or loaded) else "#b9c4d2"
        round_rect(cv, 3, 3, w - 4, h - 4, 12, fill=bg, outline=border, width=2)

        if loaded and self.template:
            name = os.path.basename(self.template)
            cv.create_text(w / 2, 28, text="已选择模板表", fill=C_OK,
                           font=FONT_UI_B, anchor="center")
            cv.create_text(w / 2, 54, text=name, fill=C_TEXT,
                           font=("Microsoft YaHei UI", 13, "bold"), anchor="center")
            cv.create_text(w / 2, 80, text="要换一个？把新文件拖进来，或点这里重新选择",
                           fill=C_MUTED, font=FONT_HINT, anchor="center")
            return

        # 文档图标
        cx, cy = w / 2, 26
        cv.create_polygon(cx - 12, cy - 15, cx + 4, cy - 15, cx + 12, cy - 7,
                          cx + 12, cy + 15, cx - 12, cy + 15,
                          fill="#ffffff", outline=C_ACCENT, width=2)
        cv.create_line(cx + 4, cy - 15, cx + 4, cy - 7, cx + 12, cy - 7,
                       fill=C_ACCENT, width=2)
        for i in range(3):
            cv.create_line(cx - 6, cy - 1 + i * 6, cx + 6, cy - 1 + i * 6,
                           fill="#9fb6d6", width=1)

        tip = "把「唛头模板表.xlsx」拖到这里" if _HAS_DND else "点这里选择「唛头模板表.xlsx」"
        cv.create_text(w / 2, 58, text=tip, fill=C_TEXT,
                       font=("Microsoft YaHei UI", 13, "bold"), anchor="center")
        cv.create_text(w / 2, 82,
                       text="也可以点「选择模板表…」按钮　·　支持 .xlsx / .xlsm",
                       fill=C_MUTED, font=FONT_HINT, anchor="center")

    def _drop_hover(self, on):
        if self._drop_state == "loaded":
            return
        self._drop_state = "hover" if on else "idle"
        self._draw_drop()

    # ---- 拖放事件

    def _on_drop(self, event):
        try:
            items = self.tk.splitlist(event.data)
        except Exception:
            items = [event.data]
        files = [p for p in items if is_template_file(p)]
        if not files:
            bad = [p for p in items if os.path.isfile(p)]
            if bad:
                self._log("err", "这个文件不是 Excel 表格：%s" % os.path.basename(bad[0]))
                self._log("advice", "请拖「唛头模板表.xlsx」这类 .xlsx 文件。"
                                    "如果你拖的是 .xls（老格式），先用 Excel 另存为 .xlsx。")
            else:
                self._log("err", "没识别到文件。请直接从资源管理器里拖 Excel 文件进来。")
            return
        if len(files) > 1:
            self._log("warn", "一次只能处理一个模板表，已用第一个：%s" % os.path.basename(files[0]))
        self.set_template(files[0])

    # ---- 选择文件

    def choose_file(self):
        init = os.path.dirname(self.template) if self.template else os.path.expanduser("~")
        p = filedialog.askopenfilename(
            title="选择唛头模板表", initialdir=init,
            filetypes=[("Excel 模板表", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if p:
            self.set_template(p)

    def use_sample(self):
        p = core.find_asset("唛头模板表.xlsx", core.base_dir())
        if not p:
            self._log("err", "旁边没有「唛头模板表.xlsx」，也没有打包进程序的示例。")
            return
        self._log("info", "载入示例模板表（输出会写到程序所在文件夹）。")
        self.set_template(p)

    def set_template(self, path):
        self.template = os.path.abspath(path)
        self._drop_state = "loaded"
        self._draw_drop()
        self.btn_open_tpl.configure(state="normal")
        self.btn_open_dir.configure(state="normal")
        self.btn_res.configure(state="disabled")
        self.btn_resdir.configure(state="disabled")
        self.last_out = None
        self.lbl_out.configure(
            text="模板表：%s\n输出将写到同一个文件夹里（程序会先读表里的票名，再定文件名）"
                 % self.template, fg=C_TEXT)
        self._log("h", "已选择：%s" % os.path.basename(self.template))
        self._log("info", "位置：%s" % os.path.dirname(self.template))
        self._set_status("模板表：%s" % os.path.basename(self.template))
        # ★ 拖进来就自动检查一遍，别等用户点了才发现问题
        self.do_work("check", silent_ok=True)

    # ------------------------------------------------------------ 教程页

    def _build_tutorial_tab(self):
        outer = ttk.Frame(self.tab_tut)
        outer.pack(fill="both", expand=True, padx=2, pady=2)
        cv = tk.Canvas(outer, bg=C_BG, highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=cv.yview)
        inner = tk.Frame(cv, bg=C_BG)
        inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        win = cv.create_window((0, 0), window=inner, anchor="nw")
        cv.bind("<Configure>", lambda e: cv.itemconfigure(win, width=e.width))
        cv.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        cv.pack(side="left", fill="both", expand=True)

        def wheel(e):
            cv.yview_scroll(int(-1 * (e.delta / 120)), "units")
        cv.bind_all("<MouseWheel>", lambda e: wheel(e) if self.nb.index("current") == 1 else None)

        P = dict(padx=18)

        def title(t):
            tk.Label(inner, text=t, bg=C_BG, fg=C_TEXT, font=FONT_H1,
                     anchor="w").pack(fill="x", pady=(16, 6), **P)

        def para(t, color=C_TEXT, font=None, lpad=18):
            tk.Label(inner, text=t, bg=C_BG, fg=color, font=font or FONT_UI,
                     anchor="w", justify="left", wraplength=880).pack(fill="x", padx=lpad, pady=2)

        def bullet(t, n="•"):
            tk.Label(inner, text="%s  %s" % (n, t), bg=C_BG, fg=C_TEXT, font=FONT_UI,
                     anchor="w", justify="left", wraplength=860).pack(fill="x", padx=34, pady=1)

        def diagram(draw, height):
            f = tk.Frame(inner, bg=C_CARD, highlightthickness=1,
                         highlightbackground=C_BORDER)
            f.pack(fill="x", padx=18, pady=(8, 14))
            c = tk.Canvas(f, height=height, bg=C_CARD, highlightthickness=0)
            c.pack(fill="x")
            self.update_idletasks()
            c.bind("<Configure>", lambda e, cc=c, dd=draw: self._safe_draw(cc, dd))
            self._safe_draw(c, draw)
            return c

        title("这个程序做什么")
        para("把「唛头模板表」（Excel）里的内容，批量生成一份 Word 唛头。")
        para("一件货一个唛头、一页放一张，块与块之间自动分页 —— 你不用手动排版。", color=C_MUTED)
        diagram(self._draw_flow, 156)

        title("三步就完事")
        para("① 准备模板表：照「唛头模板表.xlsx」的样子填好，列可以从三单/四单的「PL箱单」直接复制。")
        para("② 拖进来：把那个 .xlsx 文件拖到「开始生成」页的大框里（或点「选择模板表…」）。")
        para("③ 点「开始生成」。程序会先自己检查一遍，没问题就出 Word。")
        para("检查不通过时它会停下，并把「哪一行、哪个单元格、要怎么改」一条条列在下面，"
             "照着改完再点一次就行。", color=C_MUTED)

        title("模板表怎么填")
        para("列的排法和「PL箱单」对齐，左边一大段直接整列复制粘贴，右边几列手填。")
        diagram(self._draw_table, 228)
        bullet("A 到 M 列：从「PL箱单」整列复制过来。程序只用得到 Descriptions（品名）和 "
               "Packages（件数），其余列留着方便你复制。")
        bullet("N 到 T 列：PL 里没有的，要手填 —— 票名 / 厂家英文 / 发货人 / 发货人地址 / "
               "收货人 / 收货人地址 / 原产国。")
        bullet("「票名」只写每票的第一行就行，下面留空会自动继承（同一票写一次）。")
        bullet("「件数」= 这一行要出几个唛头。比如这一行有 2 件货，就填 2，程序会出 2 张。")
        bullet("编号是自动的：PACKAGE NO = 本票件数合计-序号（如 13-1 … 13-13），"
               "本票所有唛头的 TOTAL 都写本票件数合计。")

        title("生成的文件放哪")
        para("固定写在「模板表所在的文件夹」里，和你的三单/四单在一起，不用你选路径。")
        diagram(self._draw_output, 186)

        title("程序会帮你检查什么")
        bullet("必填项有没有空着（会告诉你第几行、哪个单元格）。")
        bullet("件数是不是 1 以上的整数。")
        bullet("件数合计和同一个文件夹里「中国三单 / 越南四单」PL箱单 的 Packages 对不对得上。")
        bullet("唛头内容会不会太长、一页放不下（放不下只提醒，不拦你）。")
        para("前两条不通过就没法生成，必须改；后两条只提示，你确认就行。", color=C_MUTED)

        title("内容太长、一页放不下会怎样")
        para("一个唛头固定 11 行、18pt 字，母本页边距下大约能放 24 行 —— 也就是每行文字"
             "「折行」之后总共不能超过 24 行。", color=C_MUTED)
        para("程序没有 Word 的排版引擎，只能按字体宽度「估」：估出来超了，就会告诉你"
             "「哪一票、哪个唛头、超几行、哪一段最长、大概删几个字符就能收回来」。")
        para("这时你有两个选择：", color=C_MUTED)
        bullet("按它点名的段落精简（最快）—— 通常是收货人地址太长，或者地址里塞了多余的空格、换行。")
        bullet("不管它，点「继续生成」—— 那个唛头会被拆到两页上；因为它后面跟着硬分页符，"
               "后面的唛头会整体错位、页码也跟着变。少量几张要手工贴的话可以这么干。")
        para("想要「超了就干脆不让生成」，把 mark_config.json 里 "
             "overflow_check.on_overflow 改成 \"error\" 即可。", color=C_MUTED)

        title("几个名词")
        bullet("票名：这一批货的编号，例如 DEMO-20260101-1A。同一票的唛头编号会连在一起。")
        bullet("件数 / Packages：这一行有几件货。", n="•")
        bullet("TOTAL：本票一共有几件（写在每个唛头的右上）。")
        bullet("母本：程序旁边那个「唛头母本.docx」，决定字体、字号、页边距。"
               "想改版式就改它，别删别改名。")

    # ------------------------------------------------------------ 常见问题页

    def _build_faq_tab(self):
        outer = ttk.Frame(self.tab_faq)
        outer.pack(fill="both", expand=True, padx=2, pady=2)
        cv = tk.Canvas(outer, bg=C_BG, highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=cv.yview)
        inner = tk.Frame(cv, bg=C_BG)
        inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        win = cv.create_window((0, 0), window=inner, anchor="nw")
        cv.bind("<Configure>", lambda e: cv.itemconfigure(win, width=e.width))
        cv.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        cv.pack(side="left", fill="both", expand=True)
        cv.bind_all("<MouseWheel>", lambda e: cv.yview_scroll(int(-1 * (e.delta / 120)), "units")
                    if self.nb.index("current") == 2 else None)

        tk.Label(inner, text="出问题了？先在这里找", bg=C_BG, fg=C_TEXT, font=FONT_H1,
                 anchor="w").pack(fill="x", padx=18, pady=(16, 4))
        tk.Label(inner, text="下面按「现象 → 怎么办」列。程序报错时也会直接把办法写在信息区。",
                 bg=C_BG, fg=C_MUTED, font=FONT_UI, anchor="w").pack(fill="x", padx=18, pady=(0, 12))

        FAQ = [
            ("程序说「找不到唛头母本.docx」",
             "你把 exe 单独拷出来了。母本必须和 exe 在一起 —— 请把整个「唛头生成器」文件夹一起复制过去。"),
            ("程序说某一行的「票名」为空",
             "在那一行的「票名」列填上这一票的名字（例如 DEMO-20260101-1A）。"
             "票名只写每票第一行即可，下面的行留空会自动继承。"),
            ("程序说某个单元格是公式、没有缓存值",
             "用 Excel 或 WPS 打开模板表，按一次 Ctrl+S 保存（让公式算出结果），再回来重新生成。"),
            ("程序说件数必须是 1 以上的整数",
             "找到它说的那个单元格，把里面的文字/小数改成整数，比如 2。"
             "如果那格是文本格式的「2」，一般也能认；认不出就重新输入一遍。"),
            ("提示「件数合计与 PL箱单 的 Packages 合计不一致」",
             "这只是顺手帮你核对。如果你知道确实该不一样（比如 PL 还没更新），"
             "在弹窗里选「继续生成」就行；否则先去把模板表或 PL 改一致。"),
            ("提示「单页溢出风险」/「一页放不下」",
             "某个唛头内容太长，折行后会超过一页（大约 24 行），Word 里会被拆到两页上；"
             "因为它后面跟着硬分页符，后面的唛头会整体错位。"
             "点「先检查不生成」看它点名的那一段（通常是收货人地址），"
             "按提示删掉几个字符就能收回来。不想改就点「继续生成」，程序照出不拦你。"),
            ("想让程序对「放不下」直接拦下来",
             "打开程序旁边的 mark_config.json，把 overflow_check 里的 "
             "on_overflow 从 \"warn\" 改成 \"error\"，保存即可。"
             "改完程序会在有唛头放不下时停下、不产出。"),
            ("提示「没找到任何 PL 箱单，跳过件数核对」",
             "不影响生成。想让程序帮你核对件数，就把模板表和三单/四单放在同一个文件夹里。"),
            ("提示「目标文件正被 Word 打开，写不进去」",
             "先把 Word 里打开的那个 docx 关掉，再点一次生成。"),
            ("生成的 Word 里字变成了方框",
             "那台机器缺 Times New Roman 字体。装一下这个字体，或把母本里的字体换成机器上有的。"),
            ("拖拽没反应 / 拖不进去",
             "用「选择模板表…」按钮选文件，效果完全一样。"
             "（拖拽功能在个别精简版 Windows 上不可用，程序会自动降级。）"),
            ("想改字体、字号、页边距",
             "改程序旁边那个「唛头母本.docx」：用 Word 打开，改好保存，再生成即可。别改文件名。"),
            ("想让程序把某些列也自动转大写",
             "打开程序旁边的 mark_config.json，把列名加到 uppercase_fields 里。"
             "默认只转「品名英文」和「厂家英文」。"),
            ("想让它对错误更严 / 更松",
             "mark_config.json 里：overflow_check.on_overflow 改成 \"error\" 就会因溢出直接停下；"
             "required_fields 里删掉某项，就不再强制填那一项。"),
        ]
        for i, (q, a) in enumerate(FAQ, 1):
            card = tk.Frame(inner, bg=C_CARD, highlightthickness=1,
                            highlightbackground=C_BORDER)
            card.pack(fill="x", padx=18, pady=5)
            tk.Label(card, text="%d. %s" % (i, q), bg=C_CARD, fg=C_TEXT, font=FONT_H2,
                     anchor="w", justify="left", wraplength=860).pack(fill="x", padx=14, pady=(10, 2))
            tk.Label(card, text=a, bg=C_CARD, fg="#3a4250", font=FONT_UI,
                     anchor="w", justify="left", wraplength=860).pack(fill="x", padx=14, pady=(0, 10))
        tk.Frame(inner, bg=C_BG, height=20).pack()

    # ------------------------------------------------------------ 图示（全部用 Canvas 画的）

    def _safe_draw(self, cv, fn):
        try:
            cv.delete("all")
            fn(cv, cv.winfo_width() or 880)
        except Exception:
            cv.delete("all")
            cv.create_text(20, 20, anchor="nw", text="（图示加载失败）", fill=C_MUTED, font=FONT_HINT)

    def _draw_flow(self, cv, w):
        boxes = [
            ("①  拖进来", ["把 唛头模板表.xlsx", "拖到程序窗口里"], "#e8f1ff", C_ACCENT),
            ("②  自动检查", ["缺什么、错在哪", "一次全列出来"], "#fff4e0", "#a8620a"),
            ("③  出 Word 唛头", ["放在模板表", "同一个文件夹里"], "#e6f6ec", C_OK),
        ]
        n = len(boxes)
        gap = 46
        bw = (w - 40 - gap * (n - 1)) / n
        x = 20
        for i, (t, lines, bg, fg) in enumerate(boxes):
            round_rect(cv, x, 34, x + bw, 122, 10, fill=bg, outline=fg, width=2)
            cv.create_text(x + bw / 2, 60, text=t, fill=fg,
                           font=("Microsoft YaHei UI", 11, "bold"), anchor="center")
            for j, ln in enumerate(lines):
                cv.create_text(x + bw / 2, 86 + j * 17, text=ln, fill="#3a4250",
                               font=FONT_HINT, anchor="center")
            if i < n - 1:
                arrow(cv, x + bw + 8, 78, x + bw + gap - 8, 78, color="#9aa6b6")
            x += bw + gap

    def _draw_table(self, cv, w):
        """模板表结构示意图：左边 A..M（从 PL 复制）画成一条带，右边 N..T 逐列画出来。"""
        cv.create_text(18, 10, anchor="nw", text="模板表的列（从「PL箱单」复制 + 手填）",
                       fill=C_TEXT, font=FONT_UI_B)

        left, top, rh = 18, 38, 26
        n_pl = 7                                  # A..M 示意成 7 格
        n_hand = 7                                # N..T 真实 7 列
        total_w = w - 36
        pl_w = total_w * 0.44
        hand_w = total_w - pl_w
        cw = pl_w / n_pl

        # ---- 左侧：A..M
        cv.create_rectangle(left, top, left + pl_w, top + rh * 4, fill="#eaf2ff", outline="")
        for c in range(n_pl + 1):
            cv.create_line(left + c * cw, top, left + c * cw, top + rh * 4, fill="#cdd6e2")
        for r in range(5):
            cv.create_line(left, top + r * rh, left + pl_w, top + r * rh, fill="#cdd6e2")
        cv.create_text(left + pl_w / 2, top + rh * 0.5,
                       text="A … M　从「PL箱单」整列复制粘贴", fill=C_ACCENT,
                       font=FONT_HINT, anchor="center")
        for r, (desc, pkg) in enumerate((("镀锌方矩管", "2"), ("镀锌槽钢", "1")), start=1):
            cv.create_text(left + cw * 2.0, top + rh * r + rh / 2, text=desc,
                           fill=C_TEXT, font=FONT_HINT, anchor="center")
            cv.create_text(left + cw * 4.5, top + rh * r + rh / 2, text=pkg,
                           fill=C_TEXT, font=("Consolas", 9, "bold"), anchor="center")

        # ---- 右侧：N..T（7 列，逐列画）
        hx = left + pl_w
        heads = ["票名", "厂家英文", "发货人", "发货人地址", "收货人", "收货人地址", "原产国"]
        widths = [0.19, 0.23, 0.11, 0.14, 0.11, 0.14, 0.08]
        xs, x = [], hx
        for f in widths:
            xs.append((x, hand_w * f))
            x += hand_w * f
        cv.create_rectangle(hx, top, left + total_w, top + rh * 4, fill="#fff6e5", outline="")
        for i, (x0, ww) in enumerate(xs):
            cv.create_line(x0, top, x0, top + rh * 4, fill="#e2cfa4")
            cv.create_text(x0 + ww / 2, top + rh * 0.5, text=heads[i], fill="#8a6410",
                           font=("Microsoft YaHei UI", 8), anchor="center")
        cv.create_line(left + total_w, top, left + total_w, top + rh * 4, fill="#e2cfa4")
        for r in range(1, 5):
            cv.create_line(hx, top + r * rh, left + total_w, top + r * rh, fill="#e2cfa4")

        # 票名只写第一行，其余留空 → 向下继承
        cv.create_text(xs[0][0] + xs[0][1] / 2, top + rh * 1.5, text="…-1A",
                       fill=C_TEXT, font=("Consolas", 8, "bold"), anchor="center")
        cv.create_text(xs[1][0] + xs[1][1] / 2, top + rh * 1.5, text="HANDAN…",
                       fill=C_TEXT, font=("Consolas", 8), anchor="center")
        for r in (2, 3):
            cv.create_text(xs[0][0] + xs[0][1] / 2, top + rh * r + rh / 2, text="（留空）",
                           fill="#b9a06a", font=("Microsoft YaHei UI", 8), anchor="center")
        cv.create_text(xs[1][0] + xs[1][1] / 2, top + rh * 3.5, text="GUANGXI…",
                       fill=C_TEXT, font=("Consolas", 8), anchor="center")

        # 向下继承箭头
        ax = xs[0][0] + xs[0][1] / 2
        cv.create_line(ax, top + rh * 1.75, ax, top + rh * 2.45,
                       fill=C_ACCENT, width=1, arrow="last")

        y = top + rh * 4 + 14
        cv.create_text(18, y, anchor="nw",
                       text="★  「票名」只写每票第一行，下面留空会自动继承　　"
                            "★  「件数」= 这一行出几个唛头",
                       fill="#3a4250", font=FONT_HINT)
        cv.create_text(18, y + 20, anchor="nw",
                       text="例：第 2 行件数写 2 → 出 2 张唛头，编号接着上一行往下排",
                       fill=C_MUTED, font=FONT_HINT)

    def _draw_output(self, cv, w):
        # 文件夹
        fx, fy = 24, 30
        cv.create_polygon(fx, fy + 8, fx + 52, fy + 8, fx + 60, fy + 20, fx + 200, fy + 20,
                          fx + 200, fy + 110, fx, fy + 110,
                          fill="#ffeec2", outline="#e0b64d", width=2)
        cv.create_text(fx + 100, fy + 68, text="你的项目文件夹", fill="#8a6410",
                       font=FONT_UI_B, anchor="center")

        arrow(cv, 240, 70, 292, 70, color="#9aa6b6")

        rows = [
            ("唛头模板表.xlsx", C_TEXT, "你拖进来的"),
            ("中国三单DEMO-20260101-1A.xlsx", C_MUTED, ""),
            ("越南四单DEMO-20260101-1A.xlsx", C_MUTED, ""),
            ("唛头DEMO-20260101-1.docx", C_OK, "自动生成在这里"),
        ]
        y = 26
        for name, col, note in rows:
            cv.create_rectangle(310, y, 328, y + 16, fill="#dfe7f2", outline="#b9c4d2")
            cv.create_text(338, y + 8, anchor="w", text=name, fill=col,
                           font=("Consolas", 9) if name.endswith(".xlsx") else FONT_UI_B)
            if note:
                cv.create_text(w - 20, y + 8, anchor="e", text=note, fill=col, font=FONT_HINT)
            y += 32
        cv.create_text(310, y + 6, anchor="w",
                       text="文件名自动取表里的票名：DEMO-20260101-1A → 唛头DEMO-20260101-1.docx",
                       fill=C_MUTED, font=FONT_HINT)

    # ------------------------------------------------------------ 日志

    def _log(self, level, text):
        self.log.configure(state="normal")
        tag = level if level in ("h", "ok", "err", "warn", "info", "dim", "advice") else "info"
        self.log.insert("end", text + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")
        self.update_idletasks()

    def clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _greet(self):
        self._log("h", "准备就绪。把「唛头模板表.xlsx」拖到上面的框里，或者点「选择模板表…」。")
        if not _HAS_DND:
            self._log("warn", "这台机器上的拖拽组件没加载成功 —— 用「选择模板表…」按钮选文件，"
                              "效果完全一样。")
        self._log("info", "拖进来之后程序会先自动检查一遍，有问题会直接告诉你改哪里。")

    # ------------------------------------------------------------ 生成动作

    def _set_busy(self, on):
        self.busy = on
        state = "disabled" if on else "normal"
        for b in (self.btn_choose, self.btn_check, self.btn_go):
            b.configure(state=state)
        if not on and not self.template:
            self.btn_check.configure(state="disabled")
            self.btn_go.configure(state="disabled")
        self.configure(cursor="watch" if on else "")
        self.update_idletasks()

    def do_work(self, mode, silent_ok=False):
        if self.busy:
            return
        if not self.template:
            messagebox.showinfo(APP_TITLE, "先把「唛头模板表.xlsx」拖进来，或点「选择模板表…」选一个。")
            return
        if not os.path.exists(self.template):
            self._log("err", "文件不在了：%s" % self.template)
            self._log("advice", "可能被移动或改名了。请重新拖一次。")
            return
        self._set_busy(True)
        try:
            try:
                cfg = load_config()
            except Exception as e:
                self._log("err", str(e))
                return
            self.res = run_check(self.template, cfg)
            self._report(self.res, mode)
            if self.res.fatal:
                return
            if self.res.errors:
                return
            if mode == "check":
                if not silent_ok:
                    self._log("ok", "检查通过。可以点「开始生成」了。")
                return
            # 需要用户拍板的两件事：PL 对不上、内容太长会挤到第二页
            if self.res.needs_confirm:
                if not self._ask_continue(self.res.needs_confirm):
                    self._log("warn", "你选了「先不生成」。改好之后再来。")
                    return
                self._log("warn", "你选择继续生成（已忽略上面的提示）。")
            run_build(self.res, self.template)
            if self.res.written:
                self.last_out = self.res.out_path
                self.btn_res.configure(state="normal")
                self.btn_resdir.configure(state="normal")
                self.lbl_out.configure(text="已生成：%s" % self.last_out, fg=C_OK)
                if messagebox.askyesno(APP_TITLE,
                                       "生成完成！\n\n%s\n\n共 %d 页。\n\n现在打开看看吗？"
                                       % (self.last_out, self.res.n_pages)):
                    open_file(self.last_out)
            else:
                self._log("err", self.res.fatal or "生成失败。")
        except Exception:
            self._log("err", "程序内部出错：\n" + traceback.format_exc(limit=4))
            self._log("advice", "把上面这段发给做这个工具的人。")
        finally:
            self._set_busy(False)
            self._update_buttons()

    def _update_buttons(self):
        ok = bool(self.template) and os.path.exists(self.template or "")
        self.btn_check.configure(state="normal" if ok else "disabled")
        self.btn_go.configure(state="normal" if ok else "disabled")
        self.btn_open_tpl.configure(state="normal" if ok else "disabled")
        self.btn_open_dir.configure(state="normal" if ok else "disabled")

    def _ask_continue(self, items):
        """生成前的「要不要继续」弹窗 —— 把后果讲清楚，再让用户决定。"""
        pl = [i for i in items if i.get("code") == "pl_mismatch"]
        ov = [i for i in items if i.get("code") == "overflow"]
        parts = []
        if ov:
            n_lines = len(ov)
            parts.append("【一页放不下】有 %d 个唛头的内容太长，预估会挤到第二页。\n"
                         "后果：那个唛头会被拆到两页上，而且它后面跟着硬分页符，"
                         "后面的唛头会整体错位、页码也跟着变。\n"
                         "建议先按下面点名的段落精简一下再生成。\n" % n_lines)
            for i in ov[:4]:
                parts.append("  · %s\n    → %s\n" % (i["text"], i["advice"]))
            if len(ov) > 4:
                parts.append("  … 还有 %d 条，见下方信息区\n" % (len(ov) - 4))
        if pl:
            parts.append("\n【件数对不上】有 %d 条：模板表的件数合计与三单/四单 PL箱单 的 "
                         "Packages 合计不一致。\n"
                         "如果你知道确实该不一样（比如 PL 还没更新），忽略即可。\n" % len(pl))
            for i in pl[:4]:
                parts.append("  · %s\n" % i["text"])
            if len(pl) > 4:
                parts.append("  … 还有 %d 条\n" % (len(pl) - 4))
        parts.append("\n确定要现在生成吗？")
        return messagebox.askyesno(APP_TITLE, "".join(parts), default="no")

    def _report(self, res, mode):
        """把检查结果按「先坏消息、后好消息」的顺序写进日志。"""
        for lvl, text in res.log:
            if lvl == "h":
                self._log("h", text)
            else:
                self._log(lvl, text)

        if res.fatal:
            self._log("err", res.fatal)
            self._log("advice", "照上面说的改完，再点一次「开始生成」。")
            return

        errs = res.errors
        if errs:
            self._log("h", "✕  有 %d 处必须先改掉，改完再点一次生成：" % len(errs))
            for i, it in enumerate(errs, 1):
                self._log("err", "%d) %s" % (i, it["text"]))
                if it["advice"]:
                    self._log("advice", "→ %s" % it["advice"])
            self._log("info", "提示：点上面的「打开模板表」可以直接跳到那个文件。")
            return

        confs = res.confirms
        if confs:
            self._log("h", "!  件数核对：有 %d 条对不上（不拦你）：" % len(confs))
            for it in confs:
                self._log("confirm", "· %s" % it["text"])
            if confs[0]["advice"]:
                self._log("advice", "→ %s" % confs[0]["advice"])

        ovs = [w for w in res.warns if w.get("code") == "overflow"]
        others = [w for w in res.warns if w.get("code") != "overflow"]
        if ovs:
            self._log("h", "!  一页放不下：有 %d 个唛头会挤到第二页（会拆成两页、后面整体错位）："
                      % len(ovs))
            for it in ovs:
                self._log("warn", "· %s" % it["text"])
                if it["advice"]:
                    self._log("advice", "→ %s" % it["advice"])
            self._log("info", "改完再点「先检查不生成」看看，行数降下来就没这个提示了。")
        for it in others:
            self._log("warn", "· %s" % it["text"])
            if it["advice"]:
                self._log("advice", "→ %s" % it["advice"])

        if mode == "check" and not confs and not ovs and not others:
            self._log("ok", "全部检查通过，没发现任何问题。")

    # ------------------------------------------------------------ 打开文件

    def open_template(self):
        if self.template:
            open_file(self.template)

    def open_template_dir(self):
        if self.template:
            open_in_explorer(self.template, select=True)

    def open_result(self):
        if self.last_out and os.path.exists(self.last_out):
            open_file(self.last_out)
        else:
            self._log("warn", "还没有生成过文件。")

    def open_result_dir(self):
        if self.last_out and os.path.exists(self.last_out):
            open_in_explorer(self.last_out, select=True)
        elif self.template:
            open_in_explorer(self.template)

    # ------------------------------------------------------------ 状态栏

    def _build_statusbar(self):
        bar = tk.Frame(self, bg="#e4e8ee", height=24)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        self.status = tk.Label(bar, text="", bg="#e4e8ee", fg=C_MUTED, font=FONT_HINT,
                               anchor="w")
        self.status.pack(side="left", padx=14)
        tk.Label(bar, text="拖拽" + ("可用" if _HAS_DND else "不可用（请用按钮选择）"),
                 bg="#e4e8ee", fg=C_MUTED, font=FONT_HINT).pack(side="right", padx=14)
        self._set_status("把模板表拖进来就能开始")

    def _set_status(self, t):
        self.status.configure(text=t)


# ---------------------------------------------------------------- 自检

def _report_sink(name="界面自检结果.txt"):
    """★ 打包成 --windowed 的 exe 后没有控制台，print 等于扔进黑洞。

    所以自检 / --check 的结果同时写一份到 exe 旁边的文件里，
    这样既能给人看，也能被自动化测试读到。
    """
    lines = []

    class _Tee(object):
        def write(self, s):
            if s:
                lines.append(s)
            return len(s or "")

        def flush(self):
            pass

        def reconfigure(self, **kw):
            pass

        def isatty(self):
            return False

        def dump(self):
            p = os.path.join(core.base_dir(), name)
            try:
                with open(p, "w", encoding="utf-8") as f:
                    f.write("".join(lines))
                return p
            except Exception:
                return None

    t = _Tee()
    return t


def _gui_selftest():
    """建一遍全部界面与图示，确认没有布局/绘制异常，然后退出（回归用）。"""
    if not _HAS_TK:
        print("界面自检：跳过（这个 Python 没有 tkinter）")
        return 0
    app = App()
    app.withdraw()
    problems = []
    try:
        app.update_idletasks()
        app.update()
        # 三个页签都要能画
        for i in range(3):
            app.nb.select(i)
            app.update_idletasks()
            app.update()
        # 拖拽框的三种状态
        for st in ("idle", "hover", "loaded"):
            app._drop_state = st
            app._draw_drop()
        # 图示逐个重画
        for cv, fn in _collect_diagrams(app):
            app._safe_draw(cv, fn)
        # 日志的每个 tag 都能用
        for lvl in ("h", "ok", "err", "warn", "confirm", "info", "dim", "advice"):
            app._log(lvl, "自检 %s" % lvl)
        app.clear_log()
        # 关键控件都在
        for name in ("drop", "log", "btn_choose", "btn_check", "btn_go",
                     "btn_open_tpl", "btn_open_dir", "btn_res", "lbl_out"):
            if not hasattr(app, name):
                problems.append("缺少控件 %s" % name)
    except Exception:
        problems.append(traceback.format_exc())
    finally:
        app.destroy()

    print("界面自检：")
    print("  拖拽组件：%s" % ("可用" if _HAS_DND else "不可用（会降级成按钮选择）"))
    print("  字体：%s" % (FONT_UI[0],))
    if problems:
        for p in problems:
            print("  × %s" % p)
        return 1
    print("  三个页签、拖拽框三态、全部图示、日志标签、关键控件 —— 全部通过 √")
    return 0


def _collect_diagrams(app):
    out = []
    for tab in (app.tab_tut,):
        stack = [tab]
        while stack:
            w = stack.pop()
            for ch in w.winfo_children():
                if isinstance(ch, tk.Canvas) and ch is not app.drop:
                    out.append(ch)
                stack.append(ch)
    # 教程页的图示：按出现顺序对应三个 draw 函数
    fns = [app._draw_flow, app._draw_table, app._draw_output]
    return [(cv, fns[i % 3]) for i, cv in enumerate(out)]


def _cli_check(path):
    """不开窗口，只跑检查并打印（给测试用）。"""
    res = run_check(path)
    for lvl, text in res.log:
        if lvl in ("h", "ok", "err", "warn", "info", "dim"):
            print(text)
    if res.fatal:
        print("FATAL: %s" % res.fatal)
        return 2
    for it in res.errors:
        print("ERROR: %s" % it["text"])
        if it["advice"]:
            print("  → %s" % it["advice"])
    for it in res.confirms:
        print("CONFIRM: %s" % it["text"])
        if it["advice"]:
            print("  → %s" % it["advice"])
    for it in res.warns:
        print("WARN: %s" % it["text"])
        if it["advice"]:
            print("  → %s" % it["advice"])
    print("OUTPUT: %s" % res.out_path)
    return 1 if res.errors else 0


def _fix_streams():
    """★ 打包成 `--windowed` 的 exe 时，Windows 下 sys.stdout / sys.stderr 是 None。

    核心逻辑里到处 `print()` 和 `sys.stderr.write()`，不兜住会直接 AttributeError。
    这里补一个黑洞 writer；真要输出时 run_check 会把它们换成日志收集器。
    """
    class _Null(object):
        def write(self, *a):
            return 0

        def flush(self):
            pass

        def reconfigure(self, **kw):
            pass

        def isatty(self):
            return False

    for name in ("stdout", "stderr"):
        if getattr(sys, name, None) is None:
            setattr(sys, name, _Null())


def _install_crash_handler():
    """任何没接住的异常都要让用户看得见、也留得下痕迹（windowed 模式没有黑框）。"""
    def hook(exc_type, exc, tb):
        import traceback as _tb
        text = "".join(_tb.format_exception(exc_type, exc, tb))
        try:
            p = os.path.join(core.base_dir(), "唛头生成器_出错日志.txt")
            with open(p, "a", encoding="utf-8") as f:
                f.write("\n" + "=" * 70 + "\n" + text)
        except Exception:
            p = "(写日志失败)"
        try:
            messagebox.showerror(
                APP_TITLE,
                "程序遇到了没预料到的问题：\n\n%s\n\n"
                "详细信息已经写到：\n%s\n\n"
                "把这个文件发给做这个工具的人即可。" % (exc, p))
        except Exception:
            pass
        sys.__excepthook__(exc_type, exc, tb)

    sys.excepthook = hook


def _cli_build(path, assume_yes=False):
    """★ 不开窗口，走**和图形界面完全一样**的流程生成一份 docx（给自动化测试用）。

    这样「输出落在模板表所在文件夹」这条规则在 exe 上也能被验证，不用手点按钮。
    退出码：0 成功 / 1 有必须改的错误 / 2 生成失败 / 3 有要确认的事但没给 --yes
    """
    res = run_check(path)
    for lvl, text in res.log:
        if lvl in ("h", "ok", "err", "warn", "info", "dim"):
            print(text)
    if res.fatal:
        print("FATAL: %s" % res.fatal)
        return 2
    for it in res.errors:
        print("ERROR: %s" % it["text"])
        if it["advice"]:
            print("  → %s" % it["advice"])
    if res.errors:
        return 1
    for it in res.needs_confirm:
        print("CONFIRM: %s" % it["text"])
        if it["advice"]:
            print("  → %s" % it["advice"])
    if res.needs_confirm and not assume_yes:
        print("（有需要确认的事项，非交互模式下先停下。加 --yes 可继续。）")
        return 3
    run_build(res, path)
    if not res.written:
        print("FATAL: %s" % (res.fatal or "生成失败"))
        return 2
    print("OUTPUT: %s" % res.out_path)
    print("PAGES: %d" % res.n_pages)
    return 0


# 这些参数 = 命令行版的，交给 make_mark 处理（同一个 exe 兼做命令行版）
CLI_ARGS = ("-i", "--input", "-o", "--out", "-c", "--config", "-V", "--version",
            "--dry-run", "--no-pl-check", "--self-test", "-h", "--help")


def _attach_console():
    """★ windowed exe 没有控制台 —— 从 cmd 调用时挂到父进程的控制台上，输出才看得见。

    挂不上也无所谓：结果同时写进了结果文件，bat 里 `type` 出来一样能看。
    """
    if not getattr(sys, "frozen", False) or os.name != "nt":
        return False
    try:
        import ctypes
        k = ctypes.windll.kernel32
        if not k.AttachConsole(-1):          # ATTACH_PARENT_PROCESS
            return False
        for name, handle, mode in (("stdout", -11, "w"), ("stderr", -12, "w")):
            try:
                h = k.GetStdHandle(handle)
                fd = _open_osfhandle(h, 0)
                setattr(sys, name, os.fdopen(fd, mode, encoding="utf-8", errors="replace"))
            except Exception:
                pass
        return True
    except Exception:
        return False


def _open_osfhandle(handle, flags):
    import msvcrt
    return msvcrt.open_osfhandle(handle, flags)


def _cli_passthrough(argv):
    """把命令行参数转给 make_mark，输出同时写到结果文件（windowed 下看不到控制台）。"""
    attached = _attach_console()
    tee = _report_sink()
    old_out, old_err, old_argv = sys.stdout, sys.stderr, sys.argv
    sys.stdout = sys.stderr = tee
    sys.argv = [APP_TITLE] + list(argv)
    try:
        rc = core.main()
    except SystemExit as e:
        rc = e.code if isinstance(e.code, int) else (0 if e.code is None else 2)
    except Exception:
        sys.stdout = sys.stderr = old_out
        tee.write(traceback.format_exc())
        rc = 2
    finally:
        sys.stdout, sys.stderr, sys.argv = old_out, old_err, old_argv
    tee.dump()
    return rc or 0


def main(argv=None):
    _fix_streams()
    argv = list(sys.argv[1:] if argv is None else argv)
    if argv and argv[0] == "--selftest-gui":
        tee = _report_sink()
        old, sys.stdout = sys.stdout, tee
        try:
            rc = _gui_selftest()
        finally:
            sys.stdout = old
        p = tee.dump()
        if p:
            print("自检结果已写入 %s（退出码 %d）" % (p, rc))
        return rc
    if argv and argv[0] == "--check":
        if len(argv) < 2:
            print("用法：唛头生成器.exe --check 模板表.xlsx")
            return 2
        tee = _report_sink()
        old, sys.stdout = sys.stdout, tee
        try:
            rc = _cli_check(argv[1])
        finally:
            sys.stdout = old
        tee.dump()
        return rc
    if argv and argv[0] == "--build":
        rest = argv[1:]
        if not rest:
            print("用法：唛头生成器.exe --build 模板表.xlsx [--yes]")
            return 2
        tee = _report_sink()
        old, sys.stdout = sys.stdout, tee
        try:
            rc = _cli_build(rest[0], assume_yes=("--yes" in rest))
        finally:
            sys.stdout = old
        tee.dump()
        return rc
    # 命令行版的参数 → 转给核心逻辑（这样只需要一个 exe）
    if argv and argv[0] in CLI_ARGS:
        return _cli_passthrough(argv)
    if not _HAS_TK:
        sys.stderr.write("这台机器上的 Python 没有 tkinter，开不了图形界面。\n"
                         "命令行用法：唛头生成器.exe -i 模板表.xlsx -o 输出.docx\n")
        return 3
    _install_crash_handler()
    app = App()
    if argv and is_template_file(argv[0]):
        app.after(200, lambda: app.set_template(argv[0]))
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
