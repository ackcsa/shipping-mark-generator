# -*- coding: utf-8 -*-
"""GUI 版回归测试：不开窗口，直接跑 `make_mark_gui.py --check`，逐条断言。

覆盖三件必须靠得住的事：
  ① 输出位置 —— 必须落在**模板表所在文件夹**（用户明确要求）
  ② 报错友好 —— 每条错误都要有「行号 + 单元格」和「怎么改」；该宽容的要宽容
  ③ 不崩       —— 各种脏数据都不能让程序抛异常

跑法：
    python gui_tests.py              全部
    python gui_tests.py 测试7        只跑名字含「测试7」的
"""
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)                       # 唛头生成器/
GUI_PY = os.path.join(ROOT, "make_mark_gui.py")
GUI_EXE = os.path.join(ROOT, "唛头生成器.exe")      # 发行版目录里只有 exe

PY = sys.executable
FROZEN = bool(getattr(sys, "frozen", False))


def run_check(xlsx):
    """跑一次 --check，返回 (退出码, stdout+stderr)。

    优先用 exe（发行版目录）；没有 exe 就用当前 Python 跑源码（开发目录）。
    """
    return _run_gui(["--check", xlsx])


def run_build(xlsx, yes=True):
    """跑一次 --build（和图形界面同一条流程，只是不开窗口）。"""
    args = ["--build", xlsx] + (["--yes"] if yes else [])
    return _run_gui(args)


def _run_gui(args):
    cmd = [GUI_EXE] if os.path.exists(GUI_EXE) else [PY, GUI_PY]
    cmd += args
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       errors="replace", env=env, stdin=subprocess.DEVNULL)
    out = (r.stdout or "") + (r.stderr or "")
    # --windowed 的 exe 没有控制台，结果会写到「界面自检结果.txt」
    if not out.strip():
        rp = os.path.join(ROOT, "界面自检结果.txt")
        if os.path.exists(rp):
            with open(rp, encoding="utf-8") as f:
                out = f.read()
    return r.returncode, out


def lines_of(text, tag):
    return [ln[len(tag) + 1:].strip() for ln in text.splitlines() if ln.startswith(tag + ":")]


class Case(object):
    def __init__(self, name, xlsx, expect_rc=None, errors=None, warns=None,
                 confirms=None, must_contain=(), must_not_contain=(),
                 advice_required=False, out_dir=None):
        self.name, self.xlsx = name, xlsx
        self.expect_rc = expect_rc
        self.errors, self.warns, self.confirms = errors, warns, confirms
        self.must_contain, self.must_not_contain = list(must_contain), list(must_not_contain)
        self.advice_required = advice_required
        self.out_dir = out_dir


CASES = [
    Case("测试1_一票多物", "测试1_一票多物.xlsx", errors=0,
         must_contain=["9 个唛头", "SQUARE PIPE", "U CHANNEL STEEL"]),
    Case("测试2_多票一物", "测试2_多票一物.xlsx", errors=0),
    Case("测试3_票名不连续", "测试3_票名不连续.xlsx", errors=0),
    Case("测试4_字符兼容性", "测试4_字符兼容性.xlsx", errors=0),
    Case("测试5_特殊格式", "测试5_特殊格式.xlsx", errors=0),
    Case("测试6_单页溢出", "测试6_单页溢出.xlsx", errors=0, warns=2,
         must_contain=["会溢出到第二页", "CNEE", "缩短约"]),
    Case("负面3_必填缺失", "负面3_必填缺失.xlsx", errors=3, advice_required=True,
         must_contain=["「票名」为空", "「品名英文」为空", "「收货人地址」为空",
                       "单元格 N2", "单元格 B3", "单元格 S2"]),
    Case("负面4_件数非法", "负面4_件数非法.xlsx", errors=2, advice_required=True,
         must_contain=["不是整数", "必须 ≥ 1", "单元格 G3", "单元格 G4"]),
    Case("测试9_多余列", "测试9_多余列.xlsx", errors=0,
         must_contain=["用不到，已跳过"]),        # 用不到的列只提示、不报错
    Case("测试10_最小单张", "测试10_最小单张.xlsx", errors=0, must_contain=["1 个唛头"]),
    Case("测试11_全角件数", "测试11_全角件数.xlsx", errors=0),
    Case("负面1_公式无缓存", "负面1_公式无缓存.xlsx", errors=1, advice_required=True,
         must_contain=["公式", "另存"]),
    Case("负面2_缺件数列", "负面2_缺件数列.xlsx", errors=None, advice_required=True,
         must_contain=["缺字段「件数」"]),
]


def check_one(c):
    xlsx = os.path.join(HERE, c.xlsx)
    problems = []
    if not os.path.exists(xlsx):
        return ["用例文件不存在：%s" % xlsx], None, None

    rc, text = run_check(xlsx)
    errs = lines_of(text, "ERROR")
    warns = lines_of(text, "WARN")
    confirms = lines_of(text, "CONFIRM")
    outs = lines_of(text, "OUTPUT")
    fatal = lines_of(text, "FATAL")

    # ③ 不许崩
    if "Traceback" in text:
        problems.append("抛异常了：\n%s" % text[-600:])
    if fatal:
        problems.append("致命错误：%s" % fatal[0])

    # ① 输出位置必须 = 模板表所在文件夹
    if outs and outs[0] not in ("None", ""):
        out = outs[0]
        if os.path.dirname(os.path.abspath(out)) != os.path.abspath(HERE):
            problems.append("输出位置不对：%s（应在 %s）" % (out, HERE))
        if not os.path.basename(out).startswith("唛头"):
            problems.append("输出文件名不以「唛头」开头：%s" % out)
    elif not outs:
        problems.append("没有打印 OUTPUT 行")
    elif c.errors in (0, None) and c.name not in ("负面3_必填缺失", "负面4_件数非法",
                                                  "负面1_公式无缓存", "负面2_缺件数列"):
        # 校验没通过的用例，输出路径可以为 None；通过了的必须有
        problems.append("没有算出输出路径（OUTPUT 为空）")

    # ② 错误条数
    if c.errors is not None and len(errs) != c.errors:
        problems.append("错误条数 %d，期望 %d\n  实际：%s"
                        % (len(errs), c.errors, " | ".join(errs)[:400]))
    if c.warns is not None and len(warns) < c.warns:
        problems.append("警告条数 %d，期望至少 %d" % (len(warns), c.warns))
    if c.confirms is not None and len(confirms) != c.confirms:
        problems.append("待确认条数 %d，期望 %d" % (len(confirms), c.confirms))

    # 退出码
    if c.expect_rc is not None and rc != c.expect_rc:
        problems.append("退出码 %d，期望 %d" % (rc, c.expect_rc))

    body = text
    for s in c.must_contain:
        if s not in body:
            problems.append("输出里找不到「%s」" % s)
    for s in c.must_not_contain:
        if s in body:
            problems.append("输出里不该出现「%s」" % s)

    # 每条错误都必须带「怎么改」
    if c.advice_required:
        for e in errs:
            if "→" not in text.split(e, 1)[1][:400]:
                problems.append("错误没给出改法：%s" % e[:90])
                break

    return problems, rc, text


def test_e2e_build():
    """端到端：把模板放到一个临时子目录里，走**和图形界面同一条流程**真出一份 docx。

    重点验三件事（都是用户明确要求的）：
      ① 产出落在**模板表所在的那个文件夹**（不是 exe 目录、不是当前目录）
      ② 页数 = 唛头个数
      ③ 重复生成能覆盖，不会报错
    """
    import shutil
    import tempfile
    import zipfile

    problems = []
    tmp = tempfile.mkdtemp(prefix="mark_e2e_")
    try:
        sub = os.path.join(tmp, "某个项目文件夹")      # 带中文的路径
        os.makedirs(sub)
        src = os.path.join(HERE, "测试1_一票多物.xlsx")
        tpl = os.path.join(sub, "唛头模板表.xlsx")
        shutil.copy(src, tpl)

        # 检查阶段
        rc, out = run_check(tpl)
        if rc != 0:
            return ["检查阶段返回 %d：\n%s" % (rc, out[-400:])]
        want = os.path.join(sub, "唛头TEST-MIX-1.docx")   # 测试1 的第一个票名是 TEST-MIX-1
        got = [ln[len("OUTPUT: "):].strip() for ln in out.splitlines()
               if ln.startswith("OUTPUT: ")]
        if not got:
            problems.append("没有打印 OUTPUT")
        else:
            if os.path.abspath(got[0]) != os.path.abspath(want):
                problems.append("输出路径 %s，期望 %s" % (got[0], want))
            if os.path.dirname(os.path.abspath(got[0])) != os.path.abspath(sub):
                problems.append("输出没有落在模板表所在文件夹")

        # 生成阶段
        rc, out = run_build(tpl)
        if rc != 0:
            return problems + ["生成阶段返回 %d：\n%s" % (rc, out[-400:])]
        if not os.path.exists(want):
            return problems + ["没有生成文件：%s" % want]

        pages = [ln[len("PAGES: "):].strip() for ln in out.splitlines()
                 if ln.startswith("PAGES: ")]
        if not pages or pages[0] != "9":
            problems.append("PAGES = %s，期望 9" % (pages or "无"))

        d = zipfile.ZipFile(want).read("word/document.xml").decode("utf-8")
        npage = d.count('w:type="page"') + 1
        if npage != 9:
            problems.append("document.xml 里分页符+1 = %d，期望 9" % npage)
        if "SQUARE PIPE" not in d or "U CHANNEL STEEL" not in d:
            problems.append("产出里没有品名")

        # 再生成一次：应当直接覆盖，不报错
        rc2, out2 = run_build(tpl)
        if rc2 != 0:
            problems.append("第二次生成失败（返回 %d）" % rc2)
        if not os.path.exists(want):
            problems.append("第二次生成后文件不见了")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return problems


def main():
    only = sys.argv[1] if len(sys.argv) > 1 else None
    cases = [c for c in CASES if not only or only in c.name]
    print("GUI 检查模式回归（%d 个用例）\n%s" % (len(cases), "=" * 88))
    failed = 0
    for c in cases:
        problems, rc, _ = check_one(c)
        if problems:
            failed += 1
            print("× %s" % c.name)
            for p in problems:
                print("    - %s" % p)
        else:
            print("√ %s" % c.name)

    if not only:
        print("-" * 88)
        problems = test_e2e_build()
        if problems:
            failed += 1
            print("× 端到端：生成一份真 docx")
            for p in problems:
                print("    - %s" % p)
        else:
            print("√ 端到端：生成一份真 docx（落在模板同目录、9 页、可覆盖）")

    total = len(cases) + (0 if only else 1)
    print("=" * 88)
    print("%d/%d 通过" % (total - failed, total))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
