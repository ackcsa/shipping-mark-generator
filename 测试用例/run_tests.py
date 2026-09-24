#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
唛头生成器 —— 测试用例一键回归

对同目录下所有 `*.xlsx` 逐个跑生成，然后把产出的 docx 逐块解析出来打印，并做断言。
文件名以 `负面` 开头的用例 = **期望失败**（退出码必须非 0，且不得产出 docx）。

用法：
    python run_tests.py                 # 跑全部
    python run_tests.py 测试4           # 只跑文件名含「测试4」的
"""

import glob
import html
import os
import re
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile

# 中文 Windows 的 cmd 默认 936 代码页 —— 控制台输出不能因编码问题崩掉
def _make_console_safe():
    for _s in (sys.stdout, sys.stderr):
        try:
            _s.reconfigure(errors="replace")
        except Exception:
            pass


_make_console_safe()


HERE = os.path.dirname(os.path.abspath(__file__))
GEN = os.path.dirname(HERE)                      # 唛头生成器/
PY = sys.executable


def gen_cmd(*args):
    """怎么调生成器：① 有 exe 就用 exe（发行版目录）；② 否则用当前 Python 跑源码。

    发行版里 GUI 和命令行是**同一个 exe**（唛头生成器.exe），带 -i/-o 就是命令行模式。
    """
    for name in ("唛头生成器.exe", "ShippingMarkGenerator.exe"):
        exe = os.path.join(GEN, name)
        if os.path.exists(exe):
            return [exe] + list(args)
    return [PY, os.path.join(GEN, "make_mark.py")] + list(args)


def parse_docx(path):
    """把一个唛头 docx 拆成块。返回 (块列表, 分页符数, paraId唯一, XML是否合法, 原始XML)。"""
    d = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8")
    try:
        ET.fromstring(d)
        xml_ok = True
    except ET.ParseError:
        xml_ok = False
    paras = re.findall(r"<w:p[ >].*?</w:p>", d, re.S)
    txt = ["".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", pa, re.S)) for pa in paras]
    blocks, cur = [], []
    for pa, t in zip(paras, txt):
        if 'w:type="page"' in pa:
            blocks.append(cur)
            cur = []
            continue
        cur.append(t)
    blocks.append(cur)
    while blocks and blocks[-1] and blocks[-1][-1] == "":
        blocks[-1].pop()

    out = []
    for b in blocks:
        if not b or not b[0].startswith("SHIPPER:"):
            continue
        g = lambda k: next((x.split(":", 1)[1].strip() for x in b if x.startswith(k + ":")), "")
        out.append({"commodity": g("COMMODITY"), "maker": g("MANUFACTURER"),
                    "total": g("TOTAL"), "no": g("PACKAGE NO"),
                    "lines": len([x for x in b if x != ""]) + 2})
    ids = re.findall(r'w14:paraId="([0-9A-Fa-f]+)"', d)
    return out, d.count('w:type="page"'), len(ids) == len(set(ids)), xml_ok, d


def run_one(x):
    name = os.path.splitext(os.path.basename(x))[0]
    negative = name.startswith("负面")
    out = os.path.join(HERE, "唛头_%s.docx" % name)
    if os.path.exists(out):
        os.remove(out)
    print("=" * 100)
    print("■ %s   %s" % (name, "（负面用例：期望报错停下）" if negative else ""))
    r = subprocess.run(gen_cmd("-i", x, "-o", out, "--no-pl-check"),
                       cwd=GEN, capture_output=True, text=True, encoding="utf-8",
                       errors="replace")
    body = (r.stdout or "") + (r.stderr or "")
    if negative:
        ok = (r.returncode != 0) and (not os.path.exists(out))
        print("   退出码 %d（应非 0）%s | 未产出文件 %s"
              % (r.returncode, "√" if r.returncode != 0 else "×",
                 "√" if not os.path.exists(out) else "×"))
        for l in body.splitlines():
            if "[错误]" in l or re.match(r"\s*\d+\)", l) or "[提示]" in l:
                print("   " + l.strip())
        print("   %s" % ("√ 符合预期" if ok else "× 与预期不符"))
        return ok

    sys.stdout.write(r.stdout)
    if r.returncode != 0:
        sys.stderr.write(r.stderr)
        print("   × 生成失败，退出码 %d" % r.returncode)
        return False

    blocks, breaks, uniq, xml_ok, xml = parse_docx(out)
    nbr = xml.count("<w:br/>")
    print("   块数 %d / 分页符 %d / paraId 唯一 %s / XML 合法 %s / 值内换行符 <w:br/> %d 个"
          % (len(blocks), breaks, uniq, xml_ok, nbr))
    print("    #   TOTAL         PACKAGE NO   COMMODITY                 MANUFACTURER")
    for i, b in enumerate(blocks, 1):
        print("    %-3d %-13s %-12s %-25s %s"
              % (i, b["total"], b["no"], html.unescape(b["commodity"]),
                 html.unescape(b["maker"])))
    ok = True
    checks = [
        ("块数 = 分页符+1", len(blocks) == breaks + 1),
        ("每块 11 行", all(b["lines"] == 11 for b in blocks)),
        ("paraId 唯一", uniq),
        ("document.xml 合法", xml_ok),
        ("PACKAGE NO 形如 N-i", all(re.fullmatch(r"\d+-\d+", b["no"]) for b in blocks)),
        ("TOTAL 与 NO 前半段一致",
         all(b["no"].split("-")[0] == b["total"].split()[0] for b in blocks)),
        ("& 已转义成 &amp;", "&amp;" in xml or "&" not in xml),
        ("无 XML 非法控制字符", not re.search(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", xml)),
        ("<w:t> 里没有裸换行", not re.search(r"<w:t[^>]*>[^<]*\n", xml)),
    ]
    for label, good in checks:
        print("   %s %s" % ("√" if good else "×", label))
        ok &= good
    return ok


def main():
    key = sys.argv[1] if len(sys.argv) > 1 else ""
    files = sorted(f for f in glob.glob(os.path.join(HERE, "*.xlsx"))
                   if not os.path.basename(f).startswith("~$"))
    if key:
        files = [f for f in files if key in os.path.basename(f)]
    if not files:
        print("没找到 xlsx 用例")
        return 1

    ok_all = True
    for x in files:
        ok_all &= run_one(x)
    print("=" * 100)
    print("全部通过 √" if ok_all else "有失败项 ×")
    return 0 if ok_all else 1


if __name__ == "__main__":
    sys.exit(main())
