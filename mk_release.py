# -*- coding: utf-8 -*-
"""组装发行版目录 `发行版/唛头生成器_v2.0.0/` 并打 zip。

★ 用 **onedir**（不是 onefile）：onefile 每次启动都要把 1000 个文件解到 %TEMP%
  再让杀软逐个扫，实测 GUI 版启动要 90 秒以上；onedir 只要 1.6 秒。

跑法（在 唛头生成器/ 下）：python mk_release.py
"""
import json
import os
import shutil

HERE = os.path.dirname(os.path.abspath(__file__))          # 唛头生成器/
ROOT = os.path.dirname(HERE)                               # DEMO-20260101/
VER = "2.0.0"
REL = os.path.join(ROOT, "发行版", "唛头生成器_v%s" % VER)
APP = os.path.join(ROOT, ".workbuddy-ai", "tmp", "guibuild", "dist_app", "唛头生成器")
EXE_NAME = "唛头生成器.exe"

README = """唛头生成器 v{ver}  —— 使用说明（公司内部）
================================================================

一、这是什么
    把「唛头模板表」（Excel）里的内容，批量生成一份 Word 唛头。
    一件货一个唛头、一页放一张，块与块之间自动分页 —— 不用手动排版。
    输出固定写在模板表所在的文件夹里，和你的三单/四单放在一起。

二、怎么用（推荐：图形界面）
    双击「唛头生成器.exe」→ 把「唛头模板表.xlsx」拖进窗口（或点「选择模板表…」）
    → 程序先自动检查一遍 → 点「开始生成」。
    窗口里有「使用教程」和「常见问题」两个页签，讲得比这份文件还细。

    拖拽不好使的机器，用「选择模板表…」按钮选文件，效果完全一样；
    或者干脆把模板表拖到「5-拖模板到这里生成.bat」上。

三、模板表怎么填
    A..M 列：与「PL箱单」同名同序，整列复制粘贴即可。
             程序只用得到 Descriptions（品名）和 Packages（件数），其余列留着方便你复制。
    N..T 列：PL 里没有的，要手填 ——
             票名 | 厂家英文 | 发货人 | 发货人地址 | 收货人 | 收货人地址 | 原产国
    票名：只写每票的第一行，下面留空会自动继承。
    件数：这一行要出几个唛头（这一行有 2 件货就写 2，程序出 2 张）。
    编号是自动的：PACKAGE NO = 本票件数合计-序号（如 13-1 … 13-13）。

四、程序会帮你检查什么
    · 必填项有没有空着 —— 会告诉你第几行、哪个单元格、怎么改。
    · 件数是不是 1 以上的整数。
    · 件数合计和同文件夹里「中国三单 / 越南四单」PL箱单 的 Packages 对不对得上。
    · 唛头内容会不会太长、一页放不下。
    前两条不通过就没法生成，必须改；后两条只提示，你确认就行。

五、★ 内容太长、一页放不下会怎样
    一个唛头固定 11 行、18pt 字，在母本的页边距下大约能放 24 行
    —— 也就是每行文字「折行」之后加起来不能超过 24 行。

    程序没有 Word 的排版引擎，只能按字体宽度「估」。估出来超了，它会告诉你：
      · 哪一票、哪个唛头（编号）、源表第几行
      · 超了几行
      · 哪一段最长、大概删几个字符就能收回来

    这时你有两个选择：
      (1) 按它点名的段落精简（最快）。通常是收货人地址太长，
          或者地址里塞了多余的空格、换行。
      (2) 不管它，点「继续生成」。那个唛头会被拆到两页上；
          因为它后面跟着硬分页符，后面的唛头会整体错位、页码也跟着变。
          少量几张准备手工贴的话可以这么干。

    想让程序「超了就干脆不让生成」：把 mark_config.json 里
    overflow_check.on_overflow 从 "warn" 改成 "error"。

六、命令行用法（给批处理 / 脚本用，跟图形界面是同一个 exe）
    唛头生成器.exe                            打开图形界面（双击这个）
    唛头生成器.exe -i 模板表.xlsx -o 输出.docx  直接生成
    唛头生成器.exe --dry-run                  只检查不写文件
    唛头生成器.exe --no-pl-check              跳过件数与三单/四单 PL 的核对
    唛头生成器.exe --self-test                跑「测试用例」里的全部用例
    唛头生成器.exe --check 模板表.xlsx         只检查，打印问题清单
    唛头生成器.exe --build 模板表.xlsx --yes   走图形界面同一条流程生成
    唛头生成器.exe -V                         看版本

    ★ 命令行模式的输出会同时写进「界面自检结果.txt」，因为这是个
      窗口程序、默认没有黑框。bat 里已经帮你 type 出来了。

七、目录说明
    唛头生成器.exe            主程序（双击这个，不用装 Python）
    _internal\\                程序运行必需的组件 —— 别删、别改、别挪
    mark_config.json          配置：字段别名、大写范围、PL 核对、溢出检查
    唛头母本.docx              版式母本（字体/字号/页边距都取自它，别删别改名）
    唛头模板表.xlsx            示例 / 空白输入表（带「使用说明」sheet）
    测试用例\\                 全部用例（含期望报错的负面用例）
    使用说明.txt               本文件
    1-生成唛头.bat             打开图形界面
    2-只校验不生成.bat         检查一个模板表（可以把表拖到它上面）
    3-自检.bat                 跑全部用例
    4-界面自检.bat             图形界面自检，结果写进「界面自检结果.txt」
    5-拖模板到这里生成.bat     不用开界面，把模板表拖到它上面直接出唛头

    ★ exe 和 _internal\\ 必须待在一起，整个文件夹一起复制。
      别只把 exe 单独拷走 —— 那样起不来（程序会明确告诉你缺什么）。

八、出问题了？
    先看程序里的「常见问题」页签，或者看信息区里给的改法。
    程序报错时不会只甩一句「失败」，都会说清楚「第几行、哪个单元格、改成什么」。
    万一真崩了，程序会在同目录留下「唛头生成器_出错日志.txt」，把它发回来即可。

九、版本
    v2.0.0   2026-09-24   加入图形界面（含教程与常见问题）；错误信息带单元格与改法；
                          溢出提示点明是哪一段、要删几个字符
    v1.0.0   2026-09-24   首个命令行发行版
"""

BAT = {
    "1-生成唛头.bat":
        '@echo off\r\nchcp 65001 >nul\r\ncd /d "%~dp0"\r\n'
        'start "" "{exe}"\r\n',

    "2-只校验不生成.bat":
        '@echo off\r\nchcp 65001 >nul\r\ncd /d "%~dp0"\r\n'
        'set "TPL=%~1"\r\n'
        'if "%TPL%"=="" set "TPL=唛头模板表.xlsx"\r\n'
        'del /q "界面自检结果.txt" 2>nul\r\n'
        '"{exe}" --check "%TPL%"\r\n'
        'echo.\r\n'
        'if exist "界面自检结果.txt" (type "界面自检结果.txt") else (echo 没有拿到结果)\r\n'
        'echo.\r\npause\r\n',

    "3-自检.bat":
        '@echo off\r\nchcp 65001 >nul\r\ncd /d "%~dp0"\r\n'
        'del /q "界面自检结果.txt" 2>nul\r\n'
        '"{exe}" --self-test\r\n'
        'echo.\r\n'
        'if exist "界面自检结果.txt" (type "界面自检结果.txt") else (echo 没有拿到结果)\r\n'
        'echo.\r\npause\r\n',

    "4-界面自检.bat":
        '@echo off\r\nchcp 65001 >nul\r\ncd /d "%~dp0"\r\n'
        'del /q "界面自检结果.txt" 2>nul\r\n'
        '"{exe}" --selftest-gui\r\n'
        'echo.\r\n'
        'if exist "界面自检结果.txt" (start "" "界面自检结果.txt") '
        'else (echo 没拿到结果，界面组件可能没起来)\r\n'
        'pause\r\n',

    "5-拖模板到这里生成.bat":
        '@echo off\r\nchcp 65001 >nul\r\ncd /d "%~dp0"\r\n'
        'if "%~1"=="" (\r\n'
        '  echo 把「唛头模板表.xlsx」拖到本文件上面再松手。\r\n'
        '  echo 输出会写在那个模板表所在的文件夹里。\r\n'
        '  echo.\r\n  pause\r\n  exit /b\r\n)\r\n'
        'del /q "界面自检结果.txt" 2>nul\r\n'
        '"{exe}" --build "%~1" --yes\r\n'
        'echo.\r\n'
        'if exist "界面自检结果.txt" (type "界面自检结果.txt") else (echo 没有拿到结果)\r\n'
        'echo.\r\npause\r\n',
}


def main():
    if not os.path.isdir(APP):
        raise SystemExit("缺 onedir 产物：%s\n先跑一次 PyInstaller --onedir。" % APP)

    if os.path.isdir(REL):
        shutil.rmtree(REL)
    os.makedirs(REL)

    # 整个 onedir 目录（exe + _internal）拷进去
    for name in os.listdir(APP):
        s, d = os.path.join(APP, name), os.path.join(REL, name)
        if os.path.isdir(s):
            shutil.copytree(s, d)
        else:
            shutil.copy2(s, d)

    # 母本 / 模板表（打包进 exe 的那份做兜底，旁边这份是给用户改的）
    for f in ("唛头母本.docx", "唛头模板表.xlsx"):
        shutil.copy2(os.path.join(HERE, f), os.path.join(REL, f))

    # 配置：pl_check.dir = [".", ".."]，让「工具文件夹放项目目录里」也能核对到 PL
    with open(os.path.join(HERE, "mark_config.json"), encoding="utf-8") as f:
        cfg = json.load(f)
    cfg["_说明"] = ("唛头生成器 v%s 配置（发行版）。"
                    "相对路径一律相对于本文件所在目录。" % VER)
    cfg.setdefault("pl_check", {})["dir"] = [".", ".."]
    with open(os.path.join(REL, "mark_config.json"), "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

    # 说明 + 批处理
    with open(os.path.join(REL, "使用说明.txt"), "w", encoding="utf-8-sig") as f:
        f.write(README.format(ver=VER))
    for name, body in BAT.items():
        with open(os.path.join(REL, name), "w", encoding="utf-8", newline="") as f:
            f.write(body.format(exe=EXE_NAME))

    # 测试用例
    src_cases = os.path.join(HERE, "测试用例")
    dst_cases = os.path.join(REL, "测试用例")
    os.makedirs(dst_cases)
    n = 0
    for f in sorted(os.listdir(src_cases)):
        if f.lower().endswith((".xlsx", ".py")):
            shutil.copy2(os.path.join(src_cases, f), os.path.join(dst_cases, f))
            n += 1

    # 清掉测试留下的临时文件（别打进发行包）
    for junk in ("界面自检结果.txt", "唛头生成器_出错日志.txt", "唛头DEMO-20260101-1.docx"):
        p = os.path.join(REL, junk)
        if os.path.exists(p):
            os.remove(p)

    # zip
    zip_base = os.path.join(ROOT, "发行版", "唛头生成器_v%s" % VER)
    if os.path.exists(zip_base + ".zip"):
        os.remove(zip_base + ".zip")
    shutil.make_archive(zip_base, "zip", os.path.dirname(REL), os.path.basename(REL))

    files = sum(len(fs) for _, _, fs in os.walk(REL))
    size = sum(os.path.getsize(os.path.join(r, f))
               for r, _, fs in os.walk(REL) for f in fs)
    print("发行目录：%s" % REL)
    print("           %d 个文件 / %.1f MB（含测试用例 %d 个）" % (files, size / 1048576, n))
    print("压缩包  ：%s.zip" % zip_base)


if __name__ == "__main__":
    main()
