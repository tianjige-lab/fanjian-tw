# 產生「fanjian.html」（單一檔、離線可用，較舊的 Chrome／Firefox 也能開）
# 用法：python -X utf8 -I build_web.py              → 重新產生網頁版（改了詞庫或 web\core.js 之後跑）
#       python -X utf8 -I build_web.py --test 樣本…  → 產生後，用 Node 拿同一批樣本比對網頁版和 Python 版的輸出
# 網頁版內嵌：OpenCC 資料、兩個詞庫、CP950 對照表、web\core.js。改規則後要重跑本檔，再把新的 html 給使用者。
import json, re, sys, shutil, subprocess, tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = HERE / "fanjian.html"

def version():
    m = re.search(r'^VERSION = "(.*)"', (HERE / "fanjian.py").read_text(encoding="utf-8"), re.M)
    return m.group(1)

def cp950_map():
    chars, codes = [], []
    for cp in range(0x80, 0x10000):
        if 0xD800 <= cp <= 0xDFFF: continue
        try: b = chr(cp).encode("cp950")
        except UnicodeEncodeError: continue
        if len(b) != 2: raise SystemExit(f"CP950 出現非雙位元組對應：U+{cp:04X}")
        chars.append(chr(cp)); codes.append(b.hex())
    return "".join(chars), "".join(codes)

def build():
    oc, lex = HERE / "opencc", HERE / "lexicon"
    chars, codes = cp950_map()
    data = {
        "STPhrases": (oc / "STPhrases.txt").read_text(encoding="utf-8"),
        "STCharacters": (oc / "STCharacters.txt").read_text(encoding="utf-8"),
        "TWVariants": (oc / "TWVariants.txt").read_text(encoding="utf-8"),
        "user": (lex / "s2t.txt").read_text(encoding="utf-8-sig"),
        "post": (lex / "post.txt").read_text(encoding="utf-8-sig"),
        "cp950chars": chars, "cp950codes": codes,
    }
    js = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    core = (HERE / "web" / "core.js").read_text(encoding="utf-8")
    html = (HERE / "web" / "template.html").read_text(encoding="utf-8")
    html = html.replace("__VERSION__", version()).replace("__CORE__", core).replace("__DATA__", js)
    OUT.write_text(html, encoding="utf-8", newline="\n")
    print(f"✓ {OUT.name}　{OUT.stat().st_size // 1024} KB")

# ---------- 比對（--test） ----------
NODE_HARNESS = r"""
const fs = require("fs"), vm = require("vm");
const [html, src, outDir, ver] = process.argv.slice(2);
const h = fs.readFileSync(html, "utf8");
const pick = id => h.match(new RegExp('<script id="' + id + '"[^>]*>([\\s\\S]*?)</script>'))[1];
vm.runInThisContext(pick("fjcore"));
const ctx = FJ.build(JSON.parse(pick("fjdata")));
const name = require("path").basename(src);
const r = FJ.convert(name, new Uint8Array(fs.readFileSync(src)), ctx, ver);
fs.writeFileSync(outDir + "/" + r.outName, r.b5);
fs.writeFileSync(outDir + "/" + r.reportName, "﻿" + r.report);
"""

def test(args):
    srcs = []
    for a in args:
        p = Path(a)
        srcs += sorted(q for q in p.iterdir() if q.suffix.lower() in (".txt", ".docx")) if p.is_dir() else [p]
    srcs = [p for p in srcs if not (p.stem.endswith("-b5") or p.stem.endswith("-報告") or "-b5-" in p.stem)]
    work = Path(tempfile.mkdtemp())
    harness = work / "h.js"; harness.write_text(NODE_HARNESS, encoding="utf-8")
    bad = 0
    for p in srcs:
        dp, dj = work / "py", work / "js"
        for d in (dp, dj): shutil.rmtree(d, ignore_errors=True); d.mkdir()
        shutil.copy2(p, dp / p.name)
        subprocess.run([sys.executable, "-X", "utf8", "-I", str(HERE / "fanjian.py"), str(dp / p.name)], check=True, capture_output=True)
        subprocess.run(["node", str(harness), str(OUT), str(p), str(dj), version() + "（網頁版）"], check=True)
        stem = p.stem
        b_py, b_js = (dp / f"{stem}-b5.txt").read_bytes(), (dj / f"{stem}-b5.txt").read_bytes()
        r_py = (dp / f"{stem}-b5-報告.txt").read_text(encoding="utf-8-sig").split("\n")[1:]
        r_js = (dj / f"{stem}-b5-報告.txt").read_text(encoding="utf-8-sig").split("\n")[1:]
        ok_b, ok_r = b_py == b_js, r_py == r_js
        if not (ok_b and ok_r):
            bad += 1
            msg = []
            if not ok_b:
                k = next((i for i in range(min(len(b_py), len(b_js))) if b_py[i] != b_js[i]), min(len(b_py), len(b_js)))
                msg.append(f"b5 不同（py {len(b_py)} / js {len(b_js)} 位元組，第 {k} 位元組起）")
            if not ok_r:
                k = next((i for i in range(min(len(r_py), len(r_js))) if r_py[i] != r_js[i]), None)
                msg.append(f"報告不同：第 {k + 2 if k is not None else '?'} 行 py「{r_py[k] if k is not None else ''}」js「{r_js[k] if k is not None else ''}」")
            print(f"✗ {p.name}：" + "；".join(msg))
        else:
            print(f"✓ {p.name}：b5 {len(b_py)} 位元組、報告完全相同")
    shutil.rmtree(work, ignore_errors=True)
    print(f"共 {len(srcs)} 份，不同 {bad} 份")
    return bad

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    build()
    if sys.argv[1:2] == ["--test"]:
        sys.exit(1 if test(sys.argv[2:]) else 0)
