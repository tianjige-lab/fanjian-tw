# 繁簡轉換（簡體稿 →  Big5 TXT）
# 用法：把 .txt／.docx／.doc（或整個資料夾）拖到「convert.bat」上；或
#   python -X utf8 -I fanjian.py 檔案或資料夾 [...]
# 輸出：原檔旁 「原檔名-b5.txt」（CP950、CRLF、無 BOM、每段後一行 ~\）＋「原檔名-b5-報告.txt」（UTF-8）
# 轉換順序：讀檔 → 簡轉繁（自訂詞庫優先，再 OpenCC 詞組、單字、台灣字形）→ 繁體後修正 → 排版 → Big5 缺字處理
import os, re, sys, zipfile, subprocess, tempfile, datetime
from pathlib import Path

VERSION = "v1.1（2026-11-05）"
HERE = Path(__file__).resolve().parent
OC = HERE / "opencc"
LEX = HERE / "lexicon"
M = "~" + chr(92)

# ---------- 詞庫 ----------
def load_tab(p):
    d = {}
    for line in p.read_text(encoding="utf-8").splitlines():
        if not line or line.startswith("#"): continue
        k, v = line.split("\t", 1)
        d[k] = v.split(" ")[0]          # 一對多取第一個（OpenCC 慣例）
    return d

def load_at(p):
    rows = []
    for n, line in enumerate(p.read_text(encoding="utf-8-sig").splitlines(), 1):
        s = line.rstrip("\r\n")
        if not s.strip() or s.lstrip().startswith("#"): continue
        if "@" not in s: raise SystemExit(f"詞庫格式錯：{p.name} 第 {n} 行沒有 @：{s}")
        a, b = s.split("@", 1)
        if not a: raise SystemExit(f"詞庫格式錯：{p.name} 第 {n} 行 @ 左邊是空的")
        rows.append((a, b))
    return rows

def build():
    phr = load_tab(OC / "STPhrases.txt")
    chars = load_tab(OC / "STCharacters.txt")
    tw = load_tab(OC / "TWVariants.txt")
    user = load_at(LEX / "s2t.txt")
    table = dict(chars); table.update(phr)
    for a, b in user: table[a] = b      # 自訂詞庫最優先
    lens = {}
    for k in table:
        lens.setdefault(k[0], set()).add(len(k))
    lens = {c: sorted(v, reverse=True) for c, v in lens.items()}
    return table, lens, tw, user

def s2t(text, table, lens, tw, hits):
    out, i, n = [], 0, len(text)
    while i < n:
        c = text[i]
        for L in lens.get(c, ()):
            w = text[i:i + L]
            if L > 1 and len(w) < L: continue
            v = table.get(w)
            if v is not None:
                if w in hits: hits[w] += 1
                out.append(v); i += L; break
        else:
            out.append(c); i += 1
    s = "".join(out)
    return "".join(tw.get(ch, ch) for ch in s)   # 台灣字形（裏→裡 等）

# ---------- 讀檔 ----------
def read_txt(p):
    b = p.read_bytes()
    if b[:3] == b"\xef\xbb\xbf": return b[3:].decode("utf-8"), "UTF-8（有 BOM）"
    if b[:2] in (b"\xff\xfe", b"\xfe\xff"): return b.decode("utf-16"), "UTF-16"
    for enc, name in (("utf-8", "UTF-8"), ("gb18030", "GB18030／GBK"), ("cp950", "Big5")):
        try: return b.decode(enc), name
        except UnicodeDecodeError: pass
    raise SystemExit(f"看不懂的編碼：{p.name}")

def read_docx(p):
    x = zipfile.ZipFile(p).read("word/document.xml").decode("utf-8")
    paras = []
    for para in re.findall(r"<w:p[ >].*?</w:p>|<w:p/>", x, flags=re.S):
        t = re.sub(r"<w:t[^>]*>([^<]*)</w:t>", r"\1", para)
        t = re.sub(r"<w:br[^>]*/>", "\n", t)
        t = re.sub(r"<[^>]+>", "", t)
        paras.append(t.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
                     .replace("&quot;", '"').replace("&apos;", "'"))
    return "\n".join(paras), "Word（docx）"

def read_doc(p):
    # 舊式 .doc 交給本機 Word 唯讀開啟取文字，不存回原檔；AutomationSecurity=3 關閉巨集
    tmp = Path(tempfile.mkdtemp()) / "doc.txt"
    ps = (
        "$ErrorActionPreference='Stop';$w=New-Object -ComObject Word.Application;$w.Visible=$false;$w.AutomationSecurity=3;"
        "$w.DisplayAlerts=0;try{$d=$w.Documents.Open($env:SRC,$false,$true,$false);"
        "[IO.File]::WriteAllText($env:DST,$d.Content.Text,[Text.Encoding]::UTF8);$d.Close(0)}finally{$w.Quit()}"
    )
    env = dict(os.environ, SRC=str(p), DST=str(tmp))
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
                   env=env, check=True, capture_output=True)
    t = tmp.read_text(encoding="utf-8-sig").replace("\r", "\n")
    tmp.unlink(); tmp.parent.rmdir()
    return t, "Word（doc，經 Word 讀出）"

def read_any(p):
    ext = p.suffix.lower()
    if ext == ".docx": return read_docx(p)
    if ext == ".doc":
        if p.read_bytes()[:2] == b"PK": return read_docx(p)   # 副檔名寫 .doc 其實是 docx
        return read_doc(p)
    return read_txt(p)

# ---------- 排版 ----------
CH = re.compile(r"^(第[0-9０-９零〇一二三四五六七八九十百千兩两]+[章節节回])[\s　]*(.*)$")
CH2 = re.compile(r"^0*([1-9][0-9]{0,4})[、.．]\s*(\S.{0,38})$")   # 例如「1384、章名」
CHAT = re.compile(r"^(ps|PS|Ps|pS)[：:]|求(訂閱|月票|推薦票|收藏|保底)|^[（(]?第[一二三四五六七八九十]更")
PUNCT = str.maketrans({"“": "「", "”": "」", "‘": "『", "’": "』", "?": "？", "!": "！", "~": "～"})

def layout(text, stats):
    lines = []
    for raw in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        t = raw.strip(" \t　 ﻿").translate(PUNCT)
        m = CH.match(t)
        m2 = None if m else CH2.match(t)
        if m2 and m2.group(2)[-1] in "。！？」…":
            m2 = None
        if m:
            stats["章"] += 1
            t = m.group(1) + ("　" + m.group(2).strip() if m.group(2).strip() else "")
        elif m2:
            stats["章"] += 1
            t = f"第{m2.group(1)}章　{m2.group(2).strip()}"
        if CHAT.search(t):
            stats["閒話"].append(t[:40])
        lines.append(t)
    out = []
    for t in lines: out += [t, ""]
    s = "\n".join(out).strip("\n")
    s = re.sub(r"\n{4,}", "\n\n\n", s); s = re.sub(r"\n\n\n", "\n\n", s)
    res = [M if l == "" else l for l in s.split("\n")]
    while res and res[-1] == M: res.pop()
    return res

# ---------- Big5 ----------
def to_big5(lines, missing):
    out = []
    for no, l in enumerate(lines, 1):
        buf = []
        for ch in l:
            try:
                ch.encode("cp950"); buf.append(ch)
            except UnicodeEncodeError:
                missing.setdefault(ch, []).append(no); buf.append("●")
        out.append("".join(buf))
    return out

def free_name(p):
    if not p.exists(): return p
    k = 2
    while True:
        q = p.with_name(f"{p.stem}-{k}{p.suffix}")
        if not q.exists(): return q
        k += 1

def convert(p, ctx):
    table, lens, tw, user, post = ctx
    text, enc = read_any(p)
    hits = {a: 0 for a, _ in user}
    t = s2t(text, table, lens, tw, hits)
    # 內建：數字後的「里」是距離（數千里、十萬里、百里挑一），OpenCC 會被「數千」「十萬」先吃掉而轉成「裡」
    t, nli = re.subn(r"(?<=[0-9０-９數一二三四五六七八九十百千萬億幾兩])[裏裡](?![面頭邊])", "里", t)
    posth = [("數字＋裡", "數字＋里（內建）", nli)]
    for a, b in post:
        n = t.count(a)
        if n: t = t.replace(a, b)
        posth.append((a, b, n))
    stats = {"章": 0, "閒話": []}
    lines = layout(t, stats)
    missing = {}
    lines = to_big5(lines, missing)
    data = ("\r\n".join(lines) + "\r\n").encode("cp950", errors="strict")
    out = free_name(p.with_name(p.stem + "-b5.txt"))
    out.write_bytes(data)
    han = len(re.findall(r"[㐀-鿿]", text))
    rep = [f"繁簡轉換 {VERSION}　{datetime.datetime.now():%Y-%m-%d %H:%M}",
           f"來源：{p.name}（{enc}）", f"輸出：{out.name}（Big5／CRLF／每段後 ~\\）",
           f"漢字數：{han}　認出章節標題：{stats['章']} 個　輸出行數：{len(lines)}", ""]
    rep.append(f"■ Big5 缺字（已換成 ●，請使用者造字或改字）：{sum(len(v) for v in missing.values())} 處")
    for ch, nos in sorted(missing.items(), key=lambda kv: -len(kv[1])):
        rep.append(f"  {ch}（U+{ord(ch):04X}）×{len(nos)}　行號 {', '.join(map(str, nos[:10]))}{' …' if len(nos) > 10 else ''}")
    rep += ["", f"■ 疑似作者閒話（ps、求訂閱、第幾更；沒有刪，請使用者確認）：{len(stats['閒話'])} 行"]
    rep += [f"  {s}" for s in stats["閒話"][:200]] + (["  …（以下省略）"] if len(stats["閒話"]) > 200 else [])
    rep += ["", "■ 自訂簡轉繁詞庫命中："] + [f"  {a}→{dict(user)[a]} ×{n}" for a, n in hits.items() if n]
    rep += ["", "■ 繁體後修正命中："] + [f"  {a}→{b} ×{n}" for a, b, n in posth if n]
    rp = out.with_name(out.stem + "-報告.txt")
    rp.write_text("\n".join(rep) + "\n", encoding="utf-8-sig")
    print(f"✓ {p.name} → {out.name}　缺字 {len(missing)} 種　章 {stats['章']}")

def main(args):
    if not args:
        print("把要轉的 .txt／.docx／.doc 或資料夾拖到「convert.bat」上。"); return
    table, lens, tw, user = build()
    post = load_at(LEX / "post.txt")
    ctx = (table, lens, tw, user, post)
    files = []
    for a in args:
        p = Path(a)
        if p.is_dir():
            files += sorted(q for q in p.iterdir() if q.suffix.lower() in (".txt", ".docx", ".doc"))
        else:
            files.append(p)
    for p in files:
        if p.name.startswith("~$") or p.stem.endswith("-b5") or p.stem.endswith("-報告") or "-b5-" in p.stem:
            continue
        try:
            convert(p, ctx)
        except Exception as e:
            print(f"✗ {p.name}：{e}")

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv[1:])
