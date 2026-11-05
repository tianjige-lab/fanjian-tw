// 繁簡轉換網頁版核心（與 fanjian.py 同一套規則；改規則時兩邊要一起改，build_web.py 會用樣本比對兩邊輸出）
(function (g) {
  "use strict";
  var M = "~\\";

  // ---------- 詞庫 ----------
  function loadTab(text) {
    var d = new Map();
    text.split("\n").forEach(function (line) {
      line = line.replace(/\r$/, "");
      if (!line || line.charAt(0) === "#") return;
      var k = line.indexOf("\t");
      d.set(line.slice(0, k), line.slice(k + 1).split(" ")[0]);   // 一對多取第一個（OpenCC 慣例）
    });
    return d;
  }

  function loadAt(text, name) {
    var rows = [];
    text.replace(/^﻿/, "").split(/\r?\n/).forEach(function (s, i) {
      if (!s.trim() || s.replace(/^\s+/, "").charAt(0) === "#") return;
      var k = s.indexOf("@");
      if (k < 0) throw new Error("詞庫格式錯：" + name + " 第 " + (i + 1) + " 行沒有 @：" + s);
      if (k === 0) throw new Error("詞庫格式錯：" + name + " 第 " + (i + 1) + " 行 @ 左邊是空的");
      rows.push([s.slice(0, k), s.slice(k + 1)]);
    });
    return rows;
  }

  function build(D) {
    var table = loadTab(D.STCharacters);
    loadTab(D.STPhrases).forEach(function (v, k) { table.set(k, v); });
    var user = loadAt(D.user, "s2t.txt");
    user.forEach(function (r) { table.set(r[0], r[1]); });   // 自訂詞庫最優先
    var lens = new Map();
    table.forEach(function (v, k) {
      var c = k.charAt(0), s = lens.get(c);
      if (!s) lens.set(c, s = []);
      if (s.indexOf(k.length) < 0) s.push(k.length);
    });
    lens.forEach(function (s) { s.sort(function (a, b) { return b - a; }); });
    return { table: table, lens: lens, tw: loadTab(D.TWVariants), user: user,
             post: loadAt(D.post, "post.txt"), cp950: buildCp950(D.cp950chars, D.cp950codes) };
  }

  function s2t(text, ctx, hits) {
    var out = [], i = 0, n = text.length, table = ctx.table, lens = ctx.lens;
    while (i < n) {
      var ls = lens.get(text.charAt(i)), hit = false;
      if (ls) {
        for (var j = 0; j < ls.length; j++) {
          var w = text.substr(i, ls[j]);
          if (w.length < ls[j]) continue;
          var v = table.get(w);
          if (v !== undefined) {
            if (hits.has(w)) hits.set(w, hits.get(w) + 1);
            out.push(v); i += ls[j]; hit = true; break;
          }
        }
      }
      if (!hit) { out.push(text.charAt(i)); i += 1; }
    }
    var s = out.join(""), tw = ctx.tw, r = [];
    for (var k = 0; k < s.length; k++) { var c = s.charAt(k); r.push(tw.has(c) ? tw.get(c) : c); }   // 台灣字形（裏→裡 等）
    return r.join("");
  }

  // ---------- 讀檔 ----------
  function decodeTxt(b) {
    if (b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) return [new TextDecoder("utf-8").decode(b.subarray(3)), "UTF-8（有 BOM）"];
    if (b[0] === 0xff && b[1] === 0xfe) return [new TextDecoder("utf-16le").decode(b.subarray(2)), "UTF-16"];
    if (b[0] === 0xfe && b[1] === 0xff) return [new TextDecoder("utf-16be").decode(b.subarray(2)), "UTF-16"];
    var tries = [["utf-8", "UTF-8"], ["gb18030", "GB18030／GBK"], ["big5", "Big5"]];
    for (var i = 0; i < tries.length; i++) {
      try { return [new TextDecoder(tries[i][0], { fatal: true }).decode(b), tries[i][1]]; } catch (e) { }
    }
    throw new Error("看不懂的編碼");
  }

  var LBASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
  var LEXT = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
  var DBASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
  var DEXT = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
  var CLORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

  // 解 docx（zip）裡的 deflate 資料；照 zlib 的 puff.c 寫法
  function inflateRaw(src) {
    var pos = 0, bb = 0, bc = 0, out = new Uint8Array(src.length * 4 + 1024), olen = 0;
    function need(n) {
      if (olen + n <= out.length) return;
      var o = new Uint8Array(Math.max(out.length * 2, olen + n)); o.set(out.subarray(0, olen)); out = o;
    }
    function bits(n) {
      while (bc < n) {
        if (pos >= src.length) throw new Error("Word 檔壓縮資料不完整");
        bb |= src[pos++] << bc; bc += 8;
      }
      var v = bb & ((1 << n) - 1); bb >>>= n; bc -= n; return v;
    }
    function huff(lengths) {
      var counts = new Uint16Array(16), offs = new Uint16Array(16), syms = new Uint16Array(lengths.length), i;
      for (i = 0; i < lengths.length; i++) counts[lengths[i]]++;
      counts[0] = 0;
      for (i = 1; i < 16; i++) offs[i] = offs[i - 1] + counts[i - 1];
      for (i = 0; i < lengths.length; i++) if (lengths[i]) syms[offs[lengths[i]]++] = i;
      return { counts: counts, syms: syms };
    }
    function decode(h) {
      var code = 0, first = 0, index = 0;
      for (var len = 1; len < 16; len++) {
        code |= bits(1);
        var c = h.counts[len];
        if (code - first < c) return h.syms[index + code - first];
        index += c; first += c; first <<= 1; code <<= 1;
      }
      throw new Error("Word 檔壓縮資料錯誤");
    }
    var fixL = null, fixD = null, last;
    do {
      last = bits(1);
      var type = bits(2), lh, dh, i;
      if (type === 0) {
        bb = 0; bc = 0;
        var len = src[pos] | (src[pos + 1] << 8); pos += 4;
        need(len); out.set(src.subarray(pos, pos + len), olen); olen += len; pos += len;
        continue;
      }
      if (type === 1) {
        if (!fixL) {
          var l = [];
          for (i = 0; i < 288; i++) l.push(i < 144 ? 8 : i < 256 ? 9 : i < 280 ? 7 : 8);
          var d = []; for (i = 0; i < 30; i++) d.push(5);
          fixL = huff(l); fixD = huff(d);
        }
        lh = fixL; dh = fixD;
      } else if (type === 2) {
        var hl = bits(5) + 257, hd = bits(5) + 1, hc = bits(4) + 4, cl = [];
        for (i = 0; i < 19; i++) cl.push(0);
        for (i = 0; i < hc; i++) cl[CLORDER[i]] = bits(3);
        var ch = huff(cl), ls = [];
        while (ls.length < hl + hd) {
          var s = decode(ch);
          if (s < 16) { ls.push(s); continue; }
          var rep, val = 0;
          if (s === 16) { val = ls[ls.length - 1]; rep = 3 + bits(2); }
          else if (s === 17) rep = 3 + bits(3);
          else rep = 11 + bits(7);
          while (rep--) ls.push(val);
        }
        lh = huff(ls.slice(0, hl)); dh = huff(ls.slice(hl));
      } else throw new Error("Word 檔壓縮格式錯誤");
      for (; ;) {
        var sym = decode(lh);
        if (sym < 256) { need(1); out[olen++] = sym; }
        else if (sym === 256) break;
        else {
          var k = sym - 257, n = LBASE[k] + bits(LEXT[k]), ds = decode(dh), dist = DBASE[ds] + bits(DEXT[ds]);
          need(n);
          for (var q = 0; q < n; q++) { out[olen] = out[olen - dist]; olen++; }
        }
      }
    } while (!last);
    return out.subarray(0, olen);
  }

  function unzipEntry(b, want) {
    var dv = new DataView(b.buffer, b.byteOffset, b.byteLength), e = -1;
    for (var i = b.length - 22; i >= Math.max(0, b.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { e = i; break; }
    if (e < 0) throw new Error("Word 檔損壞（找不到目錄）");
    var cnt = dv.getUint16(e + 10, true), p = dv.getUint32(e + 16, true), td = new TextDecoder("utf-8");
    for (var k = 0; k < cnt; k++) {
      var method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true),
        nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true),
        off = dv.getUint32(p + 42, true), name = td.decode(b.subarray(p + 46, p + 46 + nl));
      if (name === want) {
        var start = off + 30 + dv.getUint16(off + 26, true) + dv.getUint16(off + 28, true), raw = b.subarray(start, start + csize);
        if (method === 0) return raw;
        if (method === 8) return inflateRaw(raw);
        throw new Error("Word 檔用了不支援的壓縮方式");
      }
      p += 46 + nl + el + cl;
    }
    throw new Error("這不是 Word 檔（裡面沒有 " + want + "）");
  }

  function readDocx(b) {
    var x = new TextDecoder("utf-8").decode(unzipEntry(b, "word/document.xml"));
    var paras = (x.match(/<w:p[ >][\s\S]*?<\/w:p>|<w:p\/>/g) || []).map(function (p) {
      var t = p.replace(/<w:t[^>]*>([^<]*)<\/w:t>/g, "$1").replace(/<w:br[^>]*\/>/g, "\n").replace(/<[^>]+>/g, "");
      return t.split("&amp;").join("&").split("&lt;").join("<").split("&gt;").join(">")
        .split("&quot;").join('"').split("&apos;").join("'");
    });
    return [paras.join("\n"), "Word（docx）"];
  }

  function readAny(name, b) {
    var ext = (name.match(/\.[^.]*$/) || [""])[0].toLowerCase(), isZip = b[0] === 0x50 && b[1] === 0x4b;
    if (ext === ".docx" || (ext === ".doc" && isZip)) return readDocx(b);
    if (ext === ".doc") throw new Error("這是舊版 Word 檔（.doc），網頁讀不了。請用 Word 打開它，「另存新檔」選「純文字」，編碼選「Unicode」，再把存出來的 txt 拖進來。");
    return decodeTxt(b);
  }

  // ---------- 排版 ----------
  var CH = /^(第[0-9０-９零〇一二三四五六七八九十百千兩两]+[章節节回])[\s　]*(.*)$/u;
  var CH2 = /^0*([1-9][0-9]{0,4})[、.．]\s*(\S.{0,38})$/u;   // u：和 Python 一樣按「字」數，不按 UTF-16 碼元   // 例如「1384、章名」
  var CHAT = /^(ps|PS|Ps|pS)[：:]|求(訂閱|月票|推薦票|收藏|保底)|^[（(]?第[一二三四五六七八九十]更/;
  var PUNCT = { "“": "「", "”": "」", "‘": "『", "’": "』", "?": "？", "!": "！", "~": "～" };
  var STRIP = /^[ \t　 ﻿]+|[ \t　 ﻿]+$/g;

  function layout(text, stats) {
    var lines = text.split("\r\n").join("\n").split("\r").join("\n").split("\n").map(function (raw) {
      var t = raw.replace(STRIP, "").replace(/[“”‘’?!~]/g, function (c) { return PUNCT[c]; });
      var m = CH.exec(t), m2 = m ? null : CH2.exec(t);
      if (m2 && "。！？」…".indexOf(m2[2].slice(-1)) >= 0) m2 = null;
      if (m) { stats.ch++; t = m[1] + (m[2].trim() ? "　" + m[2].trim() : ""); }
      else if (m2) { stats.ch++; t = "第" + m2[1] + "章　" + m2[2].trim(); }
      if (CHAT.test(t)) stats.chat.push(Array.from(t).slice(0, 40).join(""));
      return t;
    });
    var out = [];
    lines.forEach(function (t) { out.push(t, ""); });
    var s = out.join("\n").replace(/^\n+|\n+$/g, "");
    s = s.replace(/\n{4,}/g, "\n\n\n").replace(/\n\n\n/g, "\n\n");
    var res = s.split("\n").map(function (l) { return l === "" ? M : l; });
    while (res.length && res[res.length - 1] === M) res.pop();
    return res;
  }

  // ---------- Big5 ----------
  function buildCp950(chars, codes) {
    var m = new Map(), a = Array.from(chars);
    for (var i = 0; i < a.length; i++) m.set(a[i], parseInt(codes.substr(i * 4, 4), 16));
    return m;
  }

  function toBig5(lines, cp, missing) {
    var parts = [], total = 0;
    lines.forEach(function (l, idx) {
      var buf = [];
      for (var ch of l) {
        var c = ch.codePointAt(0);
        if (c < 0x80) buf.push(c);
        else if (cp.has(ch)) { var v = cp.get(ch); buf.push(v >> 8, v & 0xff); }
        else {
          if (!missing.has(ch)) missing.set(ch, []);
          missing.get(ch).push(idx + 1);
          var dot = cp.get("●"); buf.push(dot >> 8, dot & 0xff);
        }
      }
      buf.push(13, 10);
      parts.push(buf); total += buf.length;
    });
    var out = new Uint8Array(total), o = 0;
    parts.forEach(function (p) { out.set(p, o); o += p.length; });
    return out;
  }

  function pad(n) { return (n < 10 ? "0" : "") + n; }

  // 回傳 { b5: Uint8Array, report: 字串, stem, missingKinds, chapters }
  function convert(name, bytes, ctx, version) {
    var r = readAny(name, bytes), text = r[0], enc = r[1];
    var hits = new Map(); ctx.user.forEach(function (u) { hits.set(u[0], 0); });
    var t = s2t(text, ctx, hits), nli = 0;
    // 內建：數字後的「里」是距離（數千里、十萬里、百里挑一），OpenCC 會被「數千」「十萬」先吃掉而轉成「裡」
    t = t.replace(/(?<=[0-9０-９數一二三四五六七八九十百千萬億幾兩])[裏裡](?![面頭邊])/g, function () { nli++; return "里"; });
    var posth = [["數字＋裡", "數字＋里（內建）", nli]];
    ctx.post.forEach(function (p) {
      var parts = t.split(p[0]), n = parts.length - 1;
      if (n) t = parts.join(p[1]);
      posth.push([p[0], p[1], n]);
    });
    var stats = { ch: 0, chat: [] }, lines = layout(t, stats), missing = new Map();
    var b5 = toBig5(lines, ctx.cp950, missing);
    var stem = name.replace(/\.[^.]*$/, ""), outName = stem + "-b5.txt", d = new Date();
    var han = (text.match(/[㐀-鿿]/g) || []).length, nMiss = 0;
    missing.forEach(function (v) { nMiss += v.length; });
    var rep = ["繁簡轉換 " + version + "　" + d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()) + " " + pad(d.getHours()) + ":" + pad(d.getMinutes()),
      "來源：" + name + "（" + enc + "）", "輸出：" + outName + "（Big5／CRLF／每段後 ~\\）",
      "漢字數：" + han + "　認出章節標題：" + stats.ch + " 個　輸出行數：" + lines.length, ""];
    rep.push("■ Big5 缺字（已換成 ●，請使用者造字或改字）：" + nMiss + " 處");
    Array.from(missing.entries()).sort(function (a, b) { return b[1].length - a[1].length; }).forEach(function (kv) {
      var nos = kv[1];
      rep.push("  " + kv[0] + "（U+" + kv[0].codePointAt(0).toString(16).toUpperCase().padStart(4, "0") + "）×" + nos.length +
        "　行號 " + nos.slice(0, 10).join(", ") + (nos.length > 10 ? " …" : ""));
    });
    rep.push("", "■ 疑似作者閒話（ps、求訂閱、第幾更；沒有刪，請使用者確認）：" + stats.chat.length + " 行");
    stats.chat.slice(0, 200).forEach(function (s) { rep.push("  " + s); });
    if (stats.chat.length > 200) rep.push("  …（以下省略）");
    var udict = new Map(ctx.user);
    rep.push("", "■ 自訂簡轉繁詞庫命中：");
    hits.forEach(function (n, a) { if (n) rep.push("  " + a + "→" + udict.get(a) + " ×" + n); });
    rep.push("", "■ 繁體後修正命中：");
    posth.forEach(function (p) { if (p[2]) rep.push("  " + p[0] + "→" + p[1] + " ×" + p[2]); });
    return { b5: b5, report: rep.join("\n") + "\n", outName: outName, reportName: stem + "-b5-報告.txt",
             missingKinds: missing.size, chapters: stats.ch };
  }

  g.FJ = { build: build, convert: convert };
})(typeof window !== "undefined" ? window : globalThis);
