# fanjian-tw 繁簡轉換（台灣用字）

把簡體中文文字檔轉成**繁體中文、台灣慣用字**的離線小工具。適合要把簡體小說、文章整理成繁體書稿的人。

- **全程離線**：檔案不會上傳到任何地方。
- **兩種用法**：Python 指令，或直接用瀏覽器開一個網頁檔，不用安裝任何東西。
- **台灣用字**：以 OpenCC 詞表為基礎，再加上依台灣出版慣用字調整的詞組（例如 身份→身分、想象→想像）。
- **輸出 Big5（CP950）**：方便交給只吃 Big5 的舊排版軟體。Big5 裡沒有的字會換成 ●，並在報告裡列出來，讓你回頭補字。
- **順手整理排版**：引號換成「」『』，半形 `? ! ~` 換成全形，章名格式統一，並偵測疑似作者閒話（例如「ps：」「求月票」），只列在報告裡，不會自動刪除。

> 轉換是機器做的，一定還要人工校對。同一個簡體字在不同語境可能對應不同繁體字（例如「里」和「裡」），報告會幫你標出要注意的地方。

## 用法一：網頁版（最簡單）

1. 下載 `fanjian.html`。
2. 用瀏覽器（Chrome、Edge、Firefox）開啟。
3. 把 TXT 或 DOCX 檔拖進虛線框，或點一下虛線框選檔。可以一次拖多個檔。
4. 轉好的檔案會自動下載。瀏覽器可能會問你是否允許下載多個檔案，請按允許。

舊式 `.doc` 檔請先用 Word 另存成「Unicode 純文字」或 `.docx` 再轉。

## 用法二：Python 版

需要 Python 3.7 以上（只用標準函式庫，不必另外安裝套件）。

```
python -I fanjian.py 你的檔案.txt
python -I fanjian.py 整個資料夾
```

Windows 也可以把檔案或資料夾直接拖到 `convert.bat` 上。

- 支援 TXT、DOCX；舊式 DOC 只能在 Windows 上、而且電腦有安裝 Word 才能轉（會以唯讀、關閉巨集的方式開啟）。
- 轉好的檔案放在原檔旁邊，不會覆蓋原檔（已有同名檔時會另取名字）。
- 同時會產生一份 UTF-8 報告，列出缺字、詞庫命中次數與疑似作者閒話。

## 輸出格式會做的整理

除了轉字，輸出檔還會順手整理成書稿格式，用之前請先知道：

- **章名統一**：「第一章 xxx」「第12回xxx」這類章名，章號和章名之間會統一成一個全形空白，例如「第一章　xxx」；整行只有「1384、章名」這種編號標題，會改成「第1384章　章名」。
- **段落**：每行前後的空白（含全形空白）會刪掉；段與段之間一律空一行，連續好幾行空行會壓成一個。
- **段落記號**：段與段之間那一行空行，輸出時會寫成 `~\`。這是部分直排排版軟體認得的分段記號；你的軟體不需要的話，用取代功能把 `~\` 全部刪掉即可。
- **換行與編碼**：輸出是 Big5（CP950）、Windows 換行（CRLF）。

## 自訂詞庫

覺得哪個詞轉得不對，可以自己加規則：

- `lexicon/s2t.txt`：簡轉繁時**優先**使用的詞組。
- `lexicon/post.txt`：轉成繁體**之後**再依序替換的規則。

格式是一行一條 `原文@替換文字`，檔案用 UTF-8 存，`#` 開頭的行是註解。

改完詞庫後執行 `python -I build_web.py`，網頁版 `fanjian.html` 就會重新產生，兩個版本用的規則保持一致。

## 測試

```
python -I -m unittest discover tests
python -I build_web.py --test tests/sample.txt
```

第二行會用 Node.js 跑網頁版的轉換邏輯，逐位元組比對 Python 版和網頁版的輸出是否完全相同。

## 授權

- 本專案程式碼：MIT License，見 `LICENSE`。
- 內含 OpenCC 1.1.9 詞表資料（Apache License 2.0），見 `NOTICE` 與 `opencc/LICENSE.txt`。

本工具由天機閣維護，維護為盡力而為。歡迎開 issue 回報轉錯的詞。

## English

**fanjian-tw** converts Simplified Chinese text into Traditional Chinese using Taiwan conventions. It runs fully offline, either as a Python script (standard library only) or as a single self-contained HTML page you open in a browser. It accepts TXT and DOCX, outputs CP950 (Big5) text plus a UTF-8 report listing characters Big5 cannot encode (replaced with ●), lexicon hits, and likely author's-note lines. Built on OpenCC 1.1.9 dictionary data (Apache-2.0) with extra Taiwan-oriented phrase rules you can edit in `lexicon/`. Machine conversion still needs human proofreading. Code is MIT licensed; maintained on a best-effort basis.
