"""
一次性執行：讀取 3 個 .docx 知識庫文件，建立 Gemini Corpus。
執行完後會產生 gemini_corpus.json，前端用這個來查詢。
用法：python setup_gemini_kb.py <你的GeminiKey>
"""
import json, sys, time, requests
from pathlib import Path
from docx import Document as DocxDoc

BASE = "https://generativelanguage.googleapis.com/v1beta"

DOCS = [
    ("1051107應變技術隊輻射彈事件現場應變作業程序書(正式版)-0630改封面.docx",         "應變程序書（正式版）"),
    ("NSER-NSS-012 應變技術隊輻射彈事件現場應變作業程序書-修訂3版_加RDD(全黑).docx", "應變程序書（修訂3版）"),
    ("輻射彈事件(已爆彈)偵測區域作業導則0730核定版.docx",                             "偵測區域作業導則"),
]

# ── 取得 Key（優先從命令列參數，否則互動輸入）──────────────────────────────
if len(sys.argv) > 1:
    KEY = sys.argv[1].strip()
else:
    KEY = input("請輸入 Gemini API Key：").strip()
if not KEY:
    print("未輸入 Key，結束。"); sys.exit(1)

def gem(method, path, **kw):
    url = f"{BASE}/{path}"
    r = getattr(requests, method)(url, params={"key": KEY}, **kw)
    if not r.ok:
        print(f"  !! API 錯誤 {r.status_code}: {r.text[:600]}")
        return {}
    try:
        return r.json()
    except Exception:
        print(f"  !! 無法解析回應：{r.text[:300]}")
        return {}

# ── 建立 Corpus ────────────────────────────────────────────────────────────
print("\n[1/4] 建立 Corpus …")
corpus_json = gem("post", "corpora", json={"displayName": "輻射應變知識庫"})
corpus_name = corpus_json.get("name", "")
if not corpus_name:
    print("建立 Corpus 失敗：", corpus_json); sys.exit(1)
print(f"  Corpus: {corpus_name}")

# ── 提取並分塊 ─────────────────────────────────────────────────────────────
def extract_text(path):
    doc = DocxDoc(path)
    blocks, buf = [], []
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            if buf: blocks.append(" ".join(buf)); buf = []
        else:
            buf.append(t)
    if buf: blocks.append(" ".join(buf))
    return blocks

def make_chunks(blocks, max_ch=1800):
    chunks, cur, cur_len = [], [], 0
    for b in blocks:
        if cur_len + len(b) > max_ch and cur:
            chunks.append("\n".join(cur)); cur, cur_len = [b], len(b)
        else:
            cur.append(b); cur_len += len(b)
    if cur: chunks.append("\n".join(cur))
    return chunks

# ── 上傳文件 ───────────────────────────────────────────────────────────────
for step, (fname, display) in enumerate(DOCS, 2):
    fpath = Path(fname)
    if not fpath.exists():
        print(f"  找不到 {fname}，跳過"); continue
    print(f"\n[{step}/4] 處理「{display}」…")
    blocks = extract_text(fpath)
    chunks = make_chunks(blocks)
    print(f"  {len(blocks)} 個段落 → {len(chunks)} 個 chunks")

    doc_json = gem("post", f"{corpus_name}/documents", json={"displayName": display})
    doc_name = doc_json.get("name", "")
    if not doc_name:
        print("  建立 Document 失敗：", doc_json); continue

    for i in range(0, len(chunks), 100):
        batch = chunks[i:i+100]
        payload = {"requests": [{"chunk": {"data": {"stringValue": c}}} for c in batch]}
        gem("post", f"{doc_name}/chunks:batchCreate", json=payload)
        print(f"  上傳 chunks {i+1}–{i+len(batch)} / {len(chunks)}")
        time.sleep(0.5)

# ── 儲存設定 ───────────────────────────────────────────────────────────────
cfg = {"corpusName": corpus_name, "displayName": "輻射應變知識庫"}
Path("gemini_corpus.json").write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"\n✅ 完成！Corpus 已建立。")
print(f"   名稱：{corpus_name}")
print(f"   設定已存到 gemini_corpus.json，請把 corpusName 填入網頁設定。")
