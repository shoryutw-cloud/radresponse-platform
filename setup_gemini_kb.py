"""
一次性執行：把 3 個 .docx 知識庫文件萃取成 gemini_kb.json
前端直接讀這個 JSON，不需要 Corpus API。
用法：python setup_gemini_kb.py
（不需要 API Key）
"""
import json, sys
from pathlib import Path
from docx import Document as DocxDoc

DOCS = [
    ("1051107應變技術隊輻射彈事件現場應變作業程序書(正式版)-0630改封面.docx",         "應變程序書（正式版）"),
    ("NSER-NSS-012 應變技術隊輻射彈事件現場應變作業程序書-修訂3版_加RDD(全黑).docx", "應變程序書（修訂3版）"),
    ("輻射彈事件(已爆彈)偵測區域作業導則0730核定版.docx",                             "偵測區域作業導則"),
]

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

output = {"documents": [], "totalChars": 0}
all_text = ""

for fname, display in DOCS:
    fpath = Path(fname)
    if not fpath.exists():
        print(f"找不到 {fname}，跳過"); continue
    blocks = extract_text(fpath)
    chunks = make_chunks(blocks)
    doc_text = "\n\n".join(chunks)
    output["documents"].append({
        "name":   display,
        "chunks": chunks,
        "text":   doc_text
    })
    all_text += f"\n\n===【{display}】===\n{doc_text}"
    chars = sum(len(c) for c in chunks)
    output["totalChars"] += chars
    print(f"  {display}：{len(chunks)} chunks，{chars} 字")

output["combinedText"] = all_text.strip()
out_path = Path("gemini_kb.json")
out_path.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"\n完成！總計 {output['totalChars']} 字")
print(f"已存到 gemini_kb.json（{out_path.stat().st_size//1024} KB）")
print("請重新部署（git push），前端就能讀到新知識庫。")
