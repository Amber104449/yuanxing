import sys
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document

doc_path = r'c:\Users\adone\OneDrive\桌面\AI练习\wendang\fangan\【协会反馈】绿色机场平台二期-建设方案V.020260608_修改版.docx'
doc = Document(doc_path)

changes_log = []

def find_para_by_contains(doc, text_fragment):
    for i, p in enumerate(doc.paragraphs):
        if text_fragment in p.text:
            return i
    return -1

def replace_para_text(doc, idx, new_text):
    if idx < 0 or idx >= len(doc.paragraphs):
        return False
    p = doc.paragraphs[idx]
    for run in p.runs:
        run.text = ''
    if p.runs:
        p.runs[0].text = new_text
    else:
        p.add_run(new_text)
    return True

# FIX: 核心理念中的"9 项扩展模块"
idx = find_para_by_contains(doc, '9 项扩展模块共同实现这一定位转变')
if idx >= 0:
    text = doc.paragraphs[idx].text
    new_text = text.replace('9 项扩展模块', '多项扩展模块')
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[FIX] 核心理念：9项→多项')
else:
    changes_log.append(f'[SKIP] 核心理念：未找到')

# Also check table 6 for "4.7 培训"
for ti, table in enumerate(doc.tables):
    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            if '4.7' in cell.text or '培训与认证' in cell.text:
                cell.text = cell.text.replace('4.7 培训与认证', '').replace('4.7培训', '')
                changes_log.append(f'[FIX] Table {ti} Row {ri}: 去掉4.7培训引用')

doc.save(doc_path)
print("最终修复完成！")
for log in changes_log:
    print(f"  {log}")
