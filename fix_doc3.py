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

def delete_para(doc, idx):
    if idx < 0 or idx >= len(doc.paragraphs):
        return
    p = doc.paragraphs[idx]
    p._element.getparent().remove(p._element)

# ============================================================
# CHANGE 1: 删除所有"等保三级"相关内容
# ============================================================
# 需要删除包含"等保三级"的段落或修改相关文本

# 删除"阿里云生态部署 · 等保三级合规"
idx = find_para_by_contains(doc, '阿里云生态部署 · 等保三级合规')
if idx >= 0:
    replace_para_text(doc, idx, '阿里云生态部署')
    changes_log.append(f'[CHANGE 1-1] 删除：等保三级合规（封面）')

# 删除"等保三级合规要求"
idx = find_para_by_contains(doc, '等保三级合规要求')
if idx >= 0:
    text = doc.paragraphs[idx].text
    new_text = text.replace('等保三级合规要求', '相关合规要求')
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[CHANGE 1-2] 修改：等保三级合规要求 → 相关合规要求')

# 删除"等保三级合规"
idx = find_para_by_contains(doc, '落实等保三级合规')
if idx >= 0:
    text = doc.paragraphs[idx].text
    new_text = text.replace('落实等保三级合规', '落实相关合规')
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[CHANGE 1-3] 修改：落实等保三级合规 → 落实相关合规')

# 删除7.2 等保三级合规整个章节
idx = find_para_by_contains(doc, '7.2 等保三级合规')
if idx >= 0:
    # 删除该标题段落后继续删除下面的内容段落
    delete_para(doc, idx)
    changes_log.append(f'[CHANGE 1-4] 删除：7.2 等保三级合规章节')
    # 删除下一段"阿里云已具备等保三级基础设施层资质"
    idx_next = find_para_by_contains(doc, '阿里云已具备等保三级基础设施层资质')
    if idx_next >= 0:
        delete_para(doc, idx_next)
        changes_log.append(f'[CHANGE 1-4] 删除：阿里云已具备等保三级基础设施层资质')
    # 删除"物理与环境：依托阿里云数据中心已具备的等保三级资质"
    idx_next = find_para_by_contains(doc, '物理与环境：依托阿里云数据中心已具备的等保三级资质')
    if idx_next >= 0:
        delete_para(doc, idx_next)
        changes_log.append(f'[CHANGE 1-4] 删除：物理与环境等保三级资质段落')

# 修改"等保三级所需安全组件"
idx = find_para_by_contains(doc, '等保三级所需安全组件')
if idx >= 0:
    text = doc.paragraphs[idx].text
    new_text = text.replace('等保三级所需安全组件', '所需安全组件')
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[CHANGE 1-5] 修改：等保三级所需安全组件 → 所需安全组件')

# 修改"安全组与等保三级合规组件"
idx = find_para_by_contains(doc, '安全组与等保三级合规组件')
if idx >= 0:
    text = doc.paragraphs[idx].text
    new_text = text.replace('安全组与等保三级合规组件', '安全组与相关合规组件')
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[CHANGE 1-6] 修改：安全组与等保三级合规组件 → 安全组与相关合规组件')

# 删除"等保三级"在表格中的出现
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            if '等保三级' in cell.text:
                for para in cell.paragraphs:
                    if '等保三级' in para.text:
                        for run in para.runs:
                            if '等保三级' in run.text:
                                run.text = run.text.replace('等保三级', '')
                        changes_log.append(f'[CHANGE 1-7] 表格中删除：等保三级')

# ============================================================
# CHANGE 2: "多项增值服务"改为具体数字（7项）
# ============================================================
# 修改1.2目标二中的描述
idx = find_para_by_contains(doc, '面向协会会员机场单位，新多项项增值服务功能模块')
if idx >= 0:
    new_text = '目标二：面向协会会员机场单位，新增7项增值服务功能模块（详见第四章），强化协会对行业的服务与引领；'
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[CHANGE 2-1] 1.2目标二：多项 → 7项')

# 修改4.1中的描述
idx = find_para_by_contains(doc, '规划多项面向会员机场单位的增值服务模块。这些模块将协会')
if idx >= 0:
    new_text = '本章从协会业务视角出发，规划7项面向会员机场单位的增值服务模块。这些模块将协会从"评价工作组织方"进一步定位为"会员机场单位的全周期数字化服务伙伴"与"行业绿色发展的标准制定者与组织推动者"。'
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[CHANGE 2-2] 4.1：多项 → 7项')

# 修改4.11总览中的"9项"
idx = find_para_by_contains(doc, '9 项扩展模块共同实现这一定位转变')
if idx >= 0:
    text = doc.paragraphs[idx].text
    new_text = text.replace('9 项扩展模块', '7项扩展模块')
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[CHANGE 2-3] 4.11：9项 → 7项')

# ============================================================
# CHANGE 3: 修改"二期新建立数据库"段落
# 添加图片转文字算法逻辑，避免AI直接参与
# ============================================================
idx = find_para_by_contains(doc, '二期新建立数据库，通过关键词抓取PDF报告数据')
if idx >= 0:
    new_text = '现状：四星评价强制"碳达峰路径规划研究"，3 家四星机场都有，但报告没汇总。二期新建立数据库，通过图片转文字算法逻辑做抓取，避免AI的直接参与导致数据信息泄露。'
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[CHANGE 3] 修改：添加图片转文字算法逻辑，避免AI直接参与')
else:
    changes_log.append(f'[SKIP 3] 未找到：二期新建立数据库段落')

# 保存文档
doc.save(doc_path)

# 输出修改日志
print("=" * 50)
print("文档修改完成！")
print("=" * 50)
for log in changes_log:
    print(log)
print("=" * 50)
print(f"共完成 {len(changes_log)} 项修改")