import sys
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import copy

doc_path = r'c:\Users\adone\OneDrive\桌面\AI练习\wendang\fangan\【协会反馈】绿色机场平台二期-建设方案V.020260608.docx'
doc = Document(doc_path)

# Helper: find paragraph index by text content
def find_para_idx(doc, text_fragment):
    for i, p in enumerate(doc.paragraphs):
        if text_fragment in p.text:
            return i
    return -1

# Helper: replace paragraph text
def replace_para_text(doc, idx, new_text):
    if idx < 0 or idx >= len(doc.paragraphs):
        return False
    p = doc.paragraphs[idx]
    # Clear all runs
    for run in p.runs:
        run.text = ''
    if p.runs:
        p.runs[0].text = new_text
    else:
        p.add_run(new_text)
    return True

# Helper: insert paragraph after index
def insert_para_after(doc, idx, text, style_name='Normal'):
    if idx < 0:
        return None
    # Get the element after which to insert
    p_elem = doc.paragraphs[idx]._element
    new_p = doc.add_paragraph(text, style=style_name)
    new_p_elem = new_p._element
    p_elem.addnext(new_p_elem)
    # Remove from end of document
    new_p._element.getparent().remove(new_p._element)
    return new_p

# Helper: delete paragraph
def delete_para(doc, idx):
    if idx < 0 or idx >= len(doc.paragraphs):
        return
    p = doc.paragraphs[idx]
    p._element.getparent().remove(p._element)

# Helper: find all paragraph indices containing text
def find_all_para_idx(doc, text_fragment):
    indices = []
    for i, p in enumerate(doc.paragraphs):
        if text_fragment in p.text:
            indices.append(i)
    return indices

changes_log = []

# ============================================================
# CHANGE 1: 修改 1.2 建设目标 - 目标二
# "新增 9 类增值服务功能模块" → 根据速记，去掉培训、会员互助等，调整为实际数量
# ============================================================
idx = find_para_idx(doc, '目标二：面向协会会员机场单位，新增 9 类增值服务功能模块')
if idx >= 0:
    new_text = '目标二：面向协会会员机场单位，新增多项增值服务功能模块（详见第四章），强化协会对行业的服务与引领；'
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[1.2] 修改目标二：去掉"9类"，改为"多项"')

# ============================================================
# CHANGE 2: 修改 4.1 设计动机与定位
# "规划 9 项面向会员机场单位的增值服务模块" → 调整
# ============================================================
idx = find_para_idx(doc, '本章从协会业务视角出发，规划 9 项面向会员机场单位的增值服务模块')
if idx >= 0:
    new_text = '本章从协会业务视角出发，规划多项面向会员机场单位的增值服务模块。这些模块将协会从"评价工作组织方"进一步定位为"会员机场单位的全周期数字化服务伙伴"与"行业绿色发展的标准制定者与组织推动者"。'
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[4.1] 修改：去掉"9项"，改为"多项"')

# ============================================================
# CHANGE 3: 修改 4.2 标题和内容 - 新技术与新产品推广→只做技术推广
# ============================================================
idx = find_para_idx(doc, '4.2 新技术与新产品推广服务')
if idx >= 0:
    replace_para_text(doc, idx, '4.2 新技术推广服务')
    changes_log.append(f'[4.2] 标题修改：去掉"与新产品"')

# 修改4.2业务痛点 - 去掉"产品"相关
idx = find_para_idx(doc, '业务痛点：协会作为行业组织，掌握大量先进技术与产品资源')
if idx >= 0:
    new_text = '业务痛点：协会作为行业组织，掌握大量先进技术资源，但向会员机场单位的推广长期依赖会议、展会、线下走访等高成本低频次方式，且推送缺乏个性化。会员单位也难以判断哪些技术真正适合自身。'
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[4.2] 业务痛点：去掉"产品"相关描述')

# 修改4.2核心功能 - 去掉"供应商资质管理"（产品相关），保留技术推广
idx = find_para_idx(doc, '供应商资质管理：协会对入库技术与供应商进行资质审核')
if idx >= 0:
    new_text = '技术入库管理：协会对入库技术进行资质审核，建立白名单机制，保障推荐质量'
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[4.2] 核心功能：供应商资质管理→技术入库管理')

# 修改4.2使用角色 - 去掉"技术供应商"
idx = find_para_idx(doc, '主要使用角色：协会管理员（运营推广）、机场用户（接收推荐）、技术供应商（提交资料）')
if idx >= 0:
    new_text = '主要使用角色：协会管理员（运营推广）、机场用户（接收推荐）'
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[4.2] 使用角色：去掉"技术供应商"')

# ============================================================
# CHANGE 4: 修改 4.3 评星预打分系统 - 重大修改
# 预打分要做，星级预测不做，差距分析不做，不给建议
# 不收集不评估不储存不分析不建议，生成就定期删除数据不做储存
# 声明不对数据结果进行负责，机场自己可以有平台历史分数报告数据
# ============================================================

# 修改4.3标题
idx = find_para_idx(doc, '4.3 评星预打分系统')
if idx >= 0:
    replace_para_text(doc, idx, '4.3 评星预打分')
    changes_log.append(f'[4.3] 标题修改：去掉"系统"')

# 修改4.3业务痛点
idx = find_para_idx(doc, '业务痛点：目前机场只有在正式申报"双碳机场"评价时才能得知自己能否达到目标星级')
if idx >= 0:
    new_text = '业务痛点：目前机场只有在正式申报"双碳机场"评价时才能得知自己能否达到目标星级，缺乏过程性反馈。机场需要一个工具，在正式申报前自助评估自身得分情况。'
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[4.3] 业务痛点：简化描述，去掉"明确差距、有针对性地提升"')

# 删除"星级预测"段落
idx = find_para_idx(doc, '星级预测：基于自评结果与历史评价数据，预测机场可能达到的星级（含置信度）')
if idx >= 0:
    delete_para(doc, idx)
    changes_log.append(f'[4.3] 删除：星级预测')

# 删除"差距分析"段落
idx = find_para_idx(doc, '差距分析：列出离下一星级仍差多少分')
if idx >= 0:
    delete_para(doc, idx)
    changes_log.append(f'[4.3] 删除：差距分析')

# 删除"整改优先级排序"段落
idx = find_para_idx(doc, '整改优先级排序：结合"提升性价比"')
if idx >= 0:
    delete_para(doc, idx)
    changes_log.append(f'[4.3] 删除：整改优先级排序')

# 修改"预打分历史追踪"段落
idx = find_para_idx(doc, '预打分历史追踪：机场可周期性自评，平台留存历史趋势')
if idx >= 0:
    new_text = '预打分历史查看：机场可周期性自评，平台短期存储打分记录供机场查看，定期自动删除，不做长期储存。平台不对预打分数据结果负责，机场可自行保留历史分数报告数据。'
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[4.3] 修改：预打分历史追踪→预打分历史查看，添加免责声明')

# 删除"协会侧汇总"段落
idx = find_para_idx(doc, '协会侧汇总：协会可看到全行业机场预打分分布')
if idx >= 0:
    delete_para(doc, idx)
    changes_log.append(f'[4.3] 删除：协会侧汇总')

# 修改4.3模块协同
idx = find_para_idx(doc, '模块协同：复用第三章的"评价指标失分热力图"分析引擎')
if idx >= 0:
    new_text = '模块协同：复用第三章的"评价指标失分热力图"分析引擎；与插件工具模块的"机场自评分速算工具"差异——后者是公众端轻量版，本模块为机场账户登录后的完整版。本模块仅做打分功能，不给建议内容。'
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[4.3] 模块协同：添加"仅做打分功能，不给建议内容"')

# ============================================================
# CHANGE 5: 删除 4.6 会员单位互助协作平台（不做）
# ============================================================
# Find and delete all paragraphs belonging to section 4.6
start_idx = find_para_idx(doc, '4.6 会员单位互助协作平台')
if start_idx >= 0:
    # Find the end of section 4.6 (next Heading 2)
    end_idx = -1
    for i in range(start_idx + 1, len(doc.paragraphs)):
        if doc.paragraphs[i].style.name == 'Heading 2':
            end_idx = i
            break
    
    if end_idx > start_idx:
        # Delete from end to start to avoid index shifting
        for i in range(end_idx - 1, start_idx - 1, -1):
            delete_para(doc, i)
        changes_log.append(f'[4.6] 删除整个章节：会员单位互助协作平台（不做）')

# ============================================================
# CHANGE 6: 删除 4.7 培训与认证管理体系（不做）
# ============================================================
start_idx = find_para_idx(doc, '4.7 培训与认证管理体系')
if start_idx >= 0:
    end_idx = -1
    for i in range(start_idx + 1, len(doc.paragraphs)):
        if doc.paragraphs[i].style.name == 'Heading 2':
            end_idx = i
            break
    
    if end_idx > start_idx:
        for i in range(end_idx - 1, start_idx - 1, -1):
            delete_para(doc, i)
        changes_log.append(f'[4.7] 删除整个章节：培训与认证管理体系（不做）')

# ============================================================
# CHANGE 7: 修改 4.4 会员单位个性化提升包 - 简化，去掉AI相关
# 根据速记：简单做，不用AI。明确需要协会提供的材料。
# 与会员单位个性化提升包结合考虑，简单的思路即可
# ============================================================

# 修改4.4业务痛点
idx = find_para_idx(doc, '业务痛点：现年度评价报告的"提升建议"章节为通用建议')
if idx >= 0:
    new_text = '业务痛点：现年度评价报告的"提升建议"章节为通用建议，机场拿到后难以与自身情况关联。协会希望为每家会员机场单位输出一份个性化的提升方案。'
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[4.4] 业务痛点：简化描述')

# 修改4.4核心功能 - 去掉AI相关，简化
idx = find_para_idx(doc, '机场画像卡：每家会员机场单位自动生成一页画像')
if idx >= 0:
    new_text = '短板个性清单：基于失分分析，为每家会员机场单位生成个性化短板清单，按优先级排序'
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[4.4] 核心功能：机场画像卡→短板个性清单')

# 修改个性化提升清单
idx = find_para_idx(doc, '个性化提升清单：基于失分分析与升星路径分析自动生成')
if idx >= 0:
    new_text = '实际案例关联：将短板清单与实际案例、榜单关联，便于机场参考改进方案节能'
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[4.4] 核心功能：个性化提升清单→实际案例关联')

# 删除ROI测算（需要模型/AI）
idx = find_para_idx(doc, 'ROI 测算：每项建议附带预估投资额与预估减排量')
if idx >= 0:
    delete_para(doc, idx)
    changes_log.append(f'[4.4] 删除：ROI测算（需要模型/AI）')

# 删除提升路径推演
idx = find_para_idx(doc, '提升路径推演：基于历史升星数据')
if idx >= 0:
    delete_para(doc, idx)
    changes_log.append(f'[4.4] 删除：提升路径推演')

# 删除年度评估闭环
idx = find_para_idx(doc, '年度评估闭环：次年评价后自动对比')
if idx >= 0:
    delete_para(doc, idx)
    changes_log.append(f'[4.4] 删除：年度评估闭环')

# ============================================================
# CHANGE 8: 修改第三章相关内容
# 失分数据采集方式（表格填写）
# ============================================================
idx = find_para_idx(doc, '方向 ⑤ 评价指标失分热力图')
if idx >= 0:
    # Find the "现状" paragraph after this heading
    for i in range(idx + 1, min(idx + 5, len(doc.paragraphs))):
        if '现状：31 项指标' in doc.paragraphs[i].text:
            new_text = '现状：31 项指标 × 20 家机场 = 620 个评价单元，足够做模式分析。失分数据通过表格填写方式采集。'
            replace_para_text(doc, i, new_text)
            changes_log.append(f'[3.3方向⑤] 添加：失分数据通过表格填写方式采集')
            break

# ============================================================
# CHANGE 9: 修改方向② - 添加案例与榜单关联
# ============================================================
idx = find_para_idx(doc, '方向 ② 投入产出 ROI 分析')
if idx >= 0:
    for i in range(idx + 1, min(idx + 5, len(doc.paragraphs))):
        if '平台输出：万元投入' in doc.paragraphs[i].text:
            new_text = '平台输出：万元投入 → 吨 CO₂减排曲线；按项目类型的边际减排成本对比；"性价比 TOP 10 项目"榜单。实际案例与榜单关联，便于机场改进方案节能。'
            replace_para_text(doc, i, new_text)
            changes_log.append(f'[3.3方向②] 添加：实际案例与榜单关联')
            break

# ============================================================
# CHANGE 10: 修改方向⑧ - PDF报告数据抓取
# ============================================================
idx = find_para_idx(doc, '方向 ⑧ 行业趋势预测与碳达峰路线汇总')
if idx >= 0:
    for i in range(idx + 1, min(idx + 5, len(doc.paragraphs))):
        if '现状：四星评价强制' in doc.paragraphs[i].text:
            new_text = '现状：四星评价强制"碳达峰路径规划研究"，3 家四星机场都有，但报告没汇总。二期新建立数据库，通过关键词抓取PDF报告数据。'
            replace_para_text(doc, i, new_text)
            changes_log.append(f'[3.3方向⑧] 添加：二期新建立数据库，通过关键词抓取PDF报告数据')
            break

# ============================================================
# CHANGE 11: 修改1.1项目背景 - 添加行业数据参考说明
# ============================================================
idx = find_para_idx(doc, '二期建设的核心使命，是把协会已经在收集的丰富数据真正"激活"')
if idx >= 0:
    new_text = '二期建设的核心使命，是把协会已经在收集的丰富数据真正"激活"——让数据从存档资料，变成行业决策工具与机场提升路径的源头；同时，从协会向会员机场单位提供服务的视角，扩展出新技术推广、评星预打分等增值功能，全面提升协会对会员单位的服务深度与行业引领作用。注：协会对于机场减排数据收集暂没计划，二期数据主要基于一期已收集数据及PDF报告抓取。'
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[1.1] 添加：协会对于机场减排数据收集暂没计划的说明')

# ============================================================
# CHANGE 12: 修改4.5 定向推送 - 去掉短信推送
# ============================================================
idx = find_para_idx(doc, '发布即推送：新政策/标准发布后，自动识别可能受影响的会员机场单位群，定向推送')
if idx >= 0:
    new_text = '发布即推送：新政策/标准发布后，自动识别可能受影响的会员机场单位群，通过平台内消息推送（定向推送短信暂时不做）'
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[4.5] 修改：定向推送短信暂时不做，改为平台内消息推送')

# ============================================================
# CHANGE 13: 修改4.2模块协同 - 去掉"画像引擎"
# ============================================================
idx = find_para_idx(doc, '模块协同：复用资源库模块的"技术产品库"作为数据底座，复用数据模块的画像引擎做匹配。')
if idx >= 0:
    new_text = '模块协同：复用资源库模块的"技术产品库"作为数据底座，通过算法和AI进行技术标注与匹配。'
    replace_para_text(doc, idx, new_text)
    changes_log.append(f'[4.2] 模块协同：改为通过算法和AI进行技术标注与匹配')

# ============================================================
# Save the modified document
# ============================================================
output_path = r'c:\Users\adone\OneDrive\桌面\AI练习\wendang\fangan\【协会反馈】绿色机场平台二期-建设方案V.020260608_修改版.docx'
doc.save(output_path)

print("修改完成！")
print(f"输出文件：{output_path}")
print("\n修改日志：")
for log in changes_log:
    print(f"  {log}")
print(f"\n共执行 {len(changes_log)} 项修改")
