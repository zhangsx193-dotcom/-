# -*- coding: utf-8 -*-
"""
航天军工领域-大国工匠精神品质热力图数据生成器
运行前：pip install python-docx pandas scikit-learn
"""
import re
import pandas as pd
from sklearn.feature_extraction.text import TfidfTransformer
from docx import Document

# 1. 词典
INDUSTRY_DICT = {
    "数控加工": ["数控铣工", "常晓飞", "数控微雕", "0.03毫米", "0.15毫米"],
    "火箭总装": ["崔蕴", "长征五号", "火箭总装", "咽喉主刀师", "喷管"],
    "飞船总装": ["张舸", "神舟", "飞船总装", "γ放射源", "盲操作"],
    "航天焊接": ["郑兴", "空间站", "舱体焊接", "300米焊缝", "气孔"],
    "雷达装配": ["顾春燕", "金线键合", "太赫兹雷达", "微米级", "芯片键合"],
    "导弹加工": ["阎敏", "导弹", "咽喉主刀师", "0.005毫米", "喷管"],
    "航天材料": ["韩利萍", "长征七号", "发射平台", "四通均流阀", "0.02毫米"],
    "航天对接": ["王曙群", "太空之吻", "对接机构", "100-1=0", "热循环"]
}

QUALITY_DICT = {
    "精度": ["毫米", "微米", "0.03", "0.005", "0.02", "丝"],
    "责任": ["生命", "安全", "航天员", "零缺陷", "100-1=0"],
    "创新": ["绝技", "首创", "改进", "发明", "操作法"],
    "专注": ["盲操作", "手感", "反复", "凌晨", "十万次"],
    "吃苦": ["高温", "60度", "晕船", "通宵", "加班"],
    "学习": ["博士", "读书", "返校", "理论", "新知"],
    "协作": ["团队", "师徒", "传帮带", "班组", "合力"],
    "极致": ["100%", "零缺陷", "万无一失", "完美", "极值"]
}

def read_docx(path):
    return "\n".join(p.text for p in Document(path).paragraphs if p.text.strip())

text = read_docx("航天军工.docx")

# 2. 匹配
sents = re.split(r'[。！？]', text)
ind2q = {ind: {q: 0 for q in QUALITY_DICT} for ind in INDUSTRY_DICT}

for sent in sents:
    for ind, kw in INDUSTRY_DICT.items():
        if any(k in sent for k in kw):
            for qual, kw_list in QUALITY_DICT.items():
                ind2q[ind][qual] += sum(sent.count(w) for w in kw_list)

# 3. TF-IDF
industries = list(INDUSTRY_DICT.keys())
qualities = list(QUALITY_DICT.keys())
freq = pd.DataFrame(0, index=industries, columns=qualities)
for ind, qdict in ind2q.items():
    for q, v in qdict.items():
        freq.loc[ind, q] = v

tfidf = TfidfTransformer(norm=None).fit_transform(freq.values)
score = ((tfidf.toarray() - tfidf.min()) / (tfidf.max() - tfidf.min() + 1e-8) * 100).round(1)
score_mat = pd.DataFrame(score, index=industries, columns=qualities)

# 4. 写CSV
stack = score_mat.stack().reset_index()
stack.columns = ["行业", "品质", "分数"]
stack.to_csv("航天军工_热力.csv", index=False, encoding="utf-8-sig")
print("已生成 航天军工_热力.csv ，共", len(stack), "条记录")
