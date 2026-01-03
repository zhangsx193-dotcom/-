# -*- coding: utf-8 -*-
"""
装备制造领域-大国工匠精神品质热力图数据生成器
"""
import re, pandas as pd
from sklearn.feature_extraction.text import TfidfTransformer
from docx import Document

INDUSTRY_DICT = {
    "ROV操控": ["韩超", "ROV", "深海一号", "1500米", "脐带缆"],
    "设备再造": ["刘云清", "改造", "1微米", "清洗机", "复兴号"],
    "货车焊接": ["易冉", "C70E", "铁路货车", "0.2毫米", "超声波"],
    "转向架研磨": ["宁允展", "定位臂", "0.05毫米", "转向架", "研磨"],
    "铸造控制": ["毛正石", "铸造", "10℃", "0误差", "叶片"],
    "深孔加工": ["戎鹏强", "深孔", "0.01毫米", "身管", "火炮"]
}

QUALITY_DICT = {
    "精度": ["微米", "0.05毫米", "0.01毫米", "10℃", "0.2毫米"],
    "责任": ["零误差", "生命", "安全", "军品", "万无一失"],
    "创新": ["改造", "再造", "首创", "操作法", "突破"],
    "专注": ["0.01毫米", "纯手工", "研磨", "深孔", "专注"],
    "吃苦": ["高温", "高空", "加班", "熬夜", "连续"],
    "学习": ["自学", "编程", "机器人", "软件", "标准"],
    "协作": ["团队", "班组", "师徒", "合力", "传帮带"],
    "极致": ["100%", "零缺陷", "极高", "完美", "顶尖"]
}

def read_docx(path):
    return "\n".join(p.text for p in Document(path).paragraphs if p.text.strip())

text = read_docx("装备制造.docx")
sents = re.split(r'[。！？]', text)
ind2q = {ind: {q: 0 for q in QUALITY_DICT} for ind in INDUSTRY_DICT}

for sent in sents:
    for ind, kw in INDUSTRY_DICT.items():
        if any(k in sent for k in kw):
            for qual, kw_list in QUALITY_DICT.items():
                ind2q[ind][qual] += sum(sent.count(w) for w in kw_list)

industries = list(INDUSTRY_DICT.keys())
qualities = list(QUALITY_DICT.keys())
freq = pd.DataFrame(0, index=industries, columns=qualities)
for ind, qdict in ind2q.items():
    for q, v in qdict.items():
        freq.loc[ind, q] = v

tfidf = TfidfTransformer(norm=None).fit_transform(freq.values)
score = ((tfidf.toarray() - tfidf.min()) / (tfidf.max() - tfidf.min() + 1e-8) * 100).round(1)
score_mat = pd.DataFrame(score, index=industries, columns=qualities)

stack = score_mat.stack().reset_index()
stack.columns = ["行业", "品质", "分数"]
stack.to_csv("装备制造_热力.csv", index=False, encoding="utf-8-sig")
print("已生成 装备制造_热力.csv ，共", len(stack), "条记录")
