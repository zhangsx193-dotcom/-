# -*- coding: utf-8 -*-
"""
能源电力领域-大国工匠精神品质热力图数据生成器
"""
import re, pandas as pd
from sklearn.feature_extraction.text import TfidfTransformer
from docx import Document

INDUSTRY_DICT = {
    "核电维修": ["陈永伟", "核电站", "反应堆", "0.1毫米", "传感器"],
    "特高压带电": ["王进", "特高压", "带电检修", "1000千伏", "秋千法"],
    "核电燃料": ["乔素凯", "核燃料", "水下修复", "四米长杆", "零失误"],
    "电能计量": ["黄金娟", "电能表", "自动化检定", "2秒钟", "58倍"],
    "火电焊接": ["胡家瑞", "热电", "焊缝", "良心", "四十年"],
    "水电吊装": ["梅琳", "白鹤滩", "转子吊装", "1毫米", "2300吨"]
}

QUALITY_DICT = {
    "精度": ["毫米", "0.1毫米", "1毫米", "2毫米", "微米"],
    "责任": ["零失误", "良心", "安全", "生命", "零缺陷"],
    "创新": ["自动化", "首创", "改进", "秋千法", "2秒钟"],
    "专注": ["四十年", "26年", "手感", "专注", "肌肉记忆"],
    "吃苦": ["高温", "高空", "60米", "暴晒", "严寒"],
    "学习": ["博士", "读书", "标准", "专利", "论文"],
    "协作": ["团队", "师徒", "班组", "合力", "传帮带"],
    "极致": ["100%", "58倍", "零缺陷", "完美", "极高"]
}

def read_docx(path):
    return "\n".join(p.text for p in Document(path).paragraphs if p.text.strip())

text = read_docx("能源电力.docx")
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
stack.to_csv("能源电力_热力.csv", index=False, encoding="utf-8-sig")
print("已生成 能源电力_热力.csv ，共", len(stack), "条记录")
