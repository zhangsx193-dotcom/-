# -*- coding: utf-8 -*-
"""
文化传承领域-大国工匠精神品质热力图数据生成器
"""
import re, pandas as pd
from sklearn.feature_extraction.text import TfidfTransformer
from docx import Document

INDUSTRY_DICT = {
    "宣纸晒纸": ["毛胜利", "晒纸", "三丈三", "11米", "头刷"],
    "宣纸捞纸": ["周东红", "捞纸", "7秒", "1800张", "100%正品率"],
    "古编钟": ["刘佑年", "编钟", "调音", "22道工序", "双音"],
    "古丝绸": ["王亚蓉", "丝绸", "东周", "2600年", "0.1毫米"],
    "书画修复": ["单嘉玖", "书画修复", "揭", "全色", "画医"],
    "硬币雕刻": ["余敏", "硬币", "牡丹币", "熊猫币", "0.07毫米"],
    "钞票雕刻": ["马荣", "人民币", "凹版雕刻", "0.16毫米", "毛泽东像"]
}

QUALITY_DICT = {
    "精度": ["毫米", "0.1毫米", "0.07毫米", "0.16毫米", "双音"],
    "责任": ["100%正品率", "画医", "传世", "零缺陷", "国家形象"],
    "创新": ["复活", "复刻", "失传", "改进", "首创"],
    "专注": ["7秒", "22道工序", "30年", "40年", "揭"],
    "吃苦": ["酷暑", "弯腰", "腰椎", "胃病", "熬夜"],
    "学习": ["自学", "实验考古", "研究", "美院", "技艺"],
    "协作": ["师徒", "三人组", "传承", "传帮带", "合作"],
    "极致": ["100%", "零瑕疵", "完美", "极高", "唯一"]
}

def read_docx(path):
    return "\n".join(p.text for p in Document(path).paragraphs if p.text.strip())

text = read_docx("文化传承.docx")
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
stack.to_csv("文化传承_热力.csv", index=False, encoding="utf-8-sig")
print("已生成 文化传承_热力.csv ，共", len(stack), "条记录")
