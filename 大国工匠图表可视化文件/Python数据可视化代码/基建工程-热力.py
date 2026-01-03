# -*- coding: utf-8 -*-
"""
行业-品质热力图数据生成器（沿用 sklearn，仅追加写 CSV）
运行前：pip install python-docx pandas scikit-learn
"""

import re
from collections import defaultdict
import pandas as pd
from sklearn.feature_extraction.text import TfidfTransformer
from docx import Document

# -------------------- 1. 词典 --------------------
INDUSTRY_DICT = {
    "航空制造": ["运-20", "C919", "大飞机", "机身", "钣金", "舱门", "蒙皮", "胡洋", "王伟"],
    "核电建设": ["核电站", "主管道", "焊工", "未晓朋", "田湾", "核电", "焊缝"],
    "射电望远镜": ["FAST", "反射面板", "吊装", "周永和", "射电望远镜", "天眼"],
    "海底隧道/桥梁": ["港珠澳大桥", "沉管", "管延安", "海底隧道", "深中通道", "大连湾"],
    "LNG船舶": ["LNG船", "殷瓦钢", "张冬伟", "液化天然气", "围护系统"]
}

QUALITY_DICT = {
    "精度": ["毫米", "0.5毫米", "0.25毫米", "误差", "精度", "严丝合缝", "缝隙"],
    "责任": ["责任", "耻辱", "荣耀", "名字", "刻", "终身", "一辈子"],
    "创新": ["数字化", "零的突破", "第一次", "首创", "革命性", "探索"],
    "专注": ["专注", "耐心", "重复", "一遍又一遍", "再检查", "强迫症"],
    "吃苦": ["高温", "汗水", "狭小", "闷热", "烫", "疼", "苦"],
    "学习": ["学习", "笔记", "日志", "培训", "请教", "啃", "坚持"],
    "协作": ["团队", "合力", "配合", "并肩", "专家团", "几十人"],
    "极致": ["百分之百", "100%", "0漏点", "天衣无缝", "无可挑剔", "完美"]
}

# -------------------- 2. 读 Word --------------------
def read_docx(path):
    return "\n".join(p.text for p in Document(path).paragraphs if p.text.strip())

text = read_docx("基建工程.docx")   # 如需换路径，改这里

# -------------------- 3. 匹配 --------------------
sents = re.split(r'[。！？]', text)
ind2q = defaultdict(lambda: defaultdict(int))

for sent in sents:
    inds = [ind for ind, kw in INDUSTRY_DICT.items() if any(k in sent for k in kw)]
    if not inds:
        continue
    for qual, kw_list in QUALITY_DICT.items():
        cnt = sum(sent.count(k) for k in kw_list)
        if cnt:
            for ind in inds:
                ind2q[ind][qual] += cnt

# -------------------- 4. sklearn TF-IDF --------------------
industries = list(ind2q.keys())
qualities = list(QUALITY_DICT.keys())

freq_mat = pd.DataFrame(0, index=industries, columns=qualities)
for ind, qdict in ind2q.items():
    for q, v in qdict.items():
        freq_mat.loc[ind, q] = v

tfidf = TfidfTransformer(norm=None).fit_transform(freq_mat.values)
tfidf_dense = tfidf.toarray()

# 0-100 归一化
score_mat = pd.DataFrame(
    ((tfidf_dense - tfidf_dense.min()) /
     (tfidf_dense.max() - tfidf_dense.min() + 1e-8) * 100),
    index=industries,
    columns=qualities
).round(1)

# -------------------- 5. 写 CSV --------------------
stack = score_mat.stack().reset_index()
stack.columns = ["行业", "品质", "分数"]
stack.to_csv("行业_品质_分数.csv", index=False, encoding="utf-8-sig")
print("已生成 行业_品质_分数.csv ，共", len(stack), "条记录")
