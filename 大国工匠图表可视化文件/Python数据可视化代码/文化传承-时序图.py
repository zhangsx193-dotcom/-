# -*- coding: utf-8 -*-
"""
《文化传承》六维情感时序图（纯 CPU + 中文无乱码）
"""
import re
import os
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
from docx import Document

# ---------- 1. 针对「文化传承」的 6 维词典 ----------
LEX = {
    "坚守":   "坚守|传承|坚持|执着|耐得住寂寞|30年如一日|磨性情",
    "匠心":   "匠心|精益求精|零差错|毫厘|完美|一丝不苟|妙手回春",
    "敬畏":   "敬畏|敬畏心|小心翼翼|如履薄冰|文物为生命|不能断",
    "自豪":   "自豪|骄傲|国宝|皇冠明珠|中华文化|民族|传世",
    "感动":   "感动|泪目|暖心|触动|震撼|情不自禁",
    "创新":   "创新|突破|失传|复活|复制|复原|模拟考古"
}

# ---------- 2. 工具函数 ----------
def read_docx(path):
    doc = Document(path)
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]

def sent_cut(text):
    return [s.strip() for s in re.split(r'[。！？；]', text) if s.strip()]

def lexicon_score(sentence):
    scores = []
    for dim, pattern in LEX.items():
        hits = len(re.findall(pattern, sentence))
        scores.append(min(hits / 3, 1.0))
    return scores

# ---------- 3. 画图（中文无乱码） ----------
def plot_timeline(df):
    font = FontProperties(fname=r"c:\windows\fonts\simsun.ttc", size=14)
    plt.rcParams["axes.unicode_minus"] = False
    plt.style.use("seaborn-v0_8-whitegrid")
    fig, ax = plt.subplots(figsize=(16, 5))
    colors = ["#e60049", "#0bb4ff", "#50d991", "#f6ea5c", "#9b5de5", "#ff8c00"]
    for dim, c in zip(LEX.keys(), colors):
        ax.plot(df.index, df[dim], label=dim, color=c, linewidth=2)
    ax.axhline(0.8, color="crimson", linestyle="--", alpha=0.7, label="高正面阈值")
    ax.set_title('《文化传承》文本六维情感时序图（句级）', fontproperties=font, fontsize=16)
    ax.set_xlabel('句子序号', fontproperties=font, fontsize=12)
    ax.set_ylabel('情感强度（0-1）', fontproperties=font, fontsize=12)
    ax.legend(prop=font, loc='upper right', frameon=True, shadow=True)
    ax.text(0.02, 0.02,
            '数据来源：文化传承.docx | 方法：正则词典\n生成时间：' + pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
            transform=ax.transAxes, fontsize=9,
            bbox=dict(boxstyle="round,pad=0.3", facecolor="lightgray", alpha=0.5))
    plt.tight_layout()
    plt.savefig("culture_timeline_final.png", dpi=600, bbox_inches="tight")
    plt.show()

# ---------- 4. 主流程 ----------
def main():
    if not os.path.exists("文化传承.docx"):
        print("❌ 请将「文化传承.docx」放在当前目录再运行！")
        return
    paragraphs = read_docx("文化传承.docx")
    sentences  = []
    for p in paragraphs:
        sentences.extend(sent_cut(p))
    print(f"共切分 {len(sentences)} 句，开始打分...")
    df = pd.DataFrame(sentences, columns=["sentence"])
    lex_scores = df["sentence"].apply(lexicon_score)
    for idx, dim in enumerate(LEX.keys()):
        df[dim] = [v[idx] for v in lex_scores]
    df.to_csv("culture_6d.csv", index=False, encoding="utf-8-sig")
    plot_timeline(df)
    print("✅ 完成！文件：culture_timeline_final.png | culture_6d.csv")

if __name__ == "__main__":
    main()
