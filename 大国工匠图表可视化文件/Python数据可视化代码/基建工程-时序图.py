# -*- coding: utf-8 -*-
"""
《基建工程》六维情感时序图（纯 CPU + 中文无乱码）
"""
import re
import os
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
from docx import Document

# ---------- 1. 针对「基建工程」的 6 维词典 ----------
LEX = {
    "精度":   "精度|毫米|微米|丝|零误差|毫米级|厘米级",
    "匠心":   "匠心|精益求精|手工|研磨|吊装|吊装工|钳工",
    "安全":   "安全|零事故|滴水不漏|人命|风险|守护",
    "创新":   "创新|突破|首创|模拟|自动化|机器人|新工艺",
    "坚守":   "坚守|30年|40年|一辈子|传帮带|扎根|归零",
    "自豪":   "自豪|大国工程|超级工程|世界之最|中国奇迹"
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
    ax.set_title('《基建工程》文本六维情感时序图（句级）', fontproperties=font, fontsize=16)
    ax.set_xlabel('句子序号', fontproperties=font, fontsize=12)
    ax.set_ylabel('情感强度（0-1）', fontproperties=font, fontsize=12)
    ax.legend(prop=font, loc='upper right', frameon=True, shadow=True)
    ax.text(0.02, 0.02,
            '数据来源：基建工程.docx | 方法：正则词典\n生成时间：' + pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
            transform=ax.transAxes, fontsize=9,
            bbox=dict(boxstyle="round,pad=0.3", facecolor="lightgray", alpha=0.5))
    plt.tight_layout()
    plt.savefig("infrastructure_timeline_final.png", dpi=600, bbox_inches="tight")
    plt.show()

# ---------- 4. 主流程 ----------
def main():
    if not os.path.exists("基建工程.docx"):
        print("❌ 请将「基建工程.docx」放在当前目录再运行！")
        return
    paragraphs = read_docx("基建工程.docx")
    sentences  = []
    for p in paragraphs:
        sentences.extend(sent_cut(p))
    print(f"共切分 {len(sentences)} 句，开始打分...")
    df = pd.DataFrame(sentences, columns=["sentence"])
    lex_scores = df["sentence"].apply(lexicon_score)
    for idx, dim in enumerate(LEX.keys()):
        df[dim] = [v[idx] for v in lex_scores]
    df.to_csv("infrastructure_6d.csv", index=False, encoding="utf-8-sig")
    plot_timeline(df)
    print("✅ 完成！文件：infrastructure_timeline_final.png | infrastructure_6d.csv")

if __name__ == "__main__":
    main()
