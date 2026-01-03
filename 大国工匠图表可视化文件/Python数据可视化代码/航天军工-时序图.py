# -*- coding: utf-8 -*-
"""
《航天军工》六维情感时序图（纯 CPU + 中文无乱码）
"""
import re
import os
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
from docx import Document

# ---------- 1. 六维词典 ----------
LEX = {
    "自豪":   "自豪|骄傲|振奋|长脸|激动|为国争光",
    "敬佩":   "敬佩|致敬|崇高|伟大|了不起|佩服",
    "感动":   "感动|泪目|暖心|热泪盈眶|触动",
    "坚守":   "坚守|耐得住寂寞|30年如一日|扎根|默默",
    "精益求精": "精益求精|零差错|毫厘|极致|完美|一丝不苟",
    "报国":   "报国|航天梦|强军|为国铸剑|奉献|忠诚"
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
        scores.append(min(hits / 3, 1.0))   # 归一化到 0-1
    return scores

# ---------- 3. 画图 ----------
def plot_timeline(df):
    # ① 指定 Windows 自带中文字体（防止方框）
    font = FontProperties(fname=r"c:\windows\fonts\simsun.ttc", size=14)
    plt.rcParams["axes.unicode_minus"] = False   # 修复负号

    plt.style.use("seaborn-v0_8-whitegrid")
    fig, ax = plt.subplots(figsize=(16, 5))
    colors = ["#e60049", "#0bb4ff", "#50d991", "#f6ea5c", "#9b5de5", "#ff8c00"]
    for dim, c in zip(LEX.keys(), colors):
        ax.plot(df.index, df[dim], label=dim, color=c, linewidth=2)
    ax.axhline(0.8, color="crimson", linestyle="--", alpha=0.7, label="高正面阈值")

    # ② 中文标题、轴、图例
    ax.set_title('《航天军工》文本六维情感时序图（句级）', fontproperties=font, fontsize=16)
    ax.set_xlabel('句子序号', fontproperties=font, fontsize=12)
    ax.set_ylabel('情感强度（0-1）', fontproperties=font, fontsize=12)
    ax.legend(prop=font, loc='upper right', frameon=True, shadow=True)

    # ③ 脚注
    ax.text(0.02, 0.02,
            '数据来源：航天军工.docx | 方法：正则词典\n生成时间：' + pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
            transform=ax.transAxes, fontsize=9,
            bbox=dict(boxstyle="round,pad=0.3", facecolor="lightgray", alpha=0.5))

    plt.tight_layout()
    plt.savefig("sentiment_timeline_final.png", dpi=600, bbox_inches="tight")
    plt.show()

# ---------- 4. 主流程 ----------
def main():
    if not os.path.exists("航天军工.docx"):
        print("❌ 请将「航天军工.docx」放在当前目录再运行！")
        return
    paragraphs = read_docx("航天军工.docx")
    sentences  = []
    for p in paragraphs:
        sentences.extend(sent_cut(p))
    print(f"共切分 {len(sentences)} 句，开始打分...")
    df = pd.DataFrame(sentences, columns=["sentence"])
    lex_scores = df["sentence"].apply(lexicon_score)
    for idx, dim in enumerate(LEX.keys()):
        df[dim] = [v[idx] for v in lex_scores]
    df.to_csv("aerospace_6d.csv", index=False, encoding="utf-8-sig")
    plot_timeline(df)
    print("✅ 完成！文件：sentiment_timeline_final.png | aerospace_6d.csv")

if __name__ == "__main__":
    main()
