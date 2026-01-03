# -*- coding: utf-8 -*-
"""
只在太空人轮廓内部生成词云
"""
import os
import jieba
import numpy as np
import cv2
from PIL import Image
from wordcloud import WordCloud
from matplotlib import pyplot as plt
from docx import Document

# ------------------------------------------------
# 1. 读 Word 文档，jieba 分词
# ------------------------------------------------
doc_path = "基建工程.docx"
doc = Document(doc_path)
text = '\n'.join(p.text for p in doc.paragraphs if p.text.strip())
words = [w for w in jieba.lcut(text) if len(w) > 1]

# ------------------------------------------------
# 2. 生成“纯净”二值 mask（背景=0，主体=255）
# ------------------------------------------------
png_path = r"D:\python\基建工程.png"   # ← 换成你的图片
width, height = 600, 800             # 输出词云尺寸

img = Image.open(png_path)

# 优先使用透明通道，否则转灰度
if img.mode in ('RGBA', 'LA') and 'transparency' in img.info or img.mode == 'RGBA':
    img = img.convert('RGBA').resize((width, height), Image.LANCZOS)
    alpha = np.array(img)[:, :, 3]          # 0=透明 255=不透明
    _, mask = cv2.threshold(alpha, 0, 255, cv2.THRESH_BINARY)
else:
    img = img.convert('L').resize((width, height), Image.LANCZOS)
    _, mask = cv2.threshold(np.array(img), 0, 255,
                            cv2.THRESH_BINARY + cv2.THRESH_OTSU)

# 确保“主体”是 255（如果背景亮就反色）
if np.mean(mask) > 127:
    mask = 255 - mask

# 形态学清噪，让边缘更平滑
kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3))
mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=2)
mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN,  kernel, iterations=2)

print("mask 唯一值：", np.unique(mask))   # 应该只有 [0 255]
Image.fromarray(mask).save('mask_debug.png')

# ------------------------------------------------
# 3. 生成词云
# ------------------------------------------------
font = r"C:\Windows\Fonts\simhei.ttf"   # 黑体
wc = WordCloud(
        mask=mask,
        font_path=font,
        background_color='black',
        colormap='plasma',
        max_words=400,
        contour_width=1,
        contour_color='yellow',
        scale=2
).generate(' '.join(words))

# ------------------------------------------------
# 4. 保存 & 显示
# ------------------------------------------------
plt.figure(figsize=(6, 8), facecolor='black')
plt.imshow(wc, interpolation='bilinear')
plt.axis('off')
plt.tight_layout(pad=0)
out_file = '基建工程词云_剪影.png'
plt.savefig(out_file, dpi=300, bbox_inches='tight', facecolor='black')
print('已保存：', os.path.abspath(out_file))
plt.show()
