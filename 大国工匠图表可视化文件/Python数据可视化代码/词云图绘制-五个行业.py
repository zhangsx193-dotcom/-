# -*- coding: utf-8 -*-
"""
装备制造轮廓内词云（已去停用词）
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
# 1. 读 Word 并分词 + 去停用词
# ------------------------------------------------
doc_path = "装备制造.docx"
doc = Document(doc_path)
text = '\n'.join(p.text for p in doc.paragraphs if p.text.strip())

# 简单停用词表（可继续往里面加）
stop = {'我们',"一个","零件", "文墨","这个","就是","文波","自己","剑锋","马荣","王曙群","顾秋亮","延安","曹彦生","常晓飞",
        '工作', '进行', '相关', '张冬伟', '实现', '时候',
        '开展', '完成', '这些', '卢仁峰', '他们', '可以', '需要',
        '主要', '通过', '不断', '充分', '胡洋', '目前',
        '已经', '正在', '凤林', '宁允展', '这样', '今天', '胡双',
        '洪家', '闫敏', '那个', "梅琳", '的话', '之一', '一些',
        '所有', '每个', '各种', '这项', '该', '本', '将', '了', '的', '和', '与', '在', '为', '是', '对', '及', '等', '等等'}

words = [w for w in jieba.lcut(text)
         if len(w) > 1 and w not in stop]

# ------------------------------------------------
# 2. 生成“装备制造”二值 mask（轮廓=255，其余=0）
# ------------------------------------------------
png_path = r"D:\python\装备制造.png"
width, height = 600, 800

img = Image.open(png_path)
if img.mode == 'RGBA':
    img = img.resize((width, height), Image.LANCZOS)
    alpha = np.array(img)[:, :, 3]
    _, mask = cv2.threshold(alpha, 0, 255, cv2.THRESH_BINARY)
else:
    img = img.convert('L').resize((width, height), Image.LANCZOS)
    _, mask = cv2.threshold(np.array(img), 0, 255,
                            cv2.THRESH_BINARY + cv2.THRESH_OTSU)

if np.mean(mask) < 127:
    mask = 255 - mask

kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3))
mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=2)
mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel, iterations=2)

print('mask 唯一值：', np.unique(mask))
Image.fromarray(mask).save('mask_debug.png')

# ------------------------------------------------
# 3. 生成词云
# ------------------------------------------------
font = r"C:\Windows\Fonts\simhei.ttf"
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
out_file = '装备制造词云_去停用词.png'
plt.savefig(out_file, dpi=300, bbox_inches='tight', facecolor='black')
print('已保存：', os.path.abspath(out_file))
plt.show()
