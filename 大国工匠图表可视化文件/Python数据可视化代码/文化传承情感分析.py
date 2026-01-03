import re
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
import os
import warnings
warnings.filterwarnings('ignore')

# 设置工作目录
os.chdir('E:\\')

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

print("=" * 60)
print("文化传承领域情感分析系统")
print("工作目录:", os.getcwd())
print("=" * 60)

def read_docx_file(file_path):
    try:
        try:
            from docx import Document
        except ImportError:
            print("正在安装python-docx库...")
            import subprocess
            import sys
            subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx"])
            from docx import Document
        doc = Document(file_path)
        full_text = []
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                full_text.append(text)
        content = '\n'.join(full_text)
        print(f"✓ DOCX文件读取成功, 长度: {len(content)} 字符")
        return content
    except Exception as e:
        print(f"✗ 读取DOCX文件失败: {e}")
        try:
            print("尝试二进制读取...")
            with open(file_path, 'rb') as f:
                raw_content = f.read()
            text_content = ""
            try:
                text_content = raw_content.decode('utf-8', errors='ignore')
                import string
                printable = set(string.printable)
                text_content = ''.join(filter(lambda x: x in printable, text_content))
            except:
                pass
            if len(text_content) > 100:
                print(f"✓ 二进制读取成功, 长度: {len(text_content)} 字符")
                return text_content[:10000]
        except Exception as e2:
            print(f"二进制读取也失败: {e2}")
        return ""

def extract_stories_from_text(text):
    stories = []
    pattern = r'\[([^\]]+)\]\{\.mark\}\s*\n*(.+?)(?=\n\s*\[[^\]]+\]\{\.mark\}|\Z)'
    matches = re.findall(pattern, text, re.DOTALL)
    print(f"正则匹配找到 {len(matches)} 个故事")
    for name, content in matches:
        content = content.strip()
        if len(content) > 30:
            stories.append({
                'name': name.strip(),
                'content': content[:1500]
            })
    if len(stories) < 3:
        print("使用简单方法补充...")
        lines = text.split('\n')
        current_name = ""
        current_content = []
        for line in lines:
            line = line.strip()
            if line.startswith('[') and ']' in line:
                if current_name and current_content:
                    story_text = ' '.join(current_content)
                    if len(story_text) > 30:
                        if not any(s['name'] == current_name for s in stories):
                            stories.append({
                                'name': current_name,
                                'content': story_text[:1500]
                            })
                try:
                    current_name = line[1:line.index(']')].strip()
                    if len(current_name) > 10:
                        current_name = ""
                except:
                    current_name = ""
                current_content = []
            elif current_name and line and not line.startswith('['):
                current_content.append(line)
        if current_name and current_content:
            story_text = ' '.join(current_content)
            if len(story_text) > 30:
                if not any(s['name'] == current_name for s in stories):
                    stories.append({
                        'name': current_name,
                        'content': story_text[:1500]
                    })
    if len(stories) < 3:
        print("数据不足，添加默认人物...")
        default_stories = [
            {'name': '毛胜利', 'content': '宣纸晒纸工艺传承人，坚守古法技艺，数十年如一日的坚持让人敬佩，传承之路虽然艰辛但成果丰硕'},
            {'name': '王亚蓉', 'content': '古丝绸修复专家，保护文化遗产，面对破损的文物充满心疼，修复成功后倍感自豪，为文化传承贡献巨大'},
            {'name': '周东红', 'content': '捞纸工，40年坚守，传承宣纸技艺，工作枯燥但充满热爱，看到年轻人传承技艺感到欣慰'},
            {'name': '孟剑锋', 'content': '錾刻大师，制作国礼，传承技艺，创作过程充满挑战，完成作品后充满成就感和荣誉感'},
            {'name': '单嘉玖', 'content': '古书画修复师，修复数百件文物，面对珍贵文物小心翼翼，修复成功后感到无比自豪和满足'},
            {'name': '李守白', 'content': '海派剪纸传承人，将传统剪纸与现代艺术结合，坚持创新传承，对传统文化充满热爱与责任感'},
            {'name': '张桂英', 'content': '苏绣大师，从事苏绣创作50余年，技艺精湛，耐心传授技艺给年轻人，坚守传统又不断突破'},
            {'name': '王津', 'content': '故宫钟表修复师，修复百年古钟，工作细致入微，对文物有深厚感情，充满使命感'},
            {'name': '陈巧生', 'content': '铜炉制作技艺传承人，复原古代铜炉制作工艺，数十年潜心研究，克服无数困难，成就感满满'},
            {'name': '杨玉芳', 'content': '皮影戏传承人，坚守皮影艺术数十年，面对行业困境不放弃，致力于传播皮影文化，责任感强烈'}
        ]
        for default in default_stories:
            if not any(s['name'] == default['name'] for s in stories):
                stories.append(default)
    return stories

def calculate_sentiment_scores(text):
    positive_keywords = ['自豪', '欣慰', '满足', '成就', '荣誉', '热爱', '喜悦', '骄傲', '敬佩', '成功']
    positive_count = sum(text.count(word) for word in positive_keywords)
    positive_score = min(positive_count / max(len(text)/500, 1) * 1.5, 0.95)
    persist_keywords = ['坚守', '坚持', '执着', '毅力', '数十年', '40年', '苦练', '刻苦', '专注', '恒心']
    persist_count = sum(text.count(word) for word in persist_keywords)
    persist_score = min(persist_count / max(len(text)/400, 1) * 1.3, 0.95)
    hardship_keywords = ['艰辛', '困难', '挑战', '艰苦', '不易', '枯燥', '繁琐', '耗时', '磨练', '压力']
    hardship_count = sum(text.count(word) for word in hardship_keywords)
    hardship_score = min(hardship_count / max(len(text)/450, 1) * 1.2, 0.95)
    duty_keywords = ['责任', '使命', '守护', '保护', '贡献', '传承', '弘扬', '担当', '义务', '奉献']
    duty_count = sum(text.count(word) for word in duty_keywords)
    duty_score = min(duty_count / max(len(text)/350, 1) * 1.4, 0.95)
    achievement_keywords = ['成就', '成果', '价值', '意义', '满足', '认可', '荣耀', '辉煌', '突破', '贡献']
    achievement_count = sum(text.count(word) for word in achievement_keywords)
    achievement_score = min(achievement_count / max(len(text)/425, 1) * 1.1, 0.95)
    scores = [positive_score, persist_score, hardship_score, duty_score, achievement_score]
    scores = [min(max(score, 0.5), 0.95) for score in scores]
    return scores

def create_culture_radar_chart(stories):
    print(f"开始为 {len(stories)} 位文化传承人创建情感维度雷达图...")
    dimensions = ['积极程度', '坚持程度', '艰辛程度', '责任感', '成就感']
    n_dim = len(dimensions)
    angles = np.linspace(0, 2 * np.pi, n_dim, endpoint=False).tolist()
    angles += angles[:1]
    fig = plt.figure(figsize=(16, 8), dpi=100, facecolor='#f8f9fa')
    
    # ---------------------- 左侧雷达图（清新配色） ----------------------
    ax1 = fig.add_subplot(121, polar=True)
    ax1.set_facecolor('#ffffff')
    
    # 清新配色方案（莫兰迪色系+低饱和度）
    colors = ['#6A994E', '#7209B7', '#F72585', '#4361EE', '#F77F00', '#06D6A0', 
              '#118AB2', '#073B4C', '#8D99AE', '#EF476F', '#FFD166', '#00A896',
              '#9381FF', '#F8A4D8', '#70D6FF', '#FF7AA2', '#80B3FF', '#B28DFF']
    linestyles = ['-', '--', '-.', ':', '-', '--', '-.', ':', '-', '--', '-.', ':',
                  '-', '--', '-.', ':', '-', '--']
    markers = ['o', 's', '^', 'D', 'v', '<', '>', 'p', '*', 'h', '+', 'x',
               'o', 's', '^', 'D', 'v', '<']
    
    all_scores = []
    for i, story in enumerate(stories):
        scores = calculate_sentiment_scores(story['content'])
        scores += scores[:1]
        all_scores.append(scores[:-1])
        color = colors[i % len(colors)]
        linestyle = linestyles[i % len(linestyles)]
        marker = markers[i % len(markers)]
        ax1.plot(angles, scores, marker=marker, linestyle=linestyle, linewidth=3,
                label=story['name'], color=color, markersize=9, alpha=0.85)
        ax1.fill(angles, scores, alpha=0.08, color=color)
    
    ax1.set_xticks(angles[:-1])
    ax1.set_xticklabels(dimensions, fontsize=12, fontweight='bold', color='#2D3047')
    ax1.set_ylim(0, 1.0)
    ax1.set_yticks([0.2, 0.4, 0.6, 0.8, 1.0])
    ax1.set_yticklabels(['0.2', '0.4', '0.6', '0.8', '1.0'], fontsize=10, color='#2D3047')
    ax1.set_title('文化传承人物情感维度对比', fontsize=14, fontweight='bold', pad=25, color='#2D3047')
    ax1.legend(loc='upper right', bbox_to_anchor=(1.4, 1.0), fontsize=10, 
              title='传承人', title_fontsize=12, ncol=2, framealpha=0.9, facecolor='#ffffff')
    ax1.grid(True, alpha=0.3, linestyle='--', color='#CED4DA')
    
    # ---------------------- 右侧雷达图（添加平均分+波动范围） ----------------------
    ax2 = fig.add_subplot(122, polar=True)
    ax2.set_facecolor('#ffffff')
    
    if all_scores:
        # 计算各维度平均分、最大值、最小值（波动范围）
        scores_array = np.array(all_scores)
        avg_scores = np.mean(scores_array, axis=0)
        max_scores = np.max(scores_array, axis=0)
        min_scores = np.min(scores_array, axis=0)
        
        # 准备雷达图数据（闭合）
        avg_radar = avg_scores.tolist() + avg_scores[:1].tolist()
        max_radar = max_scores.tolist() + max_scores[:1].tolist()
        min_radar = min_scores.tolist() + min_scores[:1].tolist()
        
        # 绘制波动范围（填充）
        ax2.fill_between(angles, min_radar, max_radar, alpha=0.15, color='#4361EE', label='分数波动范围')
        # 绘制平均分（实线）
        ax2.plot(angles, avg_radar, 'o-', linewidth=3.5, color='#4361EE', label='各维度平均分', markersize=10, markerfacecolor='#FFFFFF', markeredgewidth=2)
        # 绘制最大值/最小值参考线（虚线）
        ax2.plot(angles, max_radar, '--', linewidth=1.5, color='#F72585', alpha=0.7, label='维度最高分')
        ax2.plot(angles, min_radar, '-.', linewidth=1.5, color='#6A994E', alpha=0.7, label='维度最低分')
        
        # 添加数值标签（平均分+波动范围）
        for j, (angle, avg, max_val, min_val) in enumerate(zip(angles[:-1], avg_scores, max_scores, min_scores)):
            # 平均分标签
            ax2.text(angle, avg + 0.08, f'平均分\n{avg:.2f}', 
                    ha='center', va='center', fontsize=10, fontweight='bold', color='#2D3047',
                    bbox=dict(boxstyle='round,pad=0.3', facecolor='#FFFFFF', alpha=0.9, edgecolor='#4361EE'))
            # 波动范围标签
            ax2.text(angle, avg - 0.10, f'波动\n{min_val:.2f}-{max_val:.2f}', 
                    ha='center', va='center', fontsize=9, color='#6C757D',
                    bbox=dict(boxstyle='round,pad=0.2', facecolor='#F8F9FA', alpha=0.8))
    else:
        # 默认数据
        avg_scores = [0.85, 0.90, 0.75, 0.92, 0.88]
        max_scores = [0.95, 0.95, 0.85, 0.95, 0.95]
        min_scores = [0.70, 0.80, 0.65, 0.85, 0.75]
        avg_radar = avg_scores + avg_scores[:1]
        max_radar = max_scores + max_scores[:1]
        min_radar = min_scores + min_scores[:1]
        
        ax2.fill_between(angles, min_radar, max_radar, alpha=0.15, color='#4361EE', label='分数波动范围')
        ax2.plot(angles, avg_radar, 'o-', linewidth=3.5, color='#4361EE', label='各维度平均分', markersize=10, markerfacecolor='#FFFFFF', markeredgewidth=2)
        ax2.plot(angles, max_radar, '--', linewidth=1.5, color='#F72585', alpha=0.7, label='维度最高分')
        ax2.plot(angles, min_radar, '-.', linewidth=1.5, color='#6A994E', alpha=0.7, label='维度最低分')
        
        for j, (angle, avg, max_val, min_val) in enumerate(zip(angles[:-1], avg_scores, max_scores, min_scores)):
            ax2.text(angle, avg + 0.08, f'平均分\n{avg:.2f}', 
                    ha='center', va='center', fontsize=10, fontweight='bold', color='#2D3047',
                    bbox=dict(boxstyle='round,pad=0.3', facecolor='#FFFFFF', alpha=0.9, edgecolor='#4361EE'))
            ax2.text(angle, avg - 0.10, f'波动\n{min_val:.2f}-{max_val:.2f}', 
                    ha='center', va='center', fontsize=9, color='#6C757D',
                    bbox=dict(boxstyle='round,pad=0.2', facecolor='#F8F9FA', alpha=0.8))
    
    ax2.set_xticks(angles[:-1])
    ax2.set_xticklabels(dimensions, fontsize=12, fontweight='bold', color='#2D3047')
    ax2.set_ylim(0, 1.0)
    ax2.set_yticks([0.2, 0.4, 0.6, 0.8, 1.0])
    ax2.set_yticklabels(['0.2', '0.4', '0.6', '0.8', '1.0'], fontsize=10, color='#2D3047')
    ax2.set_title('文化传承情感维度统计（平均分+波动范围）', fontsize=14, fontweight='bold', pad=25, color='#2D3047')
    # 优化右侧图例
    ax2.legend(loc='upper right', bbox_to_anchor=(1.4, 1.0), fontsize=10, 
              title='统计维度', title_fontsize=12, framealpha=0.9, facecolor='#ffffff')
    ax2.grid(True, alpha=0.3, linestyle='--', color='#CED4DA')
    
    # 整体标题
    plt.suptitle('文化传承领域大国工匠情感维度雷达图', fontsize=16, fontweight='bold', y=0.98, color='#2D3047')
    plt.tight_layout()
    
    # 保存图片
    output_path = 'E:\\文化传承雷达图.png'
    plt.savefig(output_path, dpi=300, bbox_inches='tight', facecolor='#f8f9fa')
    print(f"\n✓ 雷达图已保存到: {output_path}")
    
    # 显示图片
    try:
        plt.show(block=True)
        print("✓ 雷达图显示成功")
    except Exception as e:
        print(f"显示图表时出错: {e}")
        print("请查看E盘根目录的PNG文件")
    return True

def analyze_culture_sentiment(text):
    positive_words = ['成功', '成就', '突破', '荣誉', '精湛', '完美', '高超', '卓越',
                     '传承', '保护', '弘扬', '复兴', '重现', '延续', '发扬', '创新',
                     '精美', '珍贵', '宝贵', '自豪', '骄傲', '热爱', '专注', '复兴',
                     '重要', '关键', '价值', '意义', '伟大', '辉煌', '荣耀', '胜利']
    negative_words = ['困难', '难题', '挑战', '艰苦', '不易', '失传', '断层', '消失',
                     '破损', '残破', '损坏', '衰退', '退化', '艰辛', '缺乏', '不足']
    text_lower = text.lower()
    pos_count = sum(text_lower.count(word) for word in positive_words)
    neg_count = sum(text_lower.count(word) for word in negative_words)
    if pos_count + neg_count > 0:
        score = pos_count / (pos_count + neg_count)
    else:
        score = 0.75
    return score, pos_count, neg_count

def main():
    print("正在分析文化传承文档...")
    file_path = 'E:\\文化传承.docx'
    if not os.path.exists(file_path):
        print(f"✗ 文件不存在: {file_path}")
        print("请确保 '文化传承.docx' 文件在E盘根目录")
        return
    text_content = read_docx_file(file_path)
    if not text_content or len(text_content) < 100:
        print("文件读取失败或内容太少，使用测试数据...")
        text_content = """[毛胜利]{.mark}宣纸是我国造纸技术皇冠上的明珠，制作有108道工序，80%以上纯手工，数十年坚守让人敬佩，虽然过程艰辛但充满成就感
[王亚蓉]{.mark}古丝绸修复专家，从泥糊中提取复制东周丝绸文物，对文物充满责任感，修复成功后倍感自豪，面对破损文物心生心疼
[周东红]{.mark}捞纸工，40年坚守，100%正品率，传承宣纸技艺，工作枯燥但充满热爱，看到年轻人传承感到欣慰
[孟剑锋]{.mark}錾刻大师，制作APEC国礼，传承錾刻技艺，创作过程充满挑战，完成作品后充满荣誉感和成就感
[单嘉玖]{.mark}古书画修复师，40年修复数百件文物，对文化遗产有强烈的使命感，修复成功后感到无比满足
[李守白]{.mark}海派剪纸传承人，将传统剪纸与现代艺术结合，坚持创新传承，对传统文化充满热爱与责任感，克服市场困境坚持创作
[张桂英]{.mark}苏绣大师，从事苏绣创作50余年，技艺精湛，耐心传授技艺给年轻人，坚守传统又不断突破，对苏绣艺术充满热爱
[王津]{.mark}故宫钟表修复师，修复百年古钟，工作细致入微，对文物有深厚感情，充满使命感，修复过程繁琐但成就感满满
[陈巧生]{.mark}铜炉制作技艺传承人，复原古代铜炉制作工艺，数十年潜心研究，克服无数技术困难，作品获得广泛认可，成就感十足
[杨玉芳]{.mark}皮影戏传承人，坚守皮影艺术数十年，面对行业困境不放弃，致力于传播皮影文化，责任感强烈，看到孩子们喜欢皮影感到欣慰"""
    stories = extract_stories_from_text(text_content)
    print(f"\n✓ 提取到 {len(stories)} 位文化传承人:")
    for i, story in enumerate(stories, 1):
        print(f"  {i:2d}. {story['name']:10s} - {len(story['content']):5d} 字符")
    results = []
    print("\n进行情感分析...")
    for story in stories:
        score, pos_count, neg_count = analyze_culture_sentiment(story['content'])
        sentiment = '积极' if score > 0.6 else '消极' if score < 0.4 else '中性'
        results.append({
            '姓名': story['name'],
            '情感倾向': sentiment,
            '综合分数': round(score, 3),
            '积极词数': pos_count,
            '消极词数': neg_count,
            '内容长度': len(story['content'])
        })
        print(f"  {story['name']:10s}: {sentiment} (分数: {score:.3f}, 积极词: {pos_count}, 消极词: {neg_count})")
    df = pd.DataFrame(results)
    csv_path = 'E:\\文化传承情感分析.csv'
    try:
        df.to_csv(csv_path, index=False, encoding='utf-8-sig')
        print(f"\n✓ 分析结果已保存: {csv_path}")
    except Exception as e:
        print(f"保存CSV文件失败: {e}")
    print("\n" + "=" * 60)
    print("正在生成雷达图...")
    print("=" * 60)
    create_culture_radar_chart(stories)
    print("\n" + "=" * 60)
    print("分析完成!")
    print("=" * 60)
    print(f"分析人数: {len(results)}")
    print(f"平均情感分数: {df['综合分数'].mean():.3f}")
    print("\n生成的文件:")
    radar_path = 'E:\\文化传承雷达图.png'
    csv_path = 'E:\\文化传承情感分析.csv'
    for path, desc in [(radar_path, "雷达图"), (csv_path, "分析数据")]:
        if os.path.exists(path):
            size = os.path.getsize(path)
            print(f"✓ {desc}: {path} ({size:,} 字节)")
        else:
            print(f"✗ {desc}: {path} (未生成)")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n程序出错: {e}")
        import traceback
        traceback.print_exc()
        print("\n尝试强制生成雷达图...")
        try:
            test_stories = [
                {'name': '毛胜利', 'content': '宣纸晒纸工艺传承人，坚守古法技艺，数十年如一日的坚持让人敬佩，传承之路虽然艰辛但成果丰硕'},
                {'name': '王亚蓉', 'content': '古丝绸修复专家，保护文化遗产，面对破损的文物充满心疼，修复成功后倍感自豪，为文化传承贡献巨大'},
                {'name': '周东红', 'content': '捞纸工，40年坚守，传承宣纸技艺，工作枯燥但充满热爱，看到年轻人传承技艺感到欣慰'},
                {'name': '孟剑锋', 'content': '錾刻大师，制作国礼，传承技艺，创作过程充满挑战，完成作品后充满成就感和荣誉感'},
                {'name': '单嘉玖', 'content': '古书画修复师，修复数百件文物，面对珍贵文物小心翼翼，修复成功后感到无比自豪和满足'},
                {'name': '李守白', 'content': '海派剪纸传承人，将传统剪纸与现代艺术结合，坚持创新传承，对传统文化充满热爱与责任感'},
                {'name': '张桂英', 'content': '苏绣大师，从事苏绣创作50余年，技艺精湛，耐心传授技艺给年轻人，坚守传统又不断突破'},
                {'name': '王津', 'content': '故宫钟表修复师，修复百年古钟，工作细致入微，对文物有深厚感情，充满使命感'},
                {'name': '陈巧生', 'content': '铜炉制作技艺传承人，复原古代铜炉制作工艺，数十年潜心研究，克服无数困难，成就感满满'},
                {'name': '杨玉芳', 'content': '皮影戏传承人，坚守皮影艺术数十年，面对行业困境不放弃，致力于传播皮影文化，责任感强烈'}
            ]
            create_culture_radar_chart(test_stories)
        except Exception as e2:
            print(f"强制生成也失败: {e2}")