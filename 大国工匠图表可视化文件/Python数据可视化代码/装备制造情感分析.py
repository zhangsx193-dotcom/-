import re
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('TkAgg')  # 强制使用TkAgg后端
import matplotlib.pyplot as plt
import os
from docx import Document
import warnings
warnings.filterwarnings('ignore')

# 设置工作目录（已按要求改为E:\）
os.chdir('E:\\')

# 设置中文字体（解决中文显示乱码问题）
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

print("=" * 60)
print("装备制造领域情感分析系统（修正版）")
print("工作目录:", os.getcwd())
print("=" * 60)

def read_docx_file(file_path):
    """读取docx文件内容，处理可能的读取异常"""
    try:
        doc = Document(file_path)
        full_text = []
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:  # 过滤空行
                full_text.append(text)
        return '\n'.join(full_text)
    except Exception as e:
        print(f"读取docx文件失败: {e}")
        return ""

def extract_stories(text):
    """提取文档中所有人物（核心+配角），适配文档表述格式"""
    stories = []
    lines = text.split('\n')
    # 文档中所有明确出现的人物名（核心+配角）
    all_names = [
        # 核心人物
        "韩超", "刘云清", "易冉", "张嘉", "张如意", 
        "潘从明", "罗昭强", "宁允展", "顾秋亮", "毛正石", 
        "戎鹏强", "张新停", "潘玉华",
        # 配角人物
        "杨卫东", "龚元龙", "刘彦冰", "董家会", "徐泽琨",
        "于文燕", "姜磊", "贺良", "王洪年", "许滨", "邹强"
    ]
    
    current_name = None
    current_content = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # 检查当前行是否为已知人物名的开始（支持“人名+冒号/逗号/空格”格式）
        name_match = next((name for name in all_names if 
                          line.startswith(f"{name}：") or 
                          line.startswith(f"{name},") or 
                          line.startswith(f"{name}，") or 
                          line == name), None)
        if name_match:
            # 保存上一个人物的故事（内容长度放宽至50字，适配配角简短描述）
            if current_name and len('\n'.join(current_content)) > 50:
                stories.append({
                    'name': current_name,
                    'content': '\n'.join(current_content)
                })
            # 初始化新人物的故事收集
            current_name = name_match
            current_content = []
            # 提取人物名后的内容（修复变量名错误：name→name_match）
            if line != current_name:
                if line.startswith(f"{name_match}："):
                    current_content.append(line.replace(f"{name_match}：", '').strip())
                elif line.startswith(f"{name_match},") or line.startswith(f"{name_match}，"):
                    current_content.append(line.replace(f"{name_match},", '').replace(f"{name_match}，", '').strip())
        elif current_name:
            # 收集当前人物的后续内容（遇到其他人物名则停止）
            next_name = next((name for name in all_names if 
                            line.startswith(f"{name}：") or 
                            line.startswith(f"{name},") or 
                            line.startswith(f"{name}，") or 
                            line == name), None)
            if not next_name:
                current_content.append(line)
    
    # 保存最后一个人物的故事
    if current_name and len('\n'.join(current_content)) > 50:
        stories.append({
            'name': current_name,
            'content': '\n'.join(current_content)
        })
    
    # 去重（避免同一人物被多次提取）
    unique_stories = []
    seen_names = set()
    for story in stories:
        if story['name'] not in seen_names:
            seen_names.add(story['name'])
            unique_stories.append(story)
    
    return unique_stories

def calculate_dimension_scores(text):
    """优化维度分数计算：扩充投入程度关键词，调整分数权重"""
    # 1. 技术难度维度（保持原有逻辑，补充行业特色关键词）
    tech_keywords = ['技术', '工艺', '操作', '设备', '机器', '系统', '精密', '精确', 
                    '高难度', '复杂', '精细', '难题', '挑战', '高级', '尖端', '高科技',
                    '微米', '毫米', '丝', '精度', '误差', '密封', '焊接', '研磨',
                    '深孔', '锻造', '铸造', '组装', '调试', '镗加工', '铅柱']
    tech_count = sum(text.count(word) for word in tech_keywords)
    tech_score = min(tech_count / (len(text)/3000 + 1) * 0.8, 0.95)
    
    # 2. 创新程度维度（保持原有逻辑）
    innov_keywords = ['创新', '创造', '发明', '研发', '改进', '优化', '首创', '独创',
                     '专利', '自主研发', '自主设计', '革新', '突破', '新方法', '新工艺']
    innov_count = sum(text.count(word) for word in innov_keywords)
    innov_score = min(innov_count / (len(text)/2500 + 1) * 0.9, 0.95)
    
    # 3. 投入程度维度（核心优化：扩充关键词，调整权重）
    commit_keywords = [
        # 原有抽象关键词
        '刻苦', '努力', '坚持', '奋斗', '钻研', '专注', '投入', '付出',
        '日夜', '加班', '训练', '练习', '苦练', '拼搏', '辛勤',
        # 新增场景化关键词（适配文档表述）
        '吃住厂', '通宵', '反复试验', '摸索', '自费', '啃资料', '查单词',
        '千百次', '上万次', '几年', '数十年', '毕生', '扎根', '坚守',
        '废寝忘食', '不辞辛劳', '克服困难', '迎难而上', '毫无保留'
    ]
    commit_count = sum(text.count(word) for word in commit_keywords)
    # 提高权重系数（从0.85→1.0），降低分母系数（从2000→1500），提升分数敏感度
    commit_score = min(commit_count / (len(text)/1500 + 1) * 1.0, 0.95)
    
    # 4. 成就高度维度（保持原有逻辑）
    achiev_keywords = ['成功', '成就', '突破', '荣誉', '奖励', '表彰', '完成', '实现',
                      '获奖', '冠军', '第一', '纪录', '成果', '胜利', '卓越', '辉煌',
                      '专家', '大师', '能手', '标兵', '先进', '首席', '顶尖', '领先']
    achiev_count = sum(text.count(word) for word in achiev_keywords)
    achiev_score = min(achiev_count / (len(text)/1500 + 1) * 1.0, 0.95)
    
    # 5. 影响广度维度（保持原有逻辑）
    impact_keywords = ['影响', '贡献', '价值', '意义', '重要', '关键', '推动', '促进',
                      '国际', '全球', '世界', '国家', '行业', '领域', '领先', '先进',
                      '出口', '自主知识产权', '大国重器', '海洋强国', '高铁名片']
    impact_count = sum(text.count(word) for word in impact_keywords)
    impact_score = min(impact_count / (len(text)/1800 + 1) * 0.9, 0.95)
    
    # 确保分数在0.2-0.95之间（提高最低分，避免过低）
    scores = [tech_score, innov_score, commit_score, achiev_score, impact_score]
    scores = [min(max(score, 0.2), 0.95) for score in scores]
    
    return scores

def create_radar_chart(stories):
    """创建雷达图：优化布局，支持更多人物显示"""
    if not stories:
        print("没有数据生成雷达图")
        return
    
    print(f"开始为 {len(stories)} 个人物生成雷达图...")
    
    # 雷达图基础配置
    dimensions = ['技术难度', '创新程度', '投入程度', '成就高度', '影响广度']
    n_dim = len(dimensions)
    angles = np.linspace(0, 2 * np.pi, n_dim, endpoint=False).tolist()
    angles += angles[:1]
    
    # 创建画布（优化尺寸，支持更多人物）
    fig = plt.figure(figsize=(16, 12))
    
    # 子图1：多人物情感维度对比（最多显示12人，优化图例布局）
    ax1 = fig.add_subplot(121, polar=True)
    # 扩充颜色列表，避免重复
    colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FECA57', 
              '#FF9FF3', '#54A0FF', '#5F27CD', '#FF7F50', '#6495ED',
              '#98FB98', '#DDA0DD', '#F0E68C', '#FFB6C1', '#87CEEB',
              '#D3D3D3', '#F4A460', '#20B2AA', '#9370DB', '#32CD32']
    
    # 遍历所有人物绘制（限制12人，避免图表拥挤）
    for i, story in enumerate(stories[:12]):
        name = story['name']
        scores = calculate_dimension_scores(story['content'])
        scores_closed = scores + scores[:1]
        
        ax1.plot(angles, scores_closed, 'o-', linewidth=2, 
                label=name, color=colors[i % len(colors)])
        ax1.fill(angles, scores_closed, alpha=0.1, color=colors[i % len(colors)])
    
    # 子图1样式优化：图例分行显示
    ax1.set_xticks(angles[:-1])
    ax1.set_xticklabels(dimensions, fontsize=12, fontweight='bold')
    ax1.set_ylim(0, 1.0)
    ax1.set_yticks([0.2, 0.4, 0.6, 0.8, 1.0])
    ax1.set_title('装备制造领域人物情感维度对比', fontsize=14, fontweight='bold', pad=20)
    ax1.legend(loc='upper right', bbox_to_anchor=(1.4, 1.0), fontsize=9, ncol=2)
    ax1.grid(True, alpha=0.3)
    
    # 子图2：平均情感维度分析（保持原有逻辑，优化数值标签）
    ax2 = fig.add_subplot(122, polar=True)
    
    all_scores = []
    for story in stories[:12]:
        all_scores.append(calculate_dimension_scores(story['content']))
    avg_scores = np.mean(all_scores, axis=0)
    std_scores = np.std(all_scores, axis=0)
    
    avg_closed = avg_scores.tolist() + avg_scores[:1].tolist()
    upper_bound = (avg_scores + std_scores/2).tolist() + (avg_scores[:1] + std_scores[:1]/2).tolist()
    lower_bound = (avg_scores - std_scores/2).tolist() + (avg_scores[:1] - std_scores[:1]/2).tolist()
    
    ax2.plot(angles, avg_closed, 'o-', linewidth=3, color='#2E86AB', label='各维度平均分')
    ax2.fill(angles, avg_closed, alpha=0.2, color='#2E86AB')
    ax2.fill_between(angles, lower_bound, upper_bound, alpha=0.1, color='#A23B72', label='分数波动范围')
    
    # 数值标签优化：位置调整，避免重叠
    for j, (angle, score) in enumerate(zip(angles[:-1], avg_scores)):
        ax2.text(angle, score + 0.03, f'{score:.2f}', ha='center', va='bottom', 
                fontsize=11, fontweight='bold', color='#2E86AB')
    
    ax2.set_xticks(angles[:-1])
    ax2.set_xticklabels(dimensions, fontsize=12, fontweight='bold')
    ax2.set_ylim(0, 1.0)
    ax2.set_yticks([0.2, 0.4, 0.6, 0.8, 1.0])
    ax2.set_title('装备制造领域平均情感维度分析', fontsize=14, fontweight='bold', pad=20)
    ax2.legend(loc='upper right', bbox_to_anchor=(1.3, 1.0), fontsize=10)
    ax2.grid(True, alpha=0.3)
    
    # 总标题优化
    plt.suptitle('装备制造领域大国工匠情感维度雷达图分析（修正版）', fontsize=16, fontweight='bold', y=0.98)
    plt.tight_layout()
    
    # 保存雷达图
    output_path = 'E:\\装备制造雷达图_修正版.png'
    plt.savefig(output_path, dpi=300, bbox_inches='tight', facecolor='white')
    print(f"✓ 雷达图已保存到: {output_path}")
    
    # 尝试显示图表
    try:
        plt.show(block=True)
        print("✓ 雷达图显示成功")
    except Exception as e:
        print(f"显示图表时出错: {e}")
        print(f"图表已保存到文件，请查看: {output_path}")

def analyze_sentiment_simple(text):
    """简单情感分析：保持原有逻辑，补充行业特色关键词"""
    positive_words = ['成功', '成就', '创新', '突破', '荣誉', '先进', '领先', '自豪',
                     '热爱', '专注', '解决', '创造', '贡献', '完美', '胜利', '优秀',
                     '卓越', '精湛', '高超', '冠军', '第一', '表彰', '能手', '大师',
                     '攻克', '完成', '实现', '节省', '高效', '精准', '优质', '好评',
                     '自主', '首创', '顶尖', '标杆', '示范', '里程碑']
    
    negative_words = ['困难', '难题', '问题', '失败', '错误', '缺陷', '故障', '挑战',
                     '头疼', '憋屈', '艰苦', '痛苦', '挫折', '障碍', '不足', '缺点',
                     '报废', '返工', '延误', '损坏', '复杂', '繁琐', '疲劳', '危险']
    
    text_lower = text.lower()
    pos_count = sum(text_lower.count(word.lower()) for word in positive_words)
    neg_count = sum(text_lower.count(word.lower()) for word in negative_words)
    
    if pos_count + neg_count > 0:
        score = pos_count / (pos_count + neg_count)
    else:
        score = 0.5
    
    return score, pos_count, neg_count

def main():
    """主程序：串联所有功能"""
    print("正在启动装备制造领域情感分析（修正版）...")
    
    # 1. 检查文件
    file_path = 'E:\\装备制造.docx'
    if not os.path.exists(file_path):
        print(f"错误: 文件不存在 - {file_path}")
        print("请确保docx文件放置在E盘根目录，并命名为'装备制造.docx'")
        return
    
    # 2. 读取文件
    content = read_docx_file(file_path)
    if len(content) < 100:
        print("无法读取有效文件内容（内容过短或为空）")
        return
    print(f"✓ 文档读取成功，总字符数: {len(content):,}")
    
    # 3. 提取人物故事（修正后：包含所有核心+配角）
    stories = extract_stories(content)
    if not stories:
        print("未提取到有效人物故事，无法进行后续分析")
        return
    print(f"✓ 提取到 {len(stories)} 个人物故事")
    
    # 显示所有提取的人物
    print("\n提取的人物列表：")
    for i, story in enumerate(stories, 1):
        content_len = len(story['content'])
        print(f"  {i:2d}. {story['name']:12s} - 内容长度: {content_len:6d} 字符")
    
    # 4. 情感分析
    results = []
    print("\n" + "-" * 50)
    print("正在进行情感分析...")
    print("-" * 50)
    
    for story in stories:
        score, pos_count, neg_count = analyze_sentiment_simple(story['content'])
        if score > 0.6:
            sentiment = '积极'
        elif score < 0.4:
            sentiment = '消极'
        else:
            sentiment = '中性'
        
        results.append({
            '姓名': story['name'],
            '情感倾向': sentiment,
            '情感分数': round(score, 3),
            '积极关键词数': pos_count,
            '消极关键词数': neg_count,
            '故事内容长度': len(story['content'])
        })
        
        print(f"  {story['name']:12s}: 情感={sentiment:4s} | 分数={score:.3f} | "
              f"积极词={pos_count:2d} | 消极词={neg_count:2d}")
    
    # 5. 保存结果到CSV
    df_results = pd.DataFrame(results)
    csv_path = 'E:\\装备制造情感分析结果_修正版.csv'
    df_results.to_csv(csv_path, index=False, encoding='utf-8-sig')
    print(f"\n✓ 情感分析结果已保存到: {csv_path}")
    
    # 6. 生成雷达图
    print("\n" + "=" * 60)
    print("正在生成情感维度雷达图（修正版）...")
    print("=" * 60)
    create_radar_chart(stories)
    
    # 7. 分析总结
    print("\n" + "=" * 60)
    print("装备制造领域情感分析总结（修正版）")
    print("=" * 60)
    total_count = len(results)
    positive_count = len([r for r in results if r['情感倾向'] == '积极'])
    neutral_count = len([r for r in results if r['情感倾向'] == '中性'])
    negative_count = len([r for r in results if r['情感倾向'] == '消极'])
    avg_score = round(df_results['情感分数'].mean(), 3)
    
    print(f"分析总人数: {total_count} 人")
    print(f"积极情感人数: {positive_count} 人 ({positive_count/total_count*100:.1f}%)")
    print(f"中性情感人数: {neutral_count} 人 ({neutral_count/total_count*100:.1f}%)")
    print(f"消极情感人数: {negative_count} 人 ({negative_count/total_count*100:.1f}%)")
    print(f"平均情感分数: {avg_score}")
    
    # 8. 文件检查
    print("\n生成文件检查：")
    generated_files = [
        (csv_path, "情感分析结果CSV文件"),
        ("E:\\装备制造雷达图_修正版.png", "情感维度雷达图")
    ]
    for file_path, file_desc in generated_files:
        if os.path.exists(file_path):
            file_size = os.path.getsize(file_path) / 1024
            print(f"✓ {file_desc}: 存在（大小: {file_size:.1f} KB）")
        else:
            print(f"✗ {file_desc}: 未生成，请检查路径或权限")
    
    print("\n" + "=" * 60)
    print("装备制造情感分析_修正版.py 执行完成！")
    print("=" * 60)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n程序执行出错: {e}")
        import traceback
        traceback.print_exc()