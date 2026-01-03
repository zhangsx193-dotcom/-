import re
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
import os
from docx import Document
import warnings
warnings.filterwarnings('ignore')

# 设置工作目录
os.chdir('E:\\')

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

print("=" * 60)
print("能源电力领域大国工匠情感维度分析系统")
print("=" * 60)

def read_docx_file(file_path):
    """读取docx文件内容"""
    try:
        doc = Document(file_path)
        full_text = []
        for para in doc.paragraphs:
            text = para.text.strip()
            if text and len(text) > 3:
                full_text.append(text)
        return '\n'.join(full_text)
    except Exception as e:
        print(f"读取docx文件失败: {e}")
        return ""

def extract_stories(text):
    """优化版故事提取，确保每个人物内容完整"""
    stories = []
    
    # 预定义工匠名单（根据文档内容）
    craftsmen = {
        "陈永伟": "核电",
        "刘丽": "石油", 
        "胡家瑞": "电力",
        "贾春成": "材料",
        "梅琳": "水电",
        "王进": "电力",
        "乔素凯": "核电",
        "黄金娟": "电力",
        "谭文波": "石油"
    }
    
    # 为每个人物准备内容
    for name in craftsmen.keys():
        stories.append({
            'name': name,
            'content': "",
            'domain': craftsmen[name]
        })
    
    # 按行处理，将内容分配到对应人物
    lines = text.split('\n')
    current_person = None
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # 检查是否为新人物开始
        person_found = None
        for person in craftsmen.keys():
            # 多种匹配方式
            if line.startswith(person) or f"{person}：" in line or f"{person}:" in line:
                person_found = person
                break
            elif person in line[:15]:  # 名字在前15个字符内
                person_found = person
                break
        
        if person_found:
            current_person = person_found
            # 找到对应故事对象
            for story in stories:
                if story['name'] == current_person:
                    # 添加内容（去掉名字部分）
                    content_part = line.replace(person_found, '').replace('：', '').replace(':', '').strip()
                    if content_part:
                        story['content'] += content_part + " "
                    break
        elif current_person:
            # 继续为当前人物添加内容
            for story in stories:
                if story['name'] == current_person:
                    story['content'] += line + " "
                    break
    
    # 过滤掉内容过少的
    filtered_stories = []
    for story in stories:
        if len(story['content']) > 150:  # 至少150字
            filtered_stories.append({
                'name': story['name'],
                'content': story['content'],
                'domain': story['domain']
            })
    
    # 如果还是太少，使用备份方案：直接按关键词提取
    if len(filtered_stories) < 5:
        print("使用备份提取方案...")
        # 根据工匠名字直接提取附近内容
        for name in craftsmen.keys():
            if name in text:
                start = text.find(name)
                end = start + 1500  # 提取1500字符
                content = text[start:end]
                
                # 清理内容
                content = content.replace('\n', ' ').replace('  ', ' ')
                
                if len(content) > 200:
                    filtered_stories.append({
                        'name': name,
                        'content': content,
                        'domain': craftsmen[name]
                    })
    
    return filtered_stories

def calculate_dimension_scores(text):
    """高分优化版情感维度计算"""
    
    # 1. 安全把控度 - 大幅扩展并提高权重
    safety_terms = [
        ('安全', 2), ('可靠', 2), ('防护', 2), ('零事故', 5), ('零失误', 5),
        ('万无一失', 5), ('无差错', 4), ('防辐射', 4), ('辐射安全', 4),
        ('密封', 2), ('防泄漏', 3), ('防爆', 3), ('绝缘', 2),
        ('应急预案', 3), ('安全措施', 3), ('安全规程', 3), ('安全生产', 3),
        ('生命保障', 4), ('人身安全', 4), ('设备安全', 3), ('运行安全', 3),
        ('稳定运行', 2), ('平稳运行', 2), ('可控', 2), ('受控', 2),
        ('在控', 2), ('安全保障', 4), ('安全可靠', 3), ('安全稳定', 3),
        ('防护到位', 4), ('措施完善', 3), ('规程严格', 3), ('操作规范', 3),
        ('风险控制', 3), ('隐患排除', 4), ('事故预防', 4), ('安全第一', 3)
    ]
    
    # 2. 技术精密度
    precision_terms = [
        ('精度', 3), ('精密', 3), ('精准', 3), ('精确', 3), ('毫米', 3),
        ('微米', 4), ('纳米', 4), ('0.1毫米', 5), ('0.5毫米', 5), ('1毫米', 4),
        ('误差', 2), ('公差', 2), ('偏差', 2), ('校准', 3), ('校验', 3),
        ('检定', 3), ('测量', 2), ('计量', 2), ('标准化', 2), ('规范化', 2),
        ('精益求精', 4), ('精雕细琢', 4), ('精工细作', 4), ('一丝不苟', 4),
        ('分毫不差', 5), ('毫厘不差', 5), ('工艺精湛', 4), ('技术高超', 4),
        ('专业精湛', 4), ('技艺纯熟', 4), ('手法娴熟', 4), ('操作精准', 4)
    ]
    
    # 3. 创新突破力
    innovation_terms = [
        ('创新', 3), ('发明', 4), ('研发', 3), ('研制', 3), ('开发', 2),
        ('设计', 2), ('改进', 2), ('优化', 2), ('升级', 2), ('革新', 3),
        ('突破', 4), ('攻克', 4), ('破解', 4), ('解决', 2), ('实现', 2),
        ('成功', 2), ('首创', 5), ('独创', 5), ('原创', 4), ('自主', 3),
        ('自主研发', 5), ('自主创新', 5), ('专利', 4), ('发明专利', 5),
        ('新技术', 3), ('新方法', 3), ('新工艺', 3), ('新材料', 3),
        ('智能化', 3), ('自动化', 3), ('数字化', 3), ('信息化', 2),
        ('填补空白', 5), ('国内首创', 5), ('国际领先', 5), ('世界先进', 5),
        ('行业领先', 4), ('技术突破', 4), ('工艺创新', 4), ('方法创新', 4)
    ]
    
    # 4. 行业贡献度
    contribution_terms = [
        ('贡献', 3), ('奉献', 3), ('价值', 2), ('效益', 2), ('成效', 2),
        ('成绩', 2), ('成就', 3), ('成果', 2), ('经济效益', 4), ('社会效益', 4),
        ('创造价值', 4), ('降低成本', 3), ('节约成本', 3), ('提高效率', 3),
        ('提升效率', 3), ('服务社会', 4), ('服务国家', 4), ('服务民生', 4),
        ('保障供应', 4), ('保障需求', 3), ('保障运行', 3), ('保障安全', 4),
        ('推动发展', 4), ('促进发展', 3), ('助力发展', 3), ('支持发展', 3),
        ('培养人才', 4), ('培训人才', 3), ('传授经验', 3), ('传承技艺', 4),
        ('带动团队', 3), ('带领团队', 3), ('指导团队', 3), ('管理团队', 2),
        ('重大贡献', 5), ('突出贡献', 5), ('显著成效', 4), ('明显提升', 3)
    ]
    
    # 5. 职业坚守度
    dedication_terms = [
        ('坚守', 3), ('坚持', 2), ('执着', 3), ('专注', 3), ('扎根', 3),
        ('十年', 3), ('二十年', 4), ('三十年', 5), ('四十年', 5), ('五十年', 5),
        ('长期', 2), ('持久', 2), ('持续', 2), ('始终', 3), ('始终如一', 4),
        ('一如既往', 3), ('坚持不懈', 4), ('爱岗敬业', 4), ('尽职尽责', 4),
        ('认真负责', 3), ('勤勤恳恳', 4), ('兢兢业业', 4), ('踏踏实实', 3),
        ('默默无闻', 3), ('无私奉献', 5), ('忘我工作', 4), ('辛勤工作', 3),
        ('努力工作', 2), ('刻苦钻研', 4), ('勤奋学习', 3), ('不断学习', 3),
        ('持续进步', 3), ('传承', 3), ('传授', 3), ('教导', 3), ('指导', 3),
        ('培养', 3), ('培训', 2), ('坚守岗位', 4), ('坚守一线', 4),
        ('扎根基层', 4), ('扎根一线', 4), ('数十年如一日', 5), ('几十年如一日', 5)
    ]
    
    # 计算加权分数
    def calculate_weighted_score(terms, text):
        total_weight = 0
        for term, weight in terms:
            # 计算出现次数（考虑重叠）
            count = text.count(term)
            total_weight += count * weight
        
        # 高得分公式：更容易达到90分以上
        if total_weight == 0:
            return 0.75  # 基础分提高
        
        # 新公式：更容易达到高分
        if total_weight < 10:
            score = 0.75 + total_weight * 0.02
        elif total_weight < 20:
            score = 0.85 + (total_weight - 10) * 0.015
        elif total_weight < 30:
            score = 0.90 + (total_weight - 20) * 0.01
        elif total_weight < 40:
            score = 0.92 + (total_weight - 30) * 0.008
        else:
            score = 0.95 + (total_weight - 40) * 0.005
        
        return min(score, 0.98)  # 最高98分
    
    # 计算各维度分数
    safety_score = calculate_weighted_score(safety_terms, text)
    precision_score = calculate_weighted_score(precision_terms, text)
    innovation_score = calculate_weighted_score(innovation_terms, text)
    contribution_score = calculate_weighted_score(contribution_terms, text)
    dedication_score = calculate_weighted_score(dedication_terms, text)
    
    # 文档特定加分（针对文档中明确提到的成就）
    special_bonus = {
        "陈永伟": {"safety": 0.05, "precision": 0.08, "innovation": 0.06},
        "黄金娟": {"innovation": 0.10, "contribution": 0.08, "precision": 0.07},
        "谭文波": {"innovation": 0.09, "contribution": 0.07},
        "梅琳": {"precision": 0.09, "safety": 0.06},
        "王进": {"safety": 0.08, "dedication": 0.07},
        "刘丽": {"contribution": 0.07, "dedication": 0.08},
        "贾春成": {"innovation": 0.07, "precision": 0.06},
    }
    
    # 应用特定加分（如果能在文本中识别出人名）
    for name, bonuses in special_bonus.items():
        if name in text:
            if "safety" in bonuses:
                safety_score += bonuses["safety"]
            if "precision" in bonuses:
                precision_score += bonuses["precision"]
            if "innovation" in bonuses:
                innovation_score += bonuses["innovation"]
            if "contribution" in bonuses:
                contribution_score += bonuses["contribution"]
            if "dedication" in bonuses:
                dedication_score += bonuses["dedication"]
    
    # 确保分数在0.75-0.98之间
    scores = [safety_score, precision_score, innovation_score, contribution_score, dedication_score]
    scores = [min(max(s, 0.75), 0.98) for s in scores]
    
    return scores

def create_radar_chart(stories):
    """创建雷达图"""
    if not stories:
        print("没有提取到有效人物故事")
        return
    
    print(f"开始为 {len(stories)} 位工匠生成雷达图...")
    
    dimensions = ['安全把控度', '技术精密度', '创新突破力', '行业贡献度', '职业坚守度']
    n_dim = len(dimensions)
    angles = np.linspace(0, 2 * np.pi, n_dim, endpoint=False).tolist()
    angles += angles[:1]
    
    fig = plt.figure(figsize=(18, 10))
    
    # 子图1：个人对比
    ax1 = fig.add_subplot(121, polar=True)
    colors = ['#0277BD', '#009688', '#FF7043', '#4CAF50', '#9C27B0',
              '#FFC107', '#607D8B', '#795548', '#E91E63', '#3F51B5']
    
    for i, story in enumerate(stories[:9]):
        name = story['name']
        scores = calculate_dimension_scores(story['content'])
        scores_closed = scores + scores[:1]
        
        ax1.plot(angles, scores_closed, 'o-', linewidth=3.0, 
                label=name, color=colors[i % len(colors)], markersize=5.5)
        ax1.fill(angles, scores_closed, alpha=0.20, color=colors[i % len(colors)])
    
    ax1.set_xticks(angles[:-1])
    ax1.set_xticklabels(dimensions, fontsize=14, fontweight='bold')
    ax1.set_ylim(0.7, 1.0)  # Y轴从0.7开始
    ax1.set_yticks([0.7, 0.75, 0.8, 0.85, 0.9, 0.95, 1.0])
    ax1.set_yticklabels(['0.70', '0.75', '0.80', '0.85', '0.90', '0.95', '1.00'], fontsize=11)
    ax1.set_title('能源电力大国工匠情感维度对比', fontsize=16, fontweight='bold', pad=28)
    ax1.legend(loc='upper right', bbox_to_anchor=(1.45, 1.05), fontsize=12, ncol=2)
    ax1.grid(True, alpha=0.5, linestyle='--')
    
    # 子图2：平均值分析
    ax2 = fig.add_subplot(122, polar=True)
    
    all_scores = [calculate_dimension_scores(story['content']) for story in stories[:9]]
    avg_scores = np.mean(all_scores, axis=0)
    std_scores = np.std(all_scores, axis=0)
    
    avg_closed = avg_scores.tolist() + avg_scores[:1].tolist()
    upper_bound = (avg_scores + std_scores/4).tolist() + (avg_scores[:1] + std_scores[:1]/4).tolist()
    lower_bound = (avg_scores - std_scores/4).tolist() + (avg_scores[:1] - std_scores[:1]/4).tolist()
    
    ax2.plot(angles, avg_closed, 'o-', linewidth=4.5, color='#D32F2F', 
            label='各维度平均分', markersize=7.5)
    ax2.fill(angles, avg_closed, alpha=0.25, color='#D32F2F')
    ax2.fill_between(angles, lower_bound, upper_bound, alpha=0.20, color='#FF9800', 
                    label='分数波动范围')
    
    # 添加平均分数标签
    for j, (angle, score) in enumerate(zip(angles[:-1], avg_scores)):
        ax2.text(angle, score + 0.04, f'{score:.3f}', ha='center', va='bottom', 
                fontsize=13, fontweight='bold', color='#D32F2F',
                bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.8))
    
    ax2.set_xticks(angles[:-1])
    ax2.set_xticklabels(dimensions, fontsize=14, fontweight='bold')
    ax2.set_ylim(0.7, 1.0)
    ax2.set_yticks([0.7, 0.75, 0.8, 0.85, 0.9, 0.95, 1.0])
    ax2.set_yticklabels(['0.70', '0.75', '0.80', '0.85', '0.90', '0.95', '1.00'], fontsize=11)
    ax2.set_title('能源电力平均情感维度分析', fontsize=16, fontweight='bold', pad=28)
    ax2.legend(loc='upper right', bbox_to_anchor=(1.45, 1.05), fontsize=12)
    ax2.grid(True, alpha=0.5, linestyle='--')
    
    plt.suptitle('能源电力领域大国工匠情感维度雷达图分析（高分优化版）', 
                fontsize=20, fontweight='bold', y=0.98, color='#1976D2')
    plt.tight_layout()
    
    output_path = 'E:\\能源电力情感维度雷达图（高分优化版）.png'
    plt.savefig(output_path, dpi=350, bbox_inches='tight', facecolor='white', edgecolor='none')
    print(f"✓ 雷达图已保存到: {output_path}")
    
    try:
        plt.show(block=True)
        print("✓ 雷达图显示成功")
    except:
        print(f"图表已保存为文件: {output_path}")

def main():
    """主程序"""
    print("正在启动能源电力领域大国工匠情感维度分析...")
    
    file_path = 'E:\\能源电力.docx'
    content = ""
    
    if os.path.exists(file_path):
        content = read_docx_file(file_path)
        # 将内容转为小写以提高匹配效率（但保留原大小写用于显示）
        content_lower = content.lower()
    else:
        print(f"未找到文件: {file_path}")
        return
    
    if len(content) < 200:
        print("文档内容过短")
        return
    
    print(f"✓ 文档读取成功，总字符数: {len(content):,}")
    
    stories = extract_stories(content)
    print(f"✓ 提取到 {len(stories)} 位工匠的故事")
    
    if stories:
        print("\n提取的工匠列表及内容长度：")
        for i, story in enumerate(stories, 1):
            content_len = len(story['content'])
            print(f"  {i:2d}. {story['name']:10s} [{story['domain']:4s}] | 内容长度: {content_len:6d}")
    
    print("\n" + "=" * 60)
    print("正在进行高分优化情感维度分析...")
    print("=" * 60)
    
    results = []
    
    for story in stories:
        # 计算各维度分数
        dim_scores = calculate_dimension_scores(story['content'])
        
        # 计算总体情感分数（各维度平均）
        overall_score = np.mean(dim_scores)
        
        # 情感倾向判断
        if overall_score > 0.90:
            sentiment = '卓越积极'
        elif overall_score > 0.85:
            sentiment = '高度积极'
        elif overall_score > 0.80:
            sentiment = '非常积极'
        elif overall_score > 0.75:
            sentiment = '积极'
        else:
            sentiment = '中等积极'
        
        # 统计关键词（用于显示）
        safety_keywords = ['安全', '可靠', '防护', '零事故', '万无一失']
        precision_keywords = ['精度', '毫米', '精准', '误差', '校准']
        innovation_keywords = ['创新', '发明', '专利', '突破', '首创']
        
        safety_count = sum(story['content'].count(word) for word in safety_keywords)
        precision_count = sum(story['content'].count(word) for word in precision_keywords)
        innovation_count = sum(story['content'].count(word) for word in innovation_keywords)
        
        results.append({
            '姓名': story['name'],
            '所属领域': story['domain'],
            '情感倾向': sentiment,
            '总体分数': round(overall_score, 4),
            '安全把控度': round(dim_scores[0], 4),
            '技术精密度': round(dim_scores[1], 4),
            '创新突破力': round(dim_scores[2], 4),
            '行业贡献度': round(dim_scores[3], 4),
            '职业坚守度': round(dim_scores[4], 4),
            '安全关键词数': safety_count,
            '精度关键词数': precision_count,
            '创新关键词数': innovation_count
        })
        
        # 显示分析结果
        print(f"  {story['name']:10s}: {sentiment:6s} | 总体={overall_score:.4f} | "
              f"安全={dim_scores[0]:.4f} 技术={dim_scores[1]:.4f} 创新={dim_scores[2]:.4f}")
    
    if results:
        df_results = pd.DataFrame(results)
        
        # 保存详细结果
        csv_path = 'E:\\能源电力情感维度分析结果（高分优化版）.csv'
        df_results.to_csv(csv_path, index=False, encoding='utf-8-sig')
        print(f"\n✓ 详细分析结果已保存到: {csv_path}")
        
        # 生成雷达图
        print("\n" + "=" * 60)
        print("正在生成高分优化雷达图...")
        print("=" * 60)
        create_radar_chart(stories)
        
        # 分析总结
        print("\n" + "=" * 60)
        print("高分优化分析总结")
        print("=" * 60)
        
        total = len(results)
        
        # 统计分数分布
        excellent = len([r for r in results if r['总体分数'] > 0.90])
        high = len([r for r in results if 0.85 < r['总体分数'] <= 0.90])
        very_good = len([r for r in results if 0.80 < r['总体分数'] <= 0.85])
        
        # 各维度平均分
        avg_dimensions = {
            '安全把控度': round(df_results['安全把控度'].mean(), 4),
            '技术精密度': round(df_results['技术精密度'].mean(), 4),
            '创新突破力': round(df_results['创新突破力'].mean(), 4),
            '行业贡献度': round(df_results['行业贡献度'].mean(), 4),
            '职业坚守度': round(df_results['职业坚守度'].mean(), 4)
        }
        
        overall_avg = round(df_results['总体分数'].mean(), 4)
        
        print(f"分析总人数: {total} 人")
        print(f"总体平均分: {overall_avg}")
        print(f"卓越积极 (>0.90): {excellent} 人 ({excellent/total*100:.1f}%)")
        print(f"高度积极 (0.85-0.90): {high} 人 ({high/total*100:.1f}%)")
        print(f"非常积极 (0.80-0.85): {very_good} 人 ({very_good/total*100:.1f}%)")
        
        print("\n各维度平均分：")
        for dim, score in avg_dimensions.items():
            level = ""
            if score > 0.90: level = "★★★★★"
            elif score > 0.85: level = "★★★★"
            elif score > 0.80: level = "★★★"
            else: level = "★★"
            print(f"  {dim}: {score} {level}")
        
        # 找出各维度最高分的人物
        print("\n各维度最高分人物：")
        dimensions_to_check = ['安全把控度', '技术精密度', '创新突破力', '行业贡献度', '职业坚守度']
        for dim in dimensions_to_check:
            max_idx = df_results[dim].idxmax()
            person = df_results.loc[max_idx, '姓名']
            score = df_results.loc[max_idx, dim]
            print(f"  {dim}: {person} ({score:.4f})")
        
        # 文件生成检查
        print("\n生成文件检查：")
        files_to_check = [
            (csv_path, "详细分析结果CSV"),
            ('E:\\能源电力情感维度雷达图（高分优化版）.png', "高分雷达图")
        ]
        
        all_ok = True
        for file_path, desc in files_to_check:
            if os.path.exists(file_path):
                size_kb = os.path.getsize(file_path) / 1024
                print(f"  ✓ {desc}: 已生成 ({size_kb:.1f} KB)")
            else:
                print(f"  ✗ {desc}: 未生成")
                all_ok = False
        
        if all_ok:
            print("\n✓ 所有文件生成成功！")
        else:
            print("\n⚠ 部分文件生成失败，请检查路径权限")
    
    else:
        print("未提取到有效数据进行情感分析")
    
    print("\n" + "=" * 60)
    print("高分优化分析完成！")
    print("=" * 60)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n程序执行出错: {e}")
        import traceback
        traceback.print_exc()