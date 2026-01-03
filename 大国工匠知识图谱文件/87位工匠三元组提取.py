import pandas as pd

# ========== 核心配置 ==========
EXCEL_FILE = '工作簿1.xlsx'  # 确保Excel与此脚本在同一文件夹
SHEET_NAME = '原始数据'

# ========== 三元组抽取函数 ==========
def extract_triples_from_table(row):
    triples = []
    name = str(row['工匠姓名']).strip()
    
    # 跳过表头和空行
    if name == '工匠姓名' or name == 'nan' or not name:
        return triples
    
    # 1. 工匠-从事-职业
    if pd.notna(row['职业/行业']):
        triples.append((name, '从事', str(row['职业/行业']).strip()))
    
    # 2. 工匠-掌握-技术
    if pd.notna(row['核心技术/绝活']):
        tech = str(row['核心技术/绝活']).strip()
        triples.append((name, '掌握', tech))
        # 技术-应用于-行业
        if pd.notna(row['行业大类']):
            triples.append((tech, '应用于', str(row['行业大类']).strip()))
    
    # 3. 工匠-体现-精神特质（三列）
    for i in range(1, 4):
        trait = str(row.get(f'精神特质{i}', '')).strip()
        if trait and trait != 'nan':
            triples.append((name, '体现', trait))
            # 精神特质-属于-行业
            if pd.notna(row['行业大类']):
                triples.append((trait, '属于', str(row['行业大类']).strip()))
    
    # 4. 工匠-属于-行业大类
    if pd.notna(row['行业大类']):
        triples.append((name, '属于', str(row['行业大类']).strip()))
    
    # 5. 工匠-采用-创新类型
    if pd.notna(row['创新类型']):
        triples.append((name, '采用', str(row['创新类型']).strip()))
    
    # 6. 工匠-使用-传承方式
    if pd.notna(row['传承方式']):
        triples.append((name, '使用', str(row['传承方式']).strip()))
    
    # 7. 工匠-风险等级-值
    if pd.notna(row['风险等级']):
        triples.append((name, '风险等级', str(row['风险等级']).strip()))
    
    # 8. 工匠-精度等级-值
    if pd.notna(row['精度等级']):
        triples.append((name, '精度等级', str(row['精度等级']).strip()))
    
    return triples

# ========== 主流程 ==========
if __name__ == '__main__':
    try:
        # 读取Excel
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
        print(f"✅ 成功读取Excel，共{len(df)}行数据")
        
        # 抽取三元组
        all_triples = []
        for idx, row in df.iterrows():
            all_triples.extend(extract_triples_from_table(row))
        
        # 去重并转为DataFrame
        triples_df = pd.DataFrame(all_triples, columns=['Subject', 'Predicate', 'Object'])
        triples_df = triples_df.drop_duplicates()
        
        # 保存三元组
        triples_df.to_csv('大国工匠_triples.csv', index=False, encoding='utf-8-sig')
        print(f"📊 已保存{len(triples_df)}个三元组到'大国工匠_triples.csv'")
        
        # ========== 生成Gephi文件 ==========
        # 创建节点
        all_entities = set(triples_df['Subject']) | set(triples_df['Object'])
        
        # 自动识别类型
        entity_types = {}
        industry_set = set(df['行业大类'].dropna().unique())
        innovation_set = {'改进型', '突破性', '传承型'}
        heritage_set = {'师徒', '自学', '院校'}
        level_set = {'低', '中', '高'}
        
        for e in all_entities:
            if e in df['工匠姓名'].values:
                entity_types[e] = '工匠'
            elif e in industry_set:
                entity_types[e] = '行业大类'
            elif e in innovation_set:
                entity_types[e] = '创新类型'
            elif e in heritage_set:
                entity_types[e] = '传承方式'
            elif e in level_set:
                entity_types[e] = '等级'
            elif any(keyword in str(e) for keyword in ['mm', 'μm', '加工', '精度', '焊接', '研磨']):
                entity_types[e] = '核心技术'
            elif any(trait in str(e) for trait in ['精益求精', '坚守执着', '创新突破', '责任担当', '传承奉献', '问题导向']):
                entity_types[e] = '精神特质'
            else:
                entity_types[e] = '其他'
        
        nodes_df = pd.DataFrame([
            {'Id': i, 'Label': e, 'Type': entity_types[e]}
            for i, e in enumerate(all_entities)
        ])
        
        # 创建边
        entity_to_id = {row['Label']: row['Id'] for _, row in nodes_df.iterrows()}
        edges_df = pd.DataFrame([
            {
                'Source': entity_to_id[row['Subject']],
                'Target': entity_to_id[row['Object']],
                'Type': row['Predicate'],
                'Weight': 1
            }
            for _, row in triples_df.iterrows()
        ])
        
        # 保存Gephi文件
        nodes_df.to_csv('gephi_nodes.csv', index=False, encoding='utf-8-sig')
        edges_df.to_csv('gephi_edges.csv', index=False, encoding='utf-8-sig')
        
        print(f"🎯 节点文件: gephi_nodes.csv ({len(nodes_df)}个节点)")
        print(f"🔗 边文件: gephi_edges.csv ({len(edges_df)}条边)")
        print("\n" + "="*50)
        
    except FileNotFoundError:
        print(f"❌ 错误: 找不到文件'{EXCEL_FILE}'，请确认文件路径！")
    except Exception as e:
        print(f"❌ 发生错误: {e}")
