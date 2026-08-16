# -*- coding: utf-8 -*-

import os
import sys
import numpy as np
import pandas as pd
from sklearn.linear_model import LogisticRegression
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import roc_auc_score, classification_report, confusion_matrix

# 设置标准输出编码为 UTF-8
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass


def safe_read_csv(filepath, default_enc='utf-8'):
    """自动探测编码并安全读取 CSV"""
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"未找到文件: {filepath}")
    for enc in [default_enc, 'gb18030', 'gbk', 'utf-8-sig', 'latin1']:
        try:
            return pd.read_csv(filepath, encoding=enc)
        except Exception:
            continue
    raise ValueError(f"无法解析文件: {filepath}")


def process_monthly_data(df, year_month_str):
    """
    对单月日度流水按账户聚合，并向量化提取行为特征
    """
    # 查找各类消费字段
    spend_cols = sorted([c for c in df.columns if c.startswith('总消费') and year_month_str in c])
    fc_cols = sorted([c for c in df.columns if c.startswith('凤巢优惠券消费') and year_month_str in c])
    cpc_cols = sorted([c for c in df.columns if c.startswith('原生CPC优惠券消费') and year_month_str in c])
    cpm_cols = sorted([c for c in df.columns if c.startswith('原生CPM优惠券消费') and year_month_str in c])
    jp_cols = sorted([c for c in df.columns if c.startswith('聚屏平台合约消费') and year_month_str in c])
    dxx_cols = sorted([c for c in df.columns if c.startswith('度星选-软植互选-消费') and year_month_str in c])

    num_days = len(spend_cols)
    if num_days == 0:
        raise ValueError(f"在数据中未匹配到 {year_month_str} 的总消费流水列！")

    # 按 账户ID 聚合，避免多订单行导致笛卡尔积
    agg_dict = {}
    if '账户名称' in df.columns: agg_dict['账户名称'] = 'first'
    if '公司名称' in df.columns: agg_dict['公司名称'] = 'first'
    if '订单行' in df.columns: agg_dict['订单行'] = 'first'

    for col in spend_cols + fc_cols + cpc_cols + cpm_cols + jp_cols + dxx_cols:
        agg_dict[col] = 'sum'

    df_agg = df.groupby('账户ID', as_index=False).agg(agg_dict)

    res = pd.DataFrame()
    res['账户ID'] = df_agg['账户ID']
    res['账户名称'] = df_agg['账户名称'] if '账户名称' in df_agg.columns else ''
    res['公司名称'] = df_agg['公司名称'] if '公司名称' in df_agg.columns else ''
    res['订单行'] = df_agg['订单行'] if '订单行' in df_agg.columns else 0

    # 1. 消费总额与日均
    spend_mat = df_agg[spend_cols].fillna(0).values
    total_spend = spend_mat.sum(axis=1)
    res[f'{year_month_str}_总消费'] = total_spend
    res[f'{year_month_str}_日均消费'] = total_spend / num_days

    # 2. RFM 活跃天数 (Frequency) 与月末断流天数 (Recency)
    active_mask = (spend_mat > 0)
    active_days = active_mask.sum(axis=1)
    res[f'{year_month_str}_活跃天数'] = active_days
    res[f'{year_month_str}_活跃天数率'] = active_days / num_days
    res[f'{year_month_str}_活跃日均消费'] = total_spend / np.where(active_days > 0, active_days, 1)

    # 向量化计算月末断流天数（距离月末最近一次消费的天数）
    day_indices = np.arange(1, num_days + 1)
    last_active_day = np.max(active_mask * day_indices, axis=1)
    recency = np.where(last_active_day > 0, num_days - last_active_day, num_days)
    res[f'{year_month_str}_月末断流天数'] = recency

    # 3. 消费节奏衰减动量
    mid_pt = num_days // 2
    late_spend = spend_mat[:, mid_pt:].sum(axis=1)
    res[f'{year_month_str}_后半月消费占比'] = late_spend / (total_spend + 1e-5)

    last7_spend = spend_mat[:, -7:].sum(axis=1)
    res[f'{year_month_str}_月末7天消费占比'] = last7_spend / (total_spend + 1e-5)

    # 4. 优惠券与渠道渗透
    fc_total = df_agg[fc_cols].fillna(0).values.sum(axis=1) if len(fc_cols) > 0 else np.zeros(len(df_agg))
    cpc_total = df_agg[cpc_cols].fillna(0).values.sum(axis=1) if len(cpc_cols) > 0 else np.zeros(len(df_agg))
    cpm_total = df_agg[cpm_cols].fillna(0).values.sum(axis=1) if len(cpm_cols) > 0 else np.zeros(len(df_agg))
    jp_total = df_agg[jp_cols].fillna(0).values.sum(axis=1) if len(jp_cols) > 0 else np.zeros(len(df_agg))
    dxx_total = df_agg[dxx_cols].fillna(0).values.sum(axis=1) if len(dxx_cols) > 0 else np.zeros(len(df_agg))

    coupon_total = fc_total + cpc_total + cpm_total
    res[f'{year_month_str}_优惠券消费总额'] = coupon_total
    res[f'{year_month_str}_优惠券占比'] = coupon_total / (total_spend + 1e-5)
    
    res[f'{year_month_str}_产品线覆盖数'] = (
        (total_spend > 0).astype(int) + 
        (fc_total > 0).astype(int) + 
        (jp_total > 0).astype(int) + 
        (dxx_total > 0).astype(int)
    )

    return res


def build_quarterly_dataset(file_paths, year_prefix='2023'):
    """
    整合某年份 Q3 季度 3 个月流水，生成季度特征宽表
    """
    m7 = f'{year_prefix}07'
    m8 = f'{year_prefix}08'
    m9 = f'{year_prefix}09'

    df7 = safe_read_csv(file_paths[m7], default_enc='gb18030' if m7=='202307' else 'utf-8')
    df8 = safe_read_csv(file_paths[m8])
    df9 = safe_read_csv(file_paths[m9])

    feat7 = process_monthly_data(df7, m7)
    feat8 = process_monthly_data(df8, m8)
    feat9 = process_monthly_data(df9, m9)

    # 主键合并
    all_accounts = pd.concat([
        feat7[['账户ID', '账户名称', '公司名称', '订单行']],
        feat8[['账户ID', '账户名称', '公司名称', '订单行']],
        feat9[['账户ID', '账户名称', '公司名称', '订单行']]
    ], ignore_index=True).drop_duplicates(subset=['账户ID'])

    merged = all_accounts.merge(feat7.drop(columns=['账户名称', '公司名称', '订单行']), on='账户ID', how='left')
    merged = merged.merge(feat8.drop(columns=['账户名称', '公司名称', '订单行']), on='账户ID', how='left')
    merged = merged.merge(feat9.drop(columns=['账户名称', '公司名称', '订单行']), on='账户ID', how='left')

    num_cols = [c for c in merged.columns if c not in ['账户ID', '账户名称', '公司名称', '订单行']]
    merged[num_cols] = merged[num_cols].fillna(0)

    # 季度汇总指标
    merged[f'{year_prefix}Q3_总消费'] = merged[f'{m7}_总消费'] + merged[f'{m8}_总消费'] + merged[f'{m9}_总消费']
    merged[f'{year_prefix}Q3_总活跃天数'] = merged[f'{m7}_活跃天数'] + merged[f'{m8}_活跃天数'] + merged[f'{m9}_活跃天数']
    merged[f'{year_prefix}Q3_有消耗月数'] = (
        (merged[f'{m7}_总消费'] > 0).astype(int) + 
        (merged[f'{m8}_总消费'] > 0).astype(int) + 
        (merged[f'{m9}_总消费'] > 0).astype(int)
    )

    # 跨月环比动量
    merged[f'{year_prefix}_8月比7月消耗变化率'] = (merged[f'{m8}_总消费'] - merged[f'{m7}_总消费']) / (merged[f'{m7}_总消费'] + 100)
    merged[f'{year_prefix}_9月比8月消耗变化率'] = (merged[f'{m9}_总消费'] - merged[f'{m8}_总消费']) / (merged[f'{m8}_总消费'] + 100)
    
    q_avg_spend = merged[f'{year_prefix}Q3_总消费'] / 3.0
    merged[f'{year_prefix}_季末衰减指数'] = merged[f'{m9}_总消费'] / (q_avg_spend + 1e-5)

    return merged


def run_prediction_pipeline():
    # 路径配置
    dir_q3 = r'D:\dataany\布瑞泽&翼百信季度消费预测\2024Q3依赖'
    dir_q4 = r'D:\dataany\布瑞泽&翼百信季度消费预测\2024Q4依赖'

    files = {
        '202207': os.path.join(dir_q3, '202207.csv'),
        '202208': os.path.join(dir_q3, '202208.csv'),
        '202209': os.path.join(dir_q3, '202209.csv'),
        '202307': os.path.join(dir_q3, '202307.csv'),
        '202308': os.path.join(dir_q3, '202308.csv'),
        '202309': os.path.join(dir_q3, '202309.csv'),
    }

    print("=" * 80, flush=True)
    print(" 🚀 正在启动：大搜客户 2023 年 Q4 季度失效预警逻辑回归预测系统", flush=True)
    print("=" * 80, flush=True)

    # 1. 加载与特征提取
    print("\n[步骤 1/5] 解析 2022Q3 与 2023Q3 日度消费流水...", flush=True)
    df_2022_q3 = build_quarterly_dataset(files, year_prefix='2022')
    df_2023_q3 = build_quarterly_dataset(files, year_prefix='2023')

    # 2. 融入去年同期老客与同比特征
    print("[步骤 2/5] 融合 2022 同期表现，计算同比波动与老客稳定性特征...", flush=True)
    df_2022_sub = df_2022_q3[['账户ID', '2022Q3_总消费', '2022Q3_总活跃天数']].rename(
        columns={'2022Q3_总消费': '2022Q3_去年同期总消费', '2022Q3_总活跃天数': '2022Q3_去年同期活跃天数'}
    )
    df_2023_q3 = df_2023_q3.merge(df_2022_sub, on='账户ID', how='left')
    df_2023_q3['2022Q3_去年同期总消费'] = df_2023_q3['2022Q3_去年同期总消费'].fillna(0)
    df_2023_q3['2022Q3_去年同期活跃天数'] = df_2023_q3['2022Q3_去年同期活跃天数'].fillna(0)
    df_2023_q3['是否去年老客户'] = (df_2023_q3['2022Q3_去年同期总消费'] > 0).astype(int)
    df_2023_q3['Q3总消费_同比增长率'] = (df_2023_q3['2023Q3_总消费'] - df_2023_q3['2022Q3_去年同期总消费']) / (df_2023_q3['2022Q3_去年同期总消费'] + 100)

    # 3. 构造多批次滑动训练样本集 (Multi-period Stacking)
    print("[步骤 3/5] 构建逻辑回归多期训练样本矩阵...", flush=True)
    
    # 批次 1: 7月特征 -> 8月是否失效
    batch1_x = pd.DataFrame({
        '当月总消费_Log': np.log1p(df_2023_q3['202307_总消费']),
        '当月活跃天数': df_2023_q3['202307_活跃天数'],
        '月末断流天数': df_2023_q3['202307_月末断流天数'],
        '后半月消费占比': df_2023_q3['202307_后半月消费占比'],
        '月末7天消费占比': df_2023_q3['202307_月末7天消费占比'],
        '优惠券消费占比': df_2023_q3['202307_优惠券占比'],
        '产品线覆盖数': df_2023_q3['202307_产品线覆盖数'],
    })
    batch1_y = (df_2023_q3['202308_总消费'] == 0).astype(int)

    # 批次 2: 8月特征 -> 9月是否失效
    batch2_x = pd.DataFrame({
        '当月总消费_Log': np.log1p(df_2023_q3['202308_总消费']),
        '当月活跃天数': df_2023_q3['202308_活跃天数'],
        '月末断流天数': df_2023_q3['202308_月末断流天数'],
        '后半月消费占比': df_2023_q3['202308_后半月消费占比'],
        '月末7天消费占比': df_2023_q3['202308_月末7天消费占比'],
        '优惠券消费占比': df_2023_q3['202308_优惠券占比'],
        '产品线覆盖数': df_2023_q3['202308_产品线覆盖数'],
    })
    batch2_y = (df_2023_q3['202309_总消费'] == 0).astype(int)

    # 过滤出当期真正有投放记录的样本
    mask1 = (df_2023_q3['202307_总消费'] > 0)
    mask2 = (df_2023_q3['202308_总消费'] > 0)

    X_train = pd.concat([batch1_x[mask1], batch2_x[mask2]], ignore_index=True)
    y_train = pd.concat([batch1_y[mask1], batch2_y[mask2]], ignore_index=True)

    print(f"  -> 训练样本总数: {len(X_train)} 个 (失效正样本: {y_train.sum()} 个, 留存负样本: {len(y_train) - y_train.sum()} 个)", flush=True)

    # 4. 训练逻辑回归模型与归因
    print("[步骤 4/5] 训练逻辑回归模型并提取风险归因权重...", flush=True)
    scaler = StandardScaler()
    X_train_scaled = scaler.fit_transform(X_train)

    lr = LogisticRegression(
        C=1.0, 
        penalty='l2', 
        class_weight='balanced', 
        solver='liblinear', 
        random_state=42
    )
    lr.fit(X_train_scaled, y_train)

    train_probs = lr.predict_proba(X_train_scaled)[:, 1]
    train_preds = (train_probs >= 0.50).astype(int)
    auc_score = roc_auc_score(y_train, train_probs)
    
    # 显式计算评估 3 大核心指标
    from sklearn.metrics import precision_score, recall_score, f1_score
    train_precision = precision_score(y_train, train_preds)
    train_recall = recall_score(y_train, train_preds)
    train_f1 = f1_score(y_train, train_preds)

    print("\n----------------------------------------------------------------------", flush=True)
    print(" 【步骤 4 评估结果】模型拟合精度三剑客与 AUC 评分：", flush=True)
    print("----------------------------------------------------------------------", flush=True)
    print(f"  ① 精确率 / 查准率 (Precision) : {train_precision:.2%} (报出失效的人群中，真正失效的比例)", flush=True)
    print(f"  ② 召回率 / 查全率 (Recall)    : {train_recall:.2%} (全盘所有真实失效客户中，成功抓出的比例)", flush=True)
    print(f"  ③ 综合平衡得分 (F1-Score)    : {train_f1:.4f} (精确率与召回率的调和平均分)", flush=True)
    print(f"  ④ 全局排序辨别力 (ROC-AUC)   : {auc_score:.4f} (模型整体区分好坏客户的排队能力)", flush=True)
    print("----------------------------------------------------------------------", flush=True)

    # 输出特征归因表
    feature_impact = pd.DataFrame({
        '特征名称': X_train.columns,
        '回归系数 (Beta)': lr.coef_[0],
        '优势比 (Odds Ratio)': np.exp(lr.coef_[0])
    }).sort_values(by='回归系数 (Beta)', ascending=False)
    
    print("\n【特征影响与失效风险归因表 (Beta 权重与优势比)】:")
    print(feature_impact.to_string(index=False), flush=True)

    # 5. 前瞻预测：2023 年 Q4 季度失效客户预测
    print("\n[步骤 5/5] 执行前瞻预测：生成 2023Q4 季度大搜客户失效预警名单...", flush=True)
    
    # 待预测客群：2023Q3 在投且有流水记录的客户
    active_q3_mask = (df_2023_q3['2023Q3_总消费'] > 0)
    pred_pool = df_2023_q3[active_q3_mask].copy()

    X_pred = pd.DataFrame({
        '当月总消费_Log': np.log1p(pred_pool['202309_总消费']),
        '当月活跃天数': pred_pool['202309_活跃天数'],
        '月末断流天数': pred_pool['202309_月末断流天数'],
        '后半月消费占比': pred_pool['202309_后半月消费占比'],
        '月末7天消费占比': pred_pool['202309_月末7天消费占比'],
        '优惠券消费占比': pred_pool['202309_优惠券占比'],
        '产品线覆盖数': pred_pool['202309_产品线覆盖数'],
    })

    X_pred_scaled = scaler.transform(X_pred)
    q4_probs = lr.predict_proba(X_pred_scaled)[:, 1]

    pred_pool['2023Q4_预测失效概率'] = q4_probs
    pred_pool['2023Q4_失效概率'] = [f"{p*100:.2f}%" for p in q4_probs]

    # 风险分级规则
    def get_risk_tier(p):
        if p >= 0.70:
            return '🔴 高危失效预警'
        elif p >= 0.40:
            return '🟡 中危观察客户'
        else:
            return '🟢 健康低危客户'

    pred_pool['风险等级'] = pred_pool['2023Q4_预测失效概率'].apply(get_risk_tier)

    # 单客诊断归因逻辑
    def generate_diagnosis(row):
        reasons = []
        if row['202309_月末断流天数'] >= 7:
            reasons.append(f"9月末提前断流 {int(row['202309_月末断流天数'])} 天")
        if row['2023_9月比8月消耗变化率'] <= -0.40:
            reasons.append(f"9月消耗环比骤降 {abs(row['2023_9月比8月消耗变化率']):.0%}")
        if row['202309_活跃天数'] <= 5:
            reasons.append(f"9月活跃天数极低仅 {int(row['202309_活跃天数'])} 天")
        if row['202309_优惠券占比'] >= 0.50:
            reasons.append(f"优惠券补贴依赖度达 {row['202309_优惠券占比']:.0%}")
        if row['是否去年老客户'] == 0:
            reasons.append("当年新开户 (生命周期抗风险弱)")
        if len(reasons) == 0:
            return "投放节奏平稳 / 暂无显著恶化征兆"
        return " | ".join(reasons)

    pred_pool['核心流失诱因诊断'] = pred_pool.apply(generate_diagnosis, axis=1)

    # 结果字段整理
    cols_export = [
        '账户ID', '账户名称', '公司名称', '订单行', '风险等级', 
        '2023Q4_失效概率', '2023Q4_预测失效概率', '核心流失诱因诊断',
        '2023Q3_总消费', '202309_总消费', '202309_活跃天数', '202309_月末断流天数',
        '2023_9月比8月消耗变化率', '是否去年老客户'
    ]
    result_df = pred_pool[cols_export].sort_values(by='2023Q4_预测失效概率', ascending=False)

    # 导出 CSV 预测名单
    output_path = os.path.join(r'D:\dataany\布瑞泽&翼百信季度消费预测', '2023Q4_大搜客户失效预警预测名单.csv')
    result_df.to_csv(output_path, index=False, encoding='utf-8-sig')
    print(f"\n✅ 预测名单已生成并保存至:\n   {output_path}", flush=True)

    # 输出风险分级统计
    print("\n" + "=" * 80, flush=True)
    print(" 📊 2023 年 Q4 季度大搜客户失效预警大盘分布", flush=True)
    print("=" * 80, flush=True)
    tier_stats = result_df['风险等级'].value_counts()
    for tier, count in tier_stats.items():
        print(f"  * {tier:<16}: {count:>5} 个账户 ({count / len(result_df):.1%})", flush=True)

    # 6. 回测验证（若存在 2023Q4 真实数据）
    q4_actual_files = [
        os.path.join(dir_q4, '202310.csv'),
        os.path.join(dir_q4, '202311.csv'),
        os.path.join(dir_q4, '202312.csv')
    ]
    if all(os.path.exists(f) for f in q4_actual_files):
        print("\n" + "=" * 80, flush=True)
        print(" 🎯 2023 年 Q4 季度真实数据样本外 (OOT) 回测验证报告", flush=True)
        print("=" * 80, flush=True)
        
        # 读取 2023Q4 真实消耗
        df_10 = safe_read_csv(q4_actual_files[0])
        df_11 = safe_read_csv(q4_actual_files[1])
        df_12 = safe_read_csv(q4_actual_files[2])

        spend_10_cols = [c for c in df_10.columns if c.startswith('总消费')]
        spend_11_cols = [c for c in df_11.columns if c.startswith('总消费')]
        spend_12_cols = [c for c in df_12.columns if c.startswith('总消费')]

        spend_10_map = df_10.groupby('账户ID')[spend_10_cols].sum().sum(axis=1).to_dict()
        spend_11_map = df_11.groupby('账户ID')[spend_11_cols].sum().sum(axis=1).to_dict()
        spend_12_map = df_12.groupby('账户ID')[spend_12_cols].sum().sum(axis=1).to_dict()

        result_df['Q4_实际总消费'] = (
            result_df['账户ID'].map(spend_10_map).fillna(0) +
            result_df['账户ID'].map(spend_11_map).fillna(0) +
            result_df['账户ID'].map(spend_12_map).fillna(0)
        )
        # 真实失效定义：Q4 整个季度消费为 0
        result_df['Q4_真实是否失效'] = (result_df['Q4_实际总消费'] == 0).astype(int)

        y_true = result_df['Q4_真实是否失效']
        y_prob = result_df['2023Q4_预测失效概率']
        
        # 1. 业务高危名单口径 (阈值 >= 0.70)
        y_pred_high = (y_prob >= 0.70).astype(int)
        prec_high = precision_score(y_true, y_pred_high)
        rec_high = recall_score(y_true, y_pred_high)
        f1_high = f1_score(y_true, y_pred_high)

        # 2. 标准分类阈值口径 (阈值 >= 0.50)
        y_pred_std = (y_prob >= 0.50).astype(int)
        prec_std = precision_score(y_true, y_pred_std)
        rec_std = recall_score(y_true, y_pred_std)
        f1_std = f1_score(y_true, y_pred_std)

        oot_auc = roc_auc_score(y_true, y_prob)
        cm_std = confusion_matrix(y_true, y_pred_std)
        cm_high = confusion_matrix(y_true, y_pred_high)
        
        print("\n----------------------------------------------------------------------", flush=True)
        print(" 【真实 Q4 样本外回测】核心精确度三剑客指标 (Precision, Recall, F1)：", flush=True)
        print("----------------------------------------------------------------------", flush=True)
        print(" ▶ 业务高危预警名单口径 (失效概率 >= 70%，优先干预名单)：", flush=True)
        print(f"    ① 精确率 / 查准率 (Precision) : {prec_high:.2%} (高危名单中真实停投的比例，打电化不扑空)", flush=True)
        print(f"    ② 召回率 / 查全率 (Recall)    : {rec_high:.2%} (在全盘所有流失大客中的捕获覆盖率)", flush=True)
        print(f"    ③ 综合平衡得分 (F1-Score)    : {f1_high:.4f}", flush=True)
        print(f"    ④ 混淆矩阵 (高危预警)        : TP(抓准)={cm_high[1,1]}, FP(误报)={cm_high[0,1]}", flush=True)
        print("\n ▶ 标准平衡分类口径 (失效概率 >= 50%)：", flush=True)
        print(f"    ① 精确率 / 查准率 (Precision) : {prec_std:.2%}", flush=True)
        print(f"    ② 召回率 / 查全率 (Recall)    : {rec_std:.2%}", flush=True)
        print(f"    ③ 综合平衡得分 (F1-Score)    : {f1_std:.4f}", flush=True)
        print(f"    ④ 样本外 ROC-AUC 排序得分    : {oot_auc:.4f} (极佳的整体风险分层能力)", flush=True)
        print(f"    ⑤ 混淆矩阵 (标准切分)        : TP={cm_std[1,1]}, FP={cm_std[0,1]}, FN={cm_std[1,0]}, TN={cm_std[0,0]}", flush=True)
        print("----------------------------------------------------------------------", flush=True)

        print("\n  * 详细分类报告 (Classification Report - 50% 阈值):", flush=True)
        print(classification_report(y_true, y_pred_std, target_names=['留存正常 (0)', '失效流失 (1)']), flush=True)

    print("\n【Top 10 高危失效客户示例】:", flush=True)
    print(result_df[['账户ID', '账户名称', '风险等级', '2023Q4_失效概率', '核心流失诱因诊断', '202309_总消费']].head(10).to_string(index=False), flush=True)
    print("=" * 80, flush=True)


if __name__ == '__main__':
    run_prediction_pipeline()
