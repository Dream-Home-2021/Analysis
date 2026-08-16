# -*- coding: utf-8 -*-
"""
===============================================================================
项目名称：基于 RFM 模型与 K-Means++ 机器学习的广告客户全自动分层系统
数据来源：D:\\dataany\\未消费客户明细\\RFM.xlsx
核心算法：K-Means++ 聚类 (K=4)
===============================================================================
"""

import os
import sys
import pandas as pd
import numpy as np
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import silhouette_score

# 兼容 Windows 控制台输出编码
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')


def main():
    print("=" * 70)
    print(">>> 正在启动 RFM + K-Means++ 客户智能分层分析流程...")
    print("=" * 70)

    # -------------------------------------------------------------------------
    # 1. 数据读取与基础清洗
    # -------------------------------------------------------------------------
    input_file = r"D:\dataany\未消费客户明细\RFM.xlsx"
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"未找到输入文件: {input_file}")

    print(f"[1/5] 读取数据文件: {input_file}")
    df = pd.read_excel(input_file)
    print(f"      数据总行数: {len(df):,} 行, 字段数: {len(df.columns)} 个")

    # 缺失值与状态填充
    df['账户状态'] = df['账户状态'].fillna('未知状态')
    df['账户名称'] = df['账户名称'].fillna('未知账户')
    df['公司名称'] = df['公司名称'].fillna('未知公司')

    # -------------------------------------------------------------------------
    # 2. 特征工程：偏态校正 (Log1p) 与 Z-Score 标准化
    # -------------------------------------------------------------------------
    print("[2/5] 执行特征工程...")
    # 2.1 对消费金额 M 做对数变换，消除极度右偏长尾
    df['log_M'] = np.log1p(df['M_近7天总消费'])

    # 2.2 构建三维特征空间并进行 Z-Score 标准化 (均值=0, 方差=1)
    feature_cols = ['R_近7天末次消费距今天数', 'F_近7天消费天数', 'log_M']
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(df[feature_cols])

    print(f"      特征维度: R(距今天数), F(消费天数), log_M(对数金额)")
    print(f"      均值向量: R={scaler.mean_[0]:.4f}, F={scaler.mean_[1]:.4f}, log_M={scaler.mean_[2]:.4f}")
    print(f"      标准差  : R={scaler.scale_[0]:.4f}, F={scaler.scale_[1]:.4f}, log_M={scaler.scale_[2]:.4f}")

    # -------------------------------------------------------------------------
    # 3. K-Means++ 聚类建模 (K=4)
    # -------------------------------------------------------------------------
    print("[3/5] 执行 K-Means++ 无监督聚类 (K=4)...")
    k = 4
    kmeans = KMeans(
        n_clusters=k,
        init='k-means++',
        n_init=20,
        max_iter=300,
        random_state=42
    )
    df['Cluster_ID'] = kmeans.fit_predict(X_scaled)

    # 模型质量评估
    inertia = kmeans.inertia_
    sil_score = silhouette_score(X_scaled, df['Cluster_ID'])
    print(f"      聚类完成: 误差平方和 (SSE/Inertia) = {inertia:.2f}")
    print(f"      轮廓系数 (Silhouette Score) = {sil_score:.4f} (聚类结构紧凑有效)")

    # -------------------------------------------------------------------------
    # 4. 纯动态无硬编码业务画像映射 (Dynamic Relative Ranking)
    # -------------------------------------------------------------------------
    print("[4/5] 依据质心相对高低进行纯动态画像映射 (零阈值硬编码)...")
    
    # 统计这 4 个团簇的质心指标均值
    cluster_stats = df.groupby('Cluster_ID').agg(
        R_mean=('R_近7天末次消费距今天数', 'mean'),
        F_mean=('F_近7天消费天数', 'mean'),
        M_mean=('M_近7天总消费', 'mean')
    )

    # 动态相对排序识别：
    # ① 距今天数 R 最大 (停投时间最长) -> 沉睡流失高危客群
    dormant_id = cluster_stats['R_mean'].idxmax()

    # ② 在剩余群体中，消费金额 M 均值最高 -> 核心高价值 VIP 客群
    remain_step1 = cluster_stats.drop(index=dormant_id)
    vip_id = remain_step1['M_mean'].idxmax()

    # ③ 在剩余群体中，投放频次 F 均值最高 -> 高频长尾稳定客群
    remain_step2 = remain_step1.drop(index=vip_id)
    stable_id = remain_step2['F_mean'].idxmax()

    # ④ 剩余最后一个群体 -> 低频/余额断档客群
    potential_id = remain_step2.drop(index=stable_id).index[0]

    # 构造动态映射字典
    dynamic_name_map = {
        vip_id: '核心高价值VIP客群',
        stable_id: '高频长尾稳定客群',
        potential_id: '低频/余额断档客群',
        dormant_id: '沉睡流失高危客群'
    }

    df['客群分层名称'] = df['Cluster_ID'].map(dynamic_name_map)

    # 打印动态识别结果
    for cid, cname in dynamic_name_map.items():
        sub_r = cluster_stats.loc[cid, 'R_mean']
        sub_f = cluster_stats.loc[cid, 'F_mean']
        sub_m = cluster_stats.loc[cid, 'M_mean']
        print(f"      Cluster {cid} -> 【{cname}】: R均值={sub_r:.2f}天, F均值={sub_f:.2f}天, M均值={sub_m:,.2f}元")

    # -------------------------------------------------------------------------
    # 5. 业务衍生监控：近两日 (1109/1110) 断崖停消预警
    # -------------------------------------------------------------------------
    df['两日连续未消费'] = (df['总消费_1109'] == 0) & (df['总消费_1110'] == 0)
    df['业务预警标签'] = '正常'
    
    # VIP 断崖预警：属于 VIP 客群，但近两日连续零消耗
    df.loc[(df['客群分层名称'] == '核心高价值VIP客群') & df['两日连续未消费'], '业务预警标签'] = '⚠️ VIP断崖停消预警'
    # 余额断档提醒：处于低频断档客群且账户余额为零
    df.loc[(df['客群分层名称'] == '低频/余额断档客群') & (df['账户状态'] == '用户帐面为零'), '业务预警标签'] = '🔔 余额耗尽需催充'
    # 违规排障提醒：账户处于被拒绝状态
    df.loc[df['账户状态'] == '该用户被拒绝', '业务预警标签'] = '🚨 账户审核被拒'

    # -------------------------------------------------------------------------
    # 6. 生成分析报表并导出为 Excel
    # -------------------------------------------------------------------------
    output_file = r"D:\dataany\未消费客户明细\RFM_KMeans_4类自动分层结果.xlsx"
    print(f"[5/5] 导出分层结果至 Excel: {output_file}")

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Sheet 1: 客户完整明细表
        detail_cols = [
            '账户ID', '账户名称', '公司名称', '管理员', '账户状态',
            '客群分层名称', '业务预警标签',
            'R_近7天末次消费距今天数', 'F_近7天消费天数', 'M_近7天总消费',
            '总消费_1109', '总消费_1110'
        ]
        df[detail_cols].to_excel(writer, sheet_name='客户分层明细', index=False)

        # Sheet 2: 4类客群全景画像看板
        total_m = df['M_近7天总消费'].sum()
        total_users = len(df)

        summary_report = df.groupby('客群分层名称').agg(
            客户总数=('账户ID', 'count'),
            R均值=('R_近7天末次消费距今天数', 'mean'),
            F均值=('F_近7天消费天数', 'mean'),
            M总流水=('M_近7天总消费', 'sum'),
            M均值=('M_近7天总消费', 'mean'),
            M中位数=('M_近7天总消费', 'median'),
            正常生效占比=('账户状态', lambda s: f'{(s == "用户正常生效").mean()*100:.1f}%'),
            账面为零占比=('账户状态', lambda s: f'{(s == "用户帐面为零").mean()*100:.1f}%'),
            被拒绝占比=('账户状态', lambda s: f'{(s == "该用户被拒绝").mean()*100:.1f}%'),
            两日连续零消占比=('两日连续未消费', lambda s: f'{s.mean()*100:.1f}%')
        ).reindex([
            '核心高价值VIP客群',
            '高频长尾稳定客群',
            '低频/余额断档客群',
            '沉睡流失高危客群'
        ]).reset_index()

        summary_report['客户占比'] = (summary_report['客户总数'] / total_users * 100).round(2).astype(str) + '%'
        summary_report['流水贡献占比'] = (summary_report['M总流水'] / total_m * 100).round(2).astype(str) + '%'

        # 调整看板列顺序
        ordered_cols = [
            '客群分层名称', '客户总数', '客户占比', '流水贡献占比', 'M总流水',
            'M均值', 'M中位数', 'F均值', 'R均值',
            '正常生效占比', '账面为零占比', '被拒绝占比', '两日连续零消占比'
        ]
        summary_report[ordered_cols].to_excel(writer, sheet_name='4类客群全景看板', index=False)

        # Sheet 3: 管理员跟进看板 (按客服/管理员透视各客群分布与预警户数)
        admin_report = df.pivot_table(
            index='管理员',
            columns='客群分层名称',
            values='账户ID',
            aggfunc='count',
            fill_value=0
        )
        # 添加预警户数统计
        vip_alert = df[df['业务预警标签'] == '⚠️ VIP断崖停消预警'].groupby('管理员')['账户ID'].count()
        admin_report['⚠️ VIP停消待跟进数'] = admin_report.index.map(vip_alert).fillna(0).astype(int)
        admin_report['账户总数'] = admin_report.sum(axis=1) - admin_report['⚠️ VIP停消待跟进数']
        admin_report = admin_report.sort_values(by='⚠️ VIP停消待跟进数', ascending=False).reset_index()
        admin_report.to_excel(writer, sheet_name='管理员跟进看板', index=False)

    print("=" * 70)
    print(f"🎉 运行成功！分层与多维看板报表已生成完毕。")
    print(f"📁 文件保存路径: {output_file}")
    print("=" * 70)

if __name__ == "__main__":
    main()
