import sys
import pandas as pd
import numpy as np
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import silhouette_score
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os

print(">>> Starting script...")

plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial']
plt.rcParams['axes.unicode_minus'] = False

file_path = r'D:\dataany\RFM与kmeans++\RFM.xlsx'
if not os.path.exists(file_path):
    print(f"Error: {file_path} not found")
    sys.exit(1)

print(">>> Loading data...")
df = pd.read_excel(file_path)
df['log_M'] = np.log1p(df['M_近7天总消费'])
X_scaled = StandardScaler().fit_transform(df[['R_近7天末次消费距今天数', 'F_近7天消费天数', 'log_M']])

sse = []
sil_scores = []
K_range = range(2, 10)

print(">>> Starting KMeans loop...")
for k in K_range:
    km = KMeans(n_clusters=k, init='k-means++', n_init=10, random_state=42)
    labels = km.fit_predict(X_scaled)
    sse.append(km.inertia_)
    sil_scores.append(silhouette_score(X_scaled, labels, sample_size=1000, random_state=42))
    print(f"Finished K={k}")

print(">>> Plotting...")
fig, ax1 = plt.subplots(figsize=(10, 5))

color = 'tab:red'
ax1.set_xlabel('K (聚类簇数)', fontsize=12)
ax1.set_ylabel('SSE (误差平方和)', color=color, fontsize=12)
ax1.plot(K_range, sse, marker='o', color=color, linewidth=2)
ax1.tick_params(axis='y', labelcolor=color)

ax2 = ax1.twinx()
color = 'tab:blue'
ax2.set_ylabel('轮廓系数 (Silhouette Score)', color=color, fontsize=12)
ax2.plot(K_range, sil_scores, marker='s', color=color, linewidth=2)
ax2.tick_params(axis='y', labelcolor=color)

fig.suptitle('K-Means 聚类评估：肘部法则 (SSE) vs 轮廓系数', fontsize=14)
fig.tight_layout()

ax1.axvline(x=4, color='gray', linestyle='--', alpha=0.7)
ax1.text(4.2, max(sse)*0.85, '拐点 K=4', color='black', fontsize=12)

out_path = r'D:\dataany\RFM与kmeans++\KMeans_评估指标图.png'
plt.savefig(out_path, dpi=300)
print(">>> Saved to", out_path)
