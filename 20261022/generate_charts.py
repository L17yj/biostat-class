import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# Set font for Traditional Chinese
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'DFKai-SB', 'MingLiU', 'Arial']
plt.rcParams['axes.unicode_minus'] = False

df = pd.read_csv('data.csv')
df['is_diseased'] = df['disease status'].apply(lambda x: 0 if str(x).strip().lower() == 'healthy' else 1)
df['is_dead'] = df['survival outcome'].apply(lambda x: 1 if str(x).strip().lower() == 'dead' else 0)

# Colors
color_deep = '#C0392B'     # Crimson Red
color_shallow = '#E67E22'  # Amber Orange
color_tap = '#2980B9'      # Ocean Blue
color_alive = '#27AE60'    # Emerald Green
color_dead = '#8E44AD'     # Purple
dark_slate = '#2C3E50'

# ==========================================
# Chart 1: 飲用水來源與生存狀況（死亡 vs 存活）
# ==========================================
fig, ax1 = plt.subplots(figsize=(8, 5.5), dpi=300)

waters = ['自來水 (Top water)', '淺井水 (Shallow well)', '深井水 (Deep well)']
water_keys = ['Top water', 'shallow well', 'Deep well']
alives = [len(df[(df['drinking water']==k) & (df['is_dead']==0)]) for k in water_keys]
deads = [len(df[(df['drinking water']==k) & (df['is_dead']==1)]) for k in water_keys]
totals = [a + d for a, d in zip(alives, deads)]
death_rates = [d / t * 100 for d, t in zip(deads, totals)]

x = np.arange(len(waters))
width = 0.35

rects1 = ax1.bar(x - width/2, alives, width, label='存活人數 (Alive)', color='#2ECC71', edgecolor='#27AE60', alpha=0.9)
rects2 = ax1.bar(x + width/2, deads, width, label='死亡人數 (Dead)', color='#E74C3C', edgecolor='#C0392B', alpha=0.9)

ax1.set_ylabel('人數 (人)', fontsize=12, fontweight='bold', color=dark_slate)
ax1.set_title('圖 1：橘郡居民不同飲用水來源之存活與死亡分佈對比', fontsize=14, fontweight='bold', pad=15, color=dark_slate)
ax1.set_xticks(x)
ax1.set_xticklabels([f'{w}\n(N={t})' for w, t in zip(waters, totals)], fontsize=11, fontweight='bold')
ax1.legend(loc='upper left', frameon=True, facecolor='white', edgecolor='none')
ax1.grid(axis='y', linestyle='--', alpha=0.3)

# Add values on bars
for rect in rects1:
    h = rect.get_height()
    if h > 0:
        ax1.annotate(f'{int(h)}人',
                    xy=(rect.get_x() + rect.get_width() / 2, h),
                    xytext=(0, 3), textcoords="offset points",
                    ha='center', va='bottom', fontsize=10, fontweight='bold')
for rect in rects2:
    h = rect.get_height()
    if h > 0:
        ax1.annotate(f'{int(h)}人',
                    xy=(rect.get_x() + rect.get_width() / 2, h),
                    xytext=(0, 3), textcoords="offset points",
                    ha='center', va='bottom', fontsize=10, fontweight='bold')

# Add death rate line on secondary axis
ax2 = ax1.twinx()
ax2.plot(x, death_rates, color='#8E44AD', marker='o', linewidth=2.5, markersize=8, label='死亡率 (%)')
ax2.set_ylabel('死亡率 (%)', fontsize=12, fontweight='bold', color='#8E44AD')
ax2.set_ylim(-5, 115)
ax2.tick_params(axis='y', labelcolor='#8E44AD')
for i, r in enumerate(death_rates):
    ax2.annotate(f'{r:.1f}%', xy=(x[i], r), xytext=(0, 10), textcoords="offset points",
                 ha='center', va='bottom', fontsize=11, fontweight='bold', color='#8E44AD',
                 bbox=dict(boxstyle='round,pad=0.2', facecolor='white', edgecolor='#8E44AD', alpha=0.8))

ax2.legend(loc='upper center', frameon=True, facecolor='white', edgecolor='none')
plt.tight_layout()
plt.savefig('chart1_water_survival.png', dpi=300)
plt.close()
print("Saved chart1_water_survival.png")

# ==========================================
# Chart 2: 飲用水來源與患病嚴重程度百分比堆疊圖
# ==========================================
fig, ax = plt.subplots(figsize=(8, 5.5), dpi=300)

healthy_pcts = [len(df[(df['drinking water']==k) & (df['disease status']=='Healthy')]) / t * 100 for k, t in zip(water_keys, totals)]
minor_pcts = [len(df[(df['drinking water']==k) & (df['disease status']=='minor')]) / t * 100 for k, t in zip(water_keys, totals)]
severe_pcts = [len(df[(df['drinking water']==k) & (df['disease status']=='severe')]) / t * 100 for k, t in zip(water_keys, totals)]

b1 = ax.bar(waters, healthy_pcts, label='健康/無病 (Healthy)', color='#3498DB', edgecolor='white', width=0.55)
b2 = ax.bar(waters, minor_pcts, bottom=healthy_pcts, label='輕度患病 (Minor)', color='#F39C12', edgecolor='white', width=0.55)
b3 = ax.bar(waters, severe_pcts, bottom=[h+m for h, m in zip(healthy_pcts, minor_pcts)], label='重度患病 (Severe)', color='#C0392B', edgecolor='white', width=0.55)

ax.set_ylabel('百分比 (%)', fontsize=12, fontweight='bold', color=dark_slate)
ax.set_title('圖 2：不同飲用水來源之疾病嚴重程度分佈（劑量反應效應）', fontsize=14, fontweight='bold', pad=15, color=dark_slate)
ax.set_ylim(0, 105)
ax.grid(axis='y', linestyle='--', alpha=0.3)
ax.legend(loc='upper right', frameon=True, facecolor='white', edgecolor='gray')

# Add percentage text
for i in range(len(waters)):
    # healthy
    h = healthy_pcts[i]
    if h > 5:
        ax.text(i, h/2, f'{h:.1f}%', ha='center', va='center', color='white', fontweight='bold', fontsize=10)
    # minor
    m = minor_pcts[i]
    if m > 5:
        ax.text(i, h + m/2, f'{m:.1f}%', ha='center', va='center', color='white', fontweight='bold', fontsize=10)
    # severe
    s = severe_pcts[i]
    if s > 5:
        ax.text(i, h + m + s/2, f'{s:.1f}%', ha='center', va='center', color='white', fontweight='bold', fontsize=10)

plt.tight_layout()
plt.savefig('chart2_water_severity.png', dpi=300)
plt.close()
print("Saved chart2_water_severity.png")

# ==========================================
# Chart 3: 危險因子森林圖 (Forest Plot of Relative Risks)
# ==========================================
fig, ax = plt.subplots(figsize=(8.5, 4.5), dpi=300)

factors = [
    '年齡因素 (41+ 歲 vs. <=40 歲)',
    '性別因素 (男性 vs. 女性)',
    '飲用水源 (深井水 vs. 淺井水)',
    '飲用水源 (深井水 vs. 非深井水組)'
]

rrs = [1.005, 0.999, 8.198, 11.905]
ci_lows = [0.925, 0.917, 5.018, 7.211]
ci_highs = [1.092, 1.088, 13.394, 19.655]
p_values = ['p = 0.903 (無關聯)', 'p = 0.978 (無關聯)', 'p < 0.001 (極顯著)', 'p < 0.001 (極顯著)']

y_pos = np.arange(len(factors))

# Horizontal error bars
ax.errorbar(rrs, y_pos, xerr=[np.array(rrs) - np.array(ci_lows), np.array(ci_highs) - np.array(rrs)],
            fmt='s', color='#1B365D', ecolor='#C0392B', elinewidth=2.5, capsize=6, capthick=2, markersize=8)

# Null effect line at RR = 1
ax.axvline(x=1.0, color='gray', linestyle='--', linewidth=1.5, alpha=0.8, label='無效線 (RR = 1.0)')

ax.set_yticks(y_pos)
ax.set_yticklabels(factors, fontsize=11, fontweight='bold')
ax.set_xlabel('風險比 (Risk Ratio, RR) [對數尺度]', fontsize=12, fontweight='bold', color=dark_slate)
ax.set_xscale('log')
ax.set_xlim(0.5, 30)
ax.set_xticks([0.5, 1.0, 2.0, 5.0, 10.0, 20.0])
ax.get_xaxis().set_major_formatter(plt.ScalarFormatter())
ax.set_title('圖 3：各項可疑危險因子之疾病死亡風險比 (RR) 與 95% 信賴區間', fontsize=14, fontweight='bold', pad=15, color=dark_slate)
ax.grid(axis='x', linestyle=':', alpha=0.6)

# Annotate points with values and p-values
for i in range(len(factors)):
    ax.text(ci_highs[i] * 1.08, y_pos[i], f'RR = {rrs[i]:.2f} (95% CI: {ci_lows[i]:.2f}-{ci_highs[i]:.2f})\n[{p_values[i]}]',
            va='center', fontsize=9.5, fontweight='bold', color='#2C3E50')

plt.tight_layout()
plt.savefig('chart3_forest_rr.png', dpi=300)
plt.close()
print("Saved chart3_forest_rr.png")
