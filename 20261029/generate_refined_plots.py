import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import shutil
import os
from lifelines import KaplanMeierFitter, NelsonAalenFitter
from lifelines.statistics import logrank_test
from lifelines.plotting import add_at_risk_counts

# Configure font
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'Segoe UI', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# Read data
data_path = 'Survival test example.xlsx'
df = pd.read_excel(data_path)

g1 = df[df['Group'] == 1]
g2 = df[df['Group'] == 2]

# Fit KM
kmf_placebo = KaplanMeierFitter()
kmf_placebo.fit(g1['Follow-up Months'], g1['Status 1=event 0=censor'], label='對照組 (Placebo, n=6)')

kmf_newdrug = KaplanMeierFitter()
kmf_newdrug.fit(g2['Follow-up Months'], g2['Status 1=event 0=censor'], label='新藥組 (New Drug, n=6)')

# Fit Nelson-Aalen for cumulative hazard
naf_placebo = NelsonAalenFitter()
naf_placebo.fit(g1['Follow-up Months'], g1['Status 1=event 0=censor'], label='對照組 (Placebo)')

naf_newdrug = NelsonAalenFitter()
naf_newdrug.fit(g2['Follow-up Months'], g2['Status 1=event 0=censor'], label='新藥組 (New Drug)')

# Log-rank test
lr = logrank_test(g1['Follow-up Months'], g2['Follow-up Months'],
                  g1['Status 1=event 0=censor'], g2['Status 1=event 0=censor'])

c_placebo = '#0072B2'  # Deep medical blue
c_newdrug = '#D55E00'  # Vermillion / medical red-orange

# -------------------------------------------------------------
# Figure 1: Clean Publication Kaplan-Meier Curve with At-Risk Table
# -------------------------------------------------------------
fig, ax = plt.subplots(figsize=(8.5, 6), dpi=300)

kmf_placebo.plot_survival_function(
    ax=ax,
    ci_show=False,
    color=c_placebo,
    linewidth=2.8,
    show_censors=True,
    censor_styles={'marker': '+', 'ms': 10, 'mew': 2.0}
)

kmf_newdrug.plot_survival_function(
    ax=ax,
    ci_show=False,
    color=c_newdrug,
    linewidth=2.8,
    show_censors=True,
    censor_styles={'marker': '+', 'ms': 10, 'mew': 2.0}
)

# Reference 50% line for median survival
ax.axhline(0.5, color='#888888', linestyle=':', linewidth=1.2, alpha=0.8, label='50% 存活基準線 (中位存活時間)')

ax.set_ylim(-0.02, 1.05)
ax.set_xlim(0, 8.2)
ax.set_xlabel('追蹤時間 (Follow-up Months)', fontsize=12, fontweight='bold', labelpad=8)
ax.set_ylabel('累積存活率 (Cumulative Survival Probability)', fontsize=12, fontweight='bold', labelpad=8)
ax.set_title('Kaplan-Meier 存活分析曲線 (Kaplan-Meier Survival Curves)', fontsize=14, fontweight='bold', pad=12)

ax.grid(True, linestyle='--', alpha=0.35, color='#b0bec5')
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_linewidth(1.2)
ax.spines['bottom'].set_linewidth(1.2)
ax.tick_params(axis='both', which='major', labelsize=10.5)

# Annotation summary box
stat_text = (
    f"【統計檢定與存活期摘要】\n"
    f"• 對照組中位存活期: 6.50 個月 (95% CI: 3.56 - 9.44)\n"
    f"• 新藥組中位存活期: 5.00 個月 (95% CI: 4.04 - 5.96)\n"
    f"• Log-Rank 檢定: χ² = {lr.test_statistic:.3f}, p = {lr.p_value:.3f} (無顯著差異)\n"
    f"• Breslow 檢定: χ² = 1.330, p = 0.249\n"
    f"• Tarone-Ware 檢定: χ² = 1.777, p = 0.183\n"
    f"註: '+' 表示設限值 (Censored data)"
)
ax.text(
    0.04, 0.08, stat_text,
    transform=ax.transAxes,
    fontsize=9.5,
    verticalalignment='bottom',
    bbox=dict(boxstyle='round,pad=0.7', facecolor='#ffffff', edgecolor='#90a4ae', alpha=0.95, linewidth=1.2)
)

ax.legend(
    loc='upper right',
    frameon=True,
    facecolor='white',
    edgecolor='#cfd8dc',
    fontsize=10.5,
    title='試驗組別 (Group)',
    title_fontsize='11'
)

add_at_risk_counts(kmf_placebo, kmf_newdrug, ax=ax, rows_to_show=['At risk'])
plt.tight_layout()

km_path = 'kaplan_meier_curve.png'
fig.savefig(km_path, dpi=300, bbox_inches='tight')
plt.close(fig)

# -------------------------------------------------------------
# Figure 2: Dual Plot (Survival Curve + Cumulative Hazard)
# -------------------------------------------------------------
fig2, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5.5), dpi=300)

# Left: Survival
kmf_placebo.plot_survival_function(ax=ax1, ci_show=True, ci_alpha=0.12, color=c_placebo, linewidth=2.4, show_censors=True)
kmf_newdrug.plot_survival_function(ax=ax1, ci_show=True, ci_alpha=0.12, color=c_newdrug, linewidth=2.4, show_censors=True)
ax1.axhline(0.5, color='#777777', linestyle=':', linewidth=1.1, alpha=0.7)
ax1.set_ylim(-0.02, 1.05)
ax1.set_xlim(0, 8.2)
ax1.set_xlabel('追蹤時間 (Follow-up Months)', fontsize=11, fontweight='bold')
ax1.set_ylabel('累積存活率 (Cumulative Survival)', fontsize=11, fontweight='bold')
ax1.set_title('(A) Kaplan-Meier 存活函數曲線', fontsize=12, fontweight='bold')
ax1.grid(True, linestyle='--', alpha=0.35)
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.legend(loc='lower left', fontsize=9.5)

# Right: Cumulative Hazard
naf_placebo.plot_cumulative_hazard(ax=ax2, ci_show=False, color=c_placebo, linewidth=2.4)
naf_newdrug.plot_cumulative_hazard(ax=ax2, ci_show=False, color=c_newdrug, linewidth=2.4)
ax2.set_xlim(0, 8.2)
ax2.set_ylim(0, 2.2)
ax2.set_xlabel('追蹤時間 (Follow-up Months)', fontsize=11, fontweight='bold')
ax2.set_ylabel('累積危險 (Cumulative Hazard)', fontsize=11, fontweight='bold')
ax2.set_title('(B) 累積危險函數曲線 (Hazard Function)', fontsize=12, fontweight='bold')
ax2.grid(True, linestyle='--', alpha=0.35)
ax2.spines['top'].set_visible(False)
ax2.spines['right'].set_visible(False)
ax2.legend(loc='upper left', fontsize=9.5)

plt.suptitle('新藥試驗存活分析：存活函數與累積危險對比 (SPSS & Lifelines 同步結果)', fontsize=14, fontweight='bold', y=1.02)
plt.tight_layout()

dual_path = 'kaplan_meier_dual_plots.png'
fig2.savefig(dual_path, dpi=300, bbox_inches='tight')
plt.close(fig2)

# Copy to brain artifact directory
artifact_dir = r'C:\Users\User\.gemini\antigravity-ide\brain\581e3b5f-7028-420a-bc7b-f0226e37cd43'
shutil.copy(km_path, os.path.join(artifact_dir, km_path))
shutil.copy(dual_path, os.path.join(artifact_dir, dual_path))

print("All plots generated and copied to artifact directory successfully.")
