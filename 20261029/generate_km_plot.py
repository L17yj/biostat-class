import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from lifelines import KaplanMeierFitter
from lifelines.statistics import logrank_test
from lifelines.plotting import add_at_risk_counts

# Configure Chinese font
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'Segoe UI', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# 1. Read data
data_path = 'Survival test example.xlsx'
df = pd.read_excel(data_path)

# Separate groups
# Group 1: Placebo (對照組/安慰劑)
# Group 2: New Drug (新藥組)
g1 = df[df['Group'] == 1]
g2 = df[df['Group'] == 2]

# 2. Fit Kaplan-Meier models
kmf_placebo = KaplanMeierFitter()
kmf_placebo.fit(
    durations=g1['Follow-up Months'],
    event_observed=g1['Status 1=event 0=censor'],
    label='對照組 (Placebo, n=6)'
)

kmf_newdrug = KaplanMeierFitter()
kmf_newdrug.fit(
    durations=g2['Follow-up Months'],
    event_observed=g2['Status 1=event 0=censor'],
    label='新藥組 (New Drug, n=6)'
)

# 3. Log-Rank Test
lr_result = logrank_test(
    g1['Follow-up Months'], g2['Follow-up Months'],
    g1['Status 1=event 0=censor'], g2['Status 1=event 0=censor']
)

# 4. Plotting (NEJM / Lancet style publication figure)
fig, ax = plt.subplots(figsize=(9, 6.5), dpi=300)

# Colors
c_placebo = '#1f77b4'  # Professional Steel Blue
c_newdrug = '#d62728'  # Coral Red / Crimson

# Plot survival curves
kmf_placebo.plot_survival_function(
    ax=ax,
    ci_show=True,
    ci_alpha=0.15,
    color=c_placebo,
    linewidth=2.5,
    show_censors=True,
    censor_styles={'marker': '+', 'ms': 10, 'mew': 2}
)

kmf_newdrug.plot_survival_function(
    ax=ax,
    ci_show=True,
    ci_alpha=0.15,
    color=c_newdrug,
    linewidth=2.5,
    show_censors=True,
    censor_styles={'marker': '+', 'ms': 10, 'mew': 2}
)

# Set axes and labels
ax.set_ylim(-0.02, 1.05)
ax.set_xlim(0, 8.5)
ax.set_xlabel('追蹤時間 (Follow-up Months)', fontsize=13, fontweight='bold', labelpad=10)
ax.set_ylabel('累積存活率 (Cumulative Survival Probability)', fontsize=13, fontweight='bold', labelpad=10)
ax.set_title('Kaplan-Meier 存活曲線 (新藥組 vs. 對照組)', fontsize=15, fontweight='bold', pad=15)

# Grid and spines styling
ax.grid(True, linestyle='--', alpha=0.4, color='#cccccc')
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_linewidth(1.2)
ax.spines['bottom'].set_linewidth(1.2)
ax.tick_params(axis='both', which='major', labelsize=11)

# Annotations box for stats
text_box = (
    f"對照組中位存活期: 6.5 個月 (95% CI: 3.56-9.44)\n"
    f"新藥組中位存活期: 5.0 個月 (95% CI: 4.04-5.96)\n"
    f"Log-Rank (Mantel-Cox) 檢定: χ² = {lr_result.test_statistic:.3f}, p = {lr_result.p_value:.3f}\n"
    f"Breslow 檢定: χ² = 1.330, p = 0.249"
)
ax.text(
    0.04, 0.12, text_box,
    transform=ax.transAxes,
    fontsize=10.5,
    verticalalignment='bottom',
    bbox=dict(boxstyle='round,pad=0.6', facecolor='#f8f9fa', edgecolor='#bdc3c7', alpha=0.95, linewidth=1)
)

# Legend styling
ax.legend(
    loc='upper right',
    frameon=True,
    facecolor='white',
    edgecolor='#e0e0e0',
    fontsize=11,
    title='組別 (Group)',
    title_fontsize='11'
)

# Add Number at Risk table
add_at_risk_counts(kmf_placebo, kmf_newdrug, ax=ax, rows_to_show=['At risk'])

plt.tight_layout()

# Save figures
output_png = 'kaplan_meier_curve.png'
output_pdf = 'kaplan_meier_curve.pdf'
fig.savefig(output_png, dpi=300, bbox_inches='tight')
fig.savefig(output_pdf, bbox_inches='tight')
print(f"Successfully generated {output_png} and {output_pdf}")
