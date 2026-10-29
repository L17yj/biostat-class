import pandas as pd
import numpy as np
from scipy.stats import chi2_contingency, fisher_exact

df = pd.read_excel('c:/生物統計實習/data/調查局20260331.xlsx')

# Clean columns
df['is_diseased'] = df['disease status'].apply(lambda x: 0 if str(x).strip().lower() == 'healthy' else 1)
df['is_dead'] = df['survival outcome'].apply(lambda x: 1 if str(x).strip().lower() == 'dead' else 0)
df['age_group'] = df['age'].apply(lambda x: '<=40' if x <= 40 else '41+')

print("=== 1. 基本人口學與分佈 ===")
print("總人數 N =", len(df))

# 1. 患病比例
num_diseased = df['is_diseased'].sum()
total = len(df)
print(f"患病人數: {num_diseased}, 比例: {num_diseased/total*100:.2f}%")

# 2. 飲用水來源種類比例
print("\n飲用水來源比例:")
water_counts = df['drinking water'].value_counts()
water_pct = df['drinking water'].value_counts(normalize=True) * 100
for k in water_counts.index:
    print(f"  {k}: {water_counts[k]} 人 ({water_pct[k]:.2f}%)")

# 3. 患病嚴重程度分佈
print("\n患病嚴重程度分佈:")
sev_counts = df['disease status'].value_counts()
sev_pct = df['disease status'].value_counts(normalize=True) * 100
for k in ['Healthy', 'minor', 'severe']:
    if k in sev_counts:
        print(f"  {k}: {sev_counts[k]} 人 ({sev_pct[k]:.2f}%)")

# 4. 疾病死亡率
num_dead = df['is_dead'].sum()
print(f"\n全體粗死亡率: {num_dead}/{total} ({num_dead/total*100:.2f}%)")
# 致死率 (Case Fatality Rate, CFR): 患病者中的死亡率
cfr = df[df['is_diseased'] == 1]['is_dead'].sum() / num_diseased
print(f"患病者致死率 (Case Fatality Rate): {df[df['is_diseased'] == 1]['is_dead'].sum()}/{num_diseased} ({cfr*100:.2f}%)")

# 5. 性別比例
print("\n性別比例:")
gender_counts = df['Gender'].value_counts()
gender_pct = df['Gender'].value_counts(normalize=True) * 100
for k in gender_counts.index:
    print(f"  {k}: {gender_counts[k]} 人 ({gender_pct[k]:.2f}%)")
print(f"男女比 (M:F): {gender_counts['M']}:{gender_counts['F']} ({gender_counts['M']/gender_counts['F']:.3f}:1)")

print("\n" + "="*50)
print("=== 2. 交叉分析與假設檢定 ===")

def calc_rr_or(contingency_2x2, exposed_label, unexposed_label, outcome_label, non_outcome_label):
    # contingency table format: [[a, b], [c, d]]
    # row 0: exposed, row 1: unexposed
    # col 0: outcome (e.g. disease/dead), col 1: non-outcome
    a = contingency_2x2[0, 0]
    b = contingency_2x2[0, 1]
    c = contingency_2x2[1, 0]
    d = contingency_2x2[1, 1]
    
    # RR = [a / (a + b)] / [c / (c + d)]
    p1 = a / (a + b) if (a + b) > 0 else 0
    p2 = c / (c + d) if (c + d) > 0 else 0
    rr = p1 / p2 if p2 > 0 else np.nan
    
    # ln(RR) SE = sqrt( (b/a)/(a+b) + (d/c)/(c+d) ) = sqrt( 1/a - 1/(a+b) + 1/c - 1/(c+d) )
    if a > 0 and c > 0:
        se_ln_rr = np.sqrt(1/a - 1/(a+b) + 1/c - 1/(c+d))
        rr_ci_low = np.exp(np.log(rr) - 1.96 * se_ln_rr)
        rr_ci_high = np.exp(np.log(rr) + 1.96 * se_ln_rr)
    else:
        rr_ci_low, rr_ci_high = np.nan, np.nan

    # OR = (a * d) / (b * c)
    # Haldane-Anscombe correction if 0
    if a == 0 or b == 0 or c == 0 or d == 0:
        a_c, b_c, c_c, d_c = a + 0.5, b + 0.5, c + 0.5, d + 0.5
        or_val = (a_c * d_c) / (b_c * c_c)
        se_ln_or = np.sqrt(1/a_c + 1/b_c + 1/c_c + 1/d_c)
    else:
        or_val = (a * d) / (b * c)
        se_ln_or = np.sqrt(1/a + 1/b + 1/c + 1/d)
        
    or_ci_low = np.exp(np.log(or_val) - 1.96 * se_ln_or)
    or_ci_high = np.exp(np.log(or_val) + 1.96 * se_ln_or)

    chi2, p, dof, ex = chi2_contingency([[a, b], [c, d]], correction=False)
    chi2_corr, p_corr, _, _ = chi2_contingency([[a, b], [c, d]], correction=True)
    odds_ratio, p_fisher = fisher_exact([[a, b], [c, d]])

    return {
        'table': [[a, b], [c, d]],
        'chi2': chi2,
        'p_val': p,
        'chi2_corr': chi2_corr,
        'p_val_corr': p_corr,
        'p_fisher': p_fisher,
        'rr': rr,
        'rr_ci': (rr_ci_low, rr_ci_high),
        'or': or_val,
        'or_ci': (or_ci_low, or_ci_high)
    }

# Analysis A: water_group (2: Deep well vs 1: Other) and is_dead
# row 0: Deep well (water_group=2), row 1: Other (water_group=1)
# col 0: Dead (1), col 1: Alive (0)
ct_water_death = np.array([
    [len(df[(df['water_group']==2) & (df['is_dead']==1)]), len(df[(df['water_group']==2) & (df['is_dead']==0)])],
    [len(df[(df['water_group']==1) & (df['is_dead']==1)]), len(df[(df['water_group']==1) & (df['is_dead']==0)])]
])
res_water_death = calc_rr_or(ct_water_death, 'Deep well', 'Other', 'Dead', 'Alive')
print("\n[深井水 vs. 死亡 (Death)]")
print("列聯表 [[Dead, Alive]]:", res_water_death['table'])
print(f"Pearson Chi-Square: {res_water_death['chi2']:.4f}, p: {res_water_death['p_val']:.4e}")
print(f"Continuity Correction: {res_water_death['chi2_corr']:.4f}, p: {res_water_death['p_val_corr']:.4e}")
print(f"Fisher Exact p: {res_water_death['p_fisher']:.4e}")
print(f"Relative Risk (RR): {res_water_death['rr']:.4f} (95% CI: {res_water_death['rr_ci'][0]:.4f} - {res_water_death['rr_ci'][1]:.4f})")
print(f"Odds Ratio (OR): {res_water_death['or']:.4f} (95% CI: {res_water_death['or_ci'][0]:.4f} - {res_water_death['or_ci'][1]:.4f})")

# Analysis B: water_group (2: Deep well vs 1: Other) and is_diseased
ct_water_dis = np.array([
    [len(df[(df['water_group']==2) & (df['is_diseased']==1)]), len(df[(df['water_group']==2) & (df['is_diseased']==0)])],
    [len(df[(df['water_group']==1) & (df['is_diseased']==1)]), len(df[(df['water_group']==1) & (df['is_diseased']==0)])]
])
res_water_dis = calc_rr_or(ct_water_dis, 'Deep well', 'Other', 'Diseased', 'Healthy')
print("\n[深井水 vs. 患病 (Diseased)]")
print("列聯表 [[Diseased, Healthy]]:", res_water_dis['table'])
print(f"Pearson Chi-Square: {res_water_dis['chi2']:.4f}, p: {res_water_dis['p_val']:.4e}")
print(f"Continuity Correction: {res_water_dis['chi2_corr']:.4f}, p: {res_water_dis['p_val_corr']:.4e}")
print(f"Fisher Exact p: {res_water_dis['p_fisher']:.4e}")
print(f"Relative Risk (RR): {res_water_dis['rr']:.4f} (95% CI: {res_water_dis['rr_ci'][0]:.4f} - {res_water_dis['rr_ci'][1]:.4f})")
print(f"Odds Ratio (OR): {res_water_dis['or']:.4f} (95% CI: {res_water_dis['or_ci'][0]:.4f} - {res_water_dis['or_ci'][1]:.4f})")

# Analysis C: Gender (M vs F) and is_dead
ct_gender_death = np.array([
    [len(df[(df['Gender']=='M') & (df['is_dead']==1)]), len(df[(df['Gender']=='M') & (df['is_dead']==0)])],
    [len(df[(df['Gender']=='F') & (df['is_dead']==1)]), len(df[(df['Gender']=='F') & (df['is_dead']==0)])]
])
res_gender_death = calc_rr_or(ct_gender_death, 'Male', 'Female', 'Dead', 'Alive')
print("\n[性別 (男 vs 女) vs. 死亡 (Death)]")
print("列聯表 [[Dead, Alive]]:", res_gender_death['table'])
print(f"Pearson Chi-Square: {res_gender_death['chi2']:.4f}, p: {res_gender_death['p_val']:.4f}")
print(f"Relative Risk (RR): {res_gender_death['rr']:.4f} (95% CI: {res_gender_death['rr_ci'][0]:.4f} - {res_gender_death['rr_ci'][1]:.4f})")
print(f"Odds Ratio (OR): {res_gender_death['or']:.4f} (95% CI: {res_gender_death['or_ci'][0]:.4f} - {res_gender_death['or_ci'][1]:.4f})")

# Analysis D: Gender (M vs F) and is_diseased
ct_gender_dis = np.array([
    [len(df[(df['Gender']=='M') & (df['is_diseased']==1)]), len(df[(df['Gender']=='M') & (df['is_diseased']==0)])],
    [len(df[(df['Gender']=='F') & (df['is_diseased']==1)]), len(df[(df['Gender']=='F') & (df['is_diseased']==0)])]
])
res_gender_dis = calc_rr_or(ct_gender_dis, 'Male', 'Female', 'Diseased', 'Healthy')
print("\n[性別 (男 vs 女) vs. 患病 (Diseased)]")
print("列聯表 [[Diseased, Healthy]]:", res_gender_dis['table'])
print(f"Pearson Chi-Square: {res_gender_dis['chi2']:.4f}, p: {res_gender_dis['p_val']:.4f}")
print(f"Relative Risk (RR): {res_gender_dis['rr']:.4f} (95% CI: {res_gender_dis['rr_ci'][0]:.4f} - {res_gender_dis['rr_ci'][1]:.4f})")
print(f"Odds Ratio (OR): {res_gender_dis['or']:.4f} (95% CI: {res_gender_dis['or_ci'][0]:.4f} - {res_gender_dis['or_ci'][1]:.4f})")

# Analysis E: Age Group (41+ vs <=40) and is_dead
ct_age_death = np.array([
    [len(df[(df['age_group']=='41+') & (df['is_dead']==1)]), len(df[(df['age_group']=='41+') & (df['is_dead']==0)])],
    [len(df[(df['age_group']=='<=40') & (df['is_dead']==1)]), len(df[(df['age_group']=='<=40') & (df['is_dead']==0)])]
])
res_age_death = calc_rr_or(ct_age_death, '41+', '<=40', 'Dead', 'Alive')
print("\n[年齡組 (41+ vs <=40) vs. 死亡 (Death)]")
print("列聯表 [[Dead, Alive]]:", res_age_death['table'])
print(f"Pearson Chi-Square: {res_age_death['chi2']:.4f}, p: {res_age_death['p_val']:.4f}")
print(f"Relative Risk (RR): {res_age_death['rr']:.4f} (95% CI: {res_age_death['rr_ci'][0]:.4f} - {res_age_death['rr_ci'][1]:.4f})")
print(f"Odds Ratio (OR): {res_age_death['or']:.4f} (95% CI: {res_age_death['or_ci'][0]:.4f} - {res_age_death['or_ci'][1]:.4f})")

# Analysis F: Drinking water (3 groups) vs Disease Status (3 groups)
ct_water_status = pd.crosstab(df['drinking water'], df['disease status'])
chi2_ws, p_ws, dof_ws, _ = chi2_contingency(ct_water_status)
print("\n[飲用水 3組 vs 疾病嚴重程度 3組 列聯表]")
print(ct_water_status)
print(f"Pearson Chi-Square: {chi2_ws:.4f}, df: {dof_ws}, p: {p_ws:.4e}")

# Analysis G: Drinking water (3 groups) vs Survival outcome (2 groups)
ct_water_surv = pd.crosstab(df['drinking water'], df['survival outcome'])
chi2_wsv, p_wsv, dof_wsv, _ = chi2_contingency(ct_water_surv)
print("\n[飲用水 3組 vs 生存狀況 2組 列聯表]")
print(ct_water_surv)
print(f"Pearson Chi-Square: {chi2_wsv:.4f}, df: {dof_wsv}, p: {p_wsv:.4e}")

# Detailed break-down of drinking water:
# Compare Deep well vs Shallow well, Deep well vs Top water, Shallow well vs Top water
for pair in [('Deep well', 'shallow well'), ('Deep well', 'Top water'), ('shallow well', 'Top water')]:
    sub = df[df['drinking water'].isin(pair)]
    ct = np.array([
        [len(sub[(sub['drinking water']==pair[0]) & (sub['is_dead']==1)]), len(sub[(sub['drinking water']==pair[0]) & (sub['is_dead']==0)])],
        [len(sub[(sub['drinking water']==pair[1]) & (sub['is_dead']==1)]), len(sub[(sub['drinking water']==pair[1]) & (sub['is_dead']==0)])]
    ])
    res = calc_rr_or(ct, pair[0], pair[1], 'Dead', 'Alive')
    print(f"\n子比較: {pair[0]} vs {pair[1]}:")
    print("Table:", res['table'])
    print(f"  Chi2: {res['chi2']:.4f}, p: {res['p_val']:.4e}")
    print(f"  RR: {res['rr']:.4f} (95% CI: {res['rr_ci'][0]:.4f} - {res['rr_ci'][1]:.4f})")
    print(f"  OR: {res['or']:.4f} (95% CI: {res['or_ci'][0]:.4f} - {res['or_ci'][1]:.4f})")
