import pandas as pd
import numpy as np
from scipy.stats import chi2_contingency, fisher_exact

df = pd.read_csv('data.csv')

# Clean columns
df['is_diseased'] = df['disease status'].apply(lambda x: 0 if str(x).strip().lower() == 'healthy' else 1)
df['is_dead'] = df['survival outcome'].apply(lambda x: 1 if str(x).strip().lower() == 'dead' else 0)
df['age_group'] = df['age'].apply(lambda x: '<=40' if x <= 40 else '41+')

total = len(df)

lines = []
def log(s=""):
    lines.append(s)

log("==================================================")
log("橘郡異常事件統計分析詳細報告")
log(f"總樣本數 N = {total}")
log("==================================================")

# 第一部分
log("\n【第一部分：人口學與分佈描述性統計】")
num_diseased = df['is_diseased'].sum()
log(f"1. 患病比例: {num_diseased}/{total} ({num_diseased/total*100:.2f}%) [無病/健康者: {total-num_diseased}人, {(total-num_diseased)/total*100:.2f}%]")

log("\n2. 飲用水來源種類比例:")
water_counts = df['drinking water'].value_counts()
for k in ['Deep well', 'shallow well', 'Top water']:
    cnt = water_counts.get(k, 0)
    log(f"   - {k}: {cnt} 人 ({cnt/total*100:.2f}%)")

log("\n3. 患病嚴重程度分布情形:")
sev_counts = df['disease status'].value_counts()
for k in ['Healthy', 'minor', 'severe']:
    cnt = sev_counts.get(k, 0)
    log(f"   - {k}: {cnt} 人 ({cnt/total*100:.2f}%)")

num_dead = df['is_dead'].sum()
log(f"\n4. 疾病死亡率:")
log(f"   - 全體居民粗死亡率 (Overall Mortality Rate): {num_dead}/{total} ({num_dead/total*100:.2f}%)")
cfr = df[df['is_diseased'] == 1]['is_dead'].sum() / num_diseased
log(f"   - 患病者致死率 (Case Fatality Rate, CFR): {df[df['is_diseased'] == 1]['is_dead'].sum()}/{num_diseased} ({cfr*100:.2f}%)")
log(f"   - 健康者死亡率: {df[df['is_diseased'] == 0]['is_dead'].sum()}/{total-num_diseased} (0.00%)")

log("\n5. 性別比例:")
gender_counts = df['Gender'].value_counts()
for k in ['F', 'M']:
    cnt = gender_counts.get(k, 0)
    log(f"   - {k}: {cnt} 人 ({cnt/total*100:.2f}%)")
log(f"   - 男女比例 (M:F): {gender_counts['M']}:{gender_counts['F']} = {gender_counts['M']/gender_counts['F']:.4f}:1")
log(f"   - 女男比例 (F:M): {gender_counts['F']}:{gender_counts['M']} = {gender_counts['F']/gender_counts['M']:.4f}:1")

log("\n6. 年齡分佈:")
age_counts = df['age_group'].value_counts()
for k in ['<=40', '41+']:
    cnt = age_counts.get(k, 0)
    log(f"   - {k}: {cnt} 人 ({cnt/total*100:.2f}%)")
log(f"   - 年齡平均數 ± 標準差: {df['age'].mean():.2f} ± {df['age'].std():.2f} 歲 (範圍: {df['age'].min()} - {df['age'].max()} 歲)")

log("\n" + "="*50)
log("【第二部分：居民患病及死亡之原因調查（假設檢定與關聯指標）】")

def calc_2x2_stats(table_2x2, exp_name, unexp_name, out_name, nonout_name):
    a, b = int(table_2x2[0, 0]), int(table_2x2[0, 1])
    c, d = int(table_2x2[1, 0]), int(table_2x2[1, 1])
    
    r1 = a / (a + b) if (a + b) > 0 else 0
    r0 = c / (c + d) if (c + d) > 0 else 0
    rr = r1 / r0 if r0 > 0 else np.nan
    
    if a > 0 and c > 0:
        se_ln_rr = np.sqrt((1 - r1)/a + (1 - r0)/c)
        rr_ci = (np.exp(np.log(rr) - 1.96 * se_ln_rr), np.exp(np.log(rr) + 1.96 * se_ln_rr))
    else:
        rr_ci = (np.nan, np.nan)
        
    if a == 0 or b == 0 or c == 0 or d == 0:
        a_c, b_c, c_c, d_c = a + 0.5, b + 0.5, c + 0.5, d + 0.5
        or_val = (a_c * d_c) / (b_c * c_c)
        se_ln_or = np.sqrt(1/a_c + 1/b_c + 1/c_c + 1/d_c)
    else:
        or_val = (a * d) / (b * c)
        se_ln_or = np.sqrt(1/a + 1/b + 1/c + 1/d)
        
    or_ci = (np.exp(np.log(or_val) - 1.96 * se_ln_or), np.exp(np.log(or_val) + 1.96 * se_ln_or))
    
    chi2, p, dof, _ = chi2_contingency([[a, b], [c, d]], correction=False)
    chi2_corr, p_corr, _, _ = chi2_contingency([[a, b], [c, d]], correction=True)
    odds_ratio, p_fisher = fisher_exact([[a, b], [c, d]])
    
    log(f"\n--- 分析項目: {exp_name} vs. {unexp_name} 對 【{out_name}】 之關聯 ---")
    log(f"2x2 列聯表:")
    log(f"  變項維度               {out_name:>12} {nonout_name:>12} {'合計 (Total)':>12} {'發生率 (Risk)':>12}")
    log(f"  {exp_name:<18} {a:12d} {b:12d} {a+b:12d} {r1*100:11.2f}%")
    log(f"  {unexp_name:<18} {c:12d} {d:12d} {c+d:12d} {r0*100:11.2f}%")
    log(f"  合計                   {a+c:12d} {b+d:12d} {a+b+c+d:12d}")
    log(f"  * Pearson 卡方值 (Chi-Square): {chi2:.4f}, 自由度 df = 1, P值 = {p:.4e}")
    log(f"  * 連續性校正卡方值 (Continuity Correction): {chi2_corr:.4f}, P值 = {p_corr:.4e}")
    log(f"  * 費雪精確檢定 (Fisher's Exact Test) 雙尾 P值: {p_fisher:.4e}")
    log(f"  * 風險比 (Risk Ratio, RR): {rr:.4f} (95% 信賴區間: {rr_ci[0]:.4f} - {rr_ci[1]:.4f})")
    log(f"  * 勝算比 (Odds Ratio, OR): {or_val:.4f} (95% 信賴區間: {or_ci[0]:.4f} - {or_ci[1]:.4f})")
    return {
        'a': a, 'b': b, 'c': c, 'd': d,
        'r1': r1, 'r0': r0, 'rr': rr, 'rr_ci': rr_ci,
        'or': or_val, 'or_ci': or_ci,
        'chi2': chi2, 'p': p, 'chi2_corr': chi2_corr, 'p_corr': p_corr,
        'p_fisher': p_fisher
    }

# 1. 深井水 vs 非深井水 (water_group 2 vs 1) 與 死亡
t1 = np.array([
    [len(df[(df['water_group']==2) & (df['is_dead']==1)]), len(df[(df['water_group']==2) & (df['is_dead']==0)])],
    [len(df[(df['water_group']==1) & (df['is_dead']==1)]), len(df[(df['water_group']==1) & (df['is_dead']==0)])]
])
calc_2x2_stats(t1, '深井水 (Deep well)', '非深井水 (淺井+自來)', '死亡 (Dead)', '存活 (Alive)')

# 2. 深井水 vs 非深井水 與 患病
t2 = np.array([
    [len(df[(df['water_group']==2) & (df['is_diseased']==1)]), len(df[(df['water_group']==2) & (df['is_diseased']==0)])],
    [len(df[(df['water_group']==1) & (df['is_diseased']==1)]), len(df[(df['water_group']==1) & (df['is_diseased']==0)])]
])
calc_2x2_stats(t2, '深井水 (Deep well)', '非深井水 (淺井+自來)', '患病 (Diseased)', '健康 (Healthy)')

# 3. 飲用水來源（三類）與 患病嚴重程度 (3x3)
log("\n" + "-"*50)
log("--- 飲用水來源 (3種) 與 患病嚴重程度 (3種) 交叉表 ---")
ct_water_dis_3x3 = pd.crosstab(df['drinking water'], df['disease status'])
log(ct_water_dis_3x3.to_string())
chi2_3x3, p_3x3, dof_3x3, _ = chi2_contingency(ct_water_dis_3x3)
log(f"Pearson 卡方值: {chi2_3x3:.4f}, 自由度 df = {dof_3x3}, P值 = {p_3x3:.4e}")

# 4. 飲用水來源（三類）與 生存狀況 (3x2)
log("\n--- 飲用水來源 (3種) 與 生存狀況 (2種) 交叉表 ---")
ct_water_surv_3x2 = pd.crosstab(df['drinking water'], df['survival outcome'])
log(ct_water_surv_3x2.to_string())
chi2_3x2, p_3x2, dof_3x2, _ = chi2_contingency(ct_water_surv_3x2)
log(f"Pearson 卡方值: {chi2_3x2:.4f}, 自由度 df = {dof_3x2}, P值 = {p_3x2:.4e}")

# 5. 性別 與 死亡
t_g_death = np.array([
    [len(df[(df['Gender']=='M') & (df['is_dead']==1)]), len(df[(df['Gender']=='M') & (df['is_dead']==0)])],
    [len(df[(df['Gender']=='F') & (df['is_dead']==1)]), len(df[(df['Gender']=='F') & (df['is_dead']==0)])]
])
calc_2x2_stats(t_g_death, '男性 (Male)', '女性 (Female)', '死亡 (Dead)', '存活 (Alive)')

# 6. 性別 與 患病
t_g_dis = np.array([
    [len(df[(df['Gender']=='M') & (df['is_diseased']==1)]), len(df[(df['Gender']=='M') & (df['is_diseased']==0)])],
    [len(df[(df['Gender']=='F') & (df['is_diseased']==1)]), len(df[(df['Gender']=='F') & (df['is_diseased']==0)])]
])
calc_2x2_stats(t_g_dis, '男性 (Male)', '女性 (Female)', '患病 (Diseased)', '健康 (Healthy)')

# 7. 年齡組 (41+ vs <=40) 與 死亡
t_a_death = np.array([
    [len(df[(df['age_group']=='41+') & (df['is_dead']==1)]), len(df[(df['age_group']=='<=40') & (df['is_dead']==0)])],
    [len(df[(df['age_group']=='<=40') & (df['is_dead']==1)]), len(df[(df['age_group']=='41+') & (df['is_dead']==0)])]
])
# Note: correct contingency table for age group
ct_age_death_tab = np.array([
    [len(df[(df['age_group']=='41+') & (df['is_dead']==1)]), len(df[(df['age_group']=='41+') & (df['is_dead']==0)])],
    [len(df[(df['age_group']=='<=40') & (df['is_dead']==1)]), len(df[(df['age_group']=='<=40') & (df['is_dead']==0)])]
])
calc_2x2_stats(ct_age_death_tab, '中高齡 (41+)', '年輕組 (<=40)', '死亡 (Dead)', '存活 (Alive)')

# 8. 年齡組 (41+ vs <=40) 與 患病
ct_age_dis_tab = np.array([
    [len(df[(df['age_group']=='41+') & (df['is_diseased']==1)]), len(df[(df['age_group']=='41+') & (df['is_diseased']==0)])],
    [len(df[(df['age_group']=='<=40') & (df['is_diseased']==1)]), len(df[(df['age_group']=='<=40') & (df['is_diseased']==0)])]
])
calc_2x2_stats(ct_age_dis_tab, '中高齡 (41+)', '年輕組 (<=40)', '患病 (Diseased)', '健康 (Healthy)')

# 9. 兩兩飲用水來源細部比較:
for pair in [('Deep well', 'shallow well'), ('Deep well', 'Top water'), ('shallow well', 'Top water')]:
    sub = df[df['drinking water'].isin(pair)]
    ct = np.array([
        [len(sub[(sub['drinking water']==pair[0]) & (sub['is_dead']==1)]), len(sub[(sub['drinking water']==pair[0]) & (sub['is_dead']==0)])],
        [len(sub[(sub['drinking water']==pair[1]) & (sub['is_dead']==1)]), len(sub[(sub['drinking water']==pair[1]) & (sub['is_dead']==0)])]
    ])
    calc_2x2_stats(ct, f"{pair[0]}", f"{pair[1]}", '死亡 (Dead)', '存活 (Alive)')

with open('stats_results.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))

print("stats_results.txt created successfully!")
