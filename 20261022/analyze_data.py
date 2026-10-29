import pandas as pd
import numpy as np
from scipy import stats

# Read excel
df = pd.read_excel('c:/生物統計實習/data/調查局20260331.xlsx')
print("Dataset columns:", df.columns.tolist())
print("Total rows:", len(df))
print(df.head())

# Let's inspect unique values of all columns
for col in df.columns:
    print(f"\n--- {col} ---")
    print(df[col].value_counts(dropna=False))

# Disease defined as: Healthy vs Diseased (minor + severe)
df['is_diseased'] = df['disease status'].apply(lambda x: 0 if str(x).strip().lower() == 'healthy' else 1)
df['is_dead'] = df['survival outcome'].apply(lambda x: 1 if str(x).strip().lower() == 'dead' else 0)
df['age_group'] = df['age'].apply(lambda x: '<=40' if x <= 40 else '41+')

print("\nCross tab: is_diseased vs is_dead:")
print(pd.crosstab(df['is_diseased'], df['is_dead'], margins=True))

print("\nCross tab: disease status vs survival outcome:")
print(pd.crosstab(df['disease status'], df['survival outcome'], margins=True))

print("\nCross tab: drinking water vs disease status:")
print(pd.crosstab(df['drinking water'], df['disease status'], margins=True))

print("\nCross tab: drinking water vs is_diseased:")
print(pd.crosstab(df['drinking water'], df['is_diseased'], margins=True))

print("\nCross tab: drinking water vs survival outcome:")
print(pd.crosstab(df['drinking water'], df['survival outcome'], margins=True))

print("\nCross tab: water_group vs survival outcome:")
print(pd.crosstab(df['water_group'], df['survival outcome'], margins=True))

print("\nCross tab: water_group vs is_diseased:")
print(pd.crosstab(df['water_group'], df['is_diseased'], margins=True))
