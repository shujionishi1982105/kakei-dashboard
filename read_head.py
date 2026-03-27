import pandas as pd
df = pd.read_csv('c:/Users/大西修司/OneDrive/Desktop/かけい整形外科_経営分析ダッシュボード/月報_品目別使用量一覧表_202203.csv', encoding='cp932', nrows=5)
with open('c:/Users/大西修司/OneDrive/Desktop/かけい整形外科_経営分析ダッシュボード/columns.txt', 'w', encoding='utf-8') as f:
    f.write(", ".join(df.columns) + "\n")
    f.write(", ".join(map(str, df.iloc[0].values)) + "\n")
