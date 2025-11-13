import pandas as pd

# 入出力ファイルパス
input_path = "検収記録簿 原本.xlsx"
output_path = "検収記録簿_加工済_空欄補完済.xlsx"

# 7・8行目をヘッダーとして読み込み
df = pd.read_excel(input_path, header=[6, 7])
df.columns = ['_'.join([str(i) for i in col if str(i) != 'nan']).strip() for col in df.columns]

# ■ 空欄を上のセルの値で補完（納品日・使用日・朝昼夕・仕入先）
fill_cols = [
    'Unnamed: 0_level_0_納品日',
    'Unnamed: 1_level_0_使用日',
    'Unnamed: 2_level_0_朝昼夕',
    'Unnamed: 3_level_0_仕入先'
]
df[fill_cols] = df[fill_cols].ffill()

# ■ 並び替え用マッピング
order_map = {'朝食': 1, '昼食': 2, '夕食': 3, '3時': 4}
df['朝昼夕_order'] = df['Unnamed: 2_level_0_朝昼夕'].map(order_map).fillna(5)

# ■ 並び替え（使用日 → 朝昼夕 → 食品名）
df_sorted = df.sort_values(by=[
    'Unnamed: 1_level_0_使用日',
    '朝昼夕_order',
    'Unnamed: 5_level_0_食品名'
])

# ■ 抽出する列の指定
cols = [
    'Unnamed: 0_level_0_納品日',
    'Unnamed: 1_level_0_使用日',
    'Unnamed: 2_level_0_朝昼夕',
    'Unnamed: 3_level_0_仕入先',
    'Unnamed: 5_level_0_食品名',
    'Unnamed: 6_level_0_換算値',
    'Unnamed: 7_level_0_総合計',
    'Unnamed: 8_level_0_単位',
    '介護老人福祉施設いわと_入所者',
    '介護老人福祉施設いわと_職員',
    'ケアハウスユー…_入所者'
]

# ■ 出力用データフレーム作成
df_out = df_sorted[cols].reset_index(drop=True)

# ■ Excelファイルとして出力
df_out.to_excel(output_path, index=False)

print("✅ 加工完了:", output_path)
