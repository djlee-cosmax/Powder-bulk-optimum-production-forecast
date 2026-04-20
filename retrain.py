import pandas as pd
import numpy as np
from sklearn.ensemble import GradientBoostingRegressor
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_absolute_error
import json
import os

# ============ 1. 데이터 로드 ============
print("데이터 로드 중...")
df = pd.read_excel("/mnt/c/Users/djlee/OneDrive - COSMAX/바탕 화면/표준 대비 실적 데이터_260420.xlsx", sheet_name="데이터")
print(f"전체 행: {len(df):,}")

# ============ 2. 전처리 ============
print("전처리 중...")

# 필요 컬럼 추출
df = df.rename(columns=lambda x: x.strip())

# 숫자 변환
for col in ['오더수량', '실적수량', '표준소요량', '투입소요량', '사용파손수량', '표준대비투입율', '로스율']:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '').str.replace('%', ''), errors='coerce').fillna(0)

# 실적수량 > 0인 행만
df = df[df['실적수량'] > 0].copy()
print(f"유효 행 (실적수량>0): {len(df):,}")

# 벌크코드 (구성부품)
df['벌크코드'] = df['구성부품'].astype(str).str.strip()
df['성형물코드'] = df['최상위자재'].astype(str).str.strip()
df['작업장'] = df['작업장내역'].astype(str).str.strip()
df['관리유형'] = df['관리유형내역'].astype(str).str.strip()

# 단위당 표준소요량
df['stdPerUnit'] = df['표준소요량'] / df['실적수량']
df['actualPerUnit'] = df['투입소요량'] / df['실적수량']

# 로스율 계산: (실제투입 - 표준소요) / 표준소요 * 100
df['calcLossRate'] = np.where(
    df['표준소요량'] > 0,
    ((df['투입소요량'] - df['표준소요량']) / df['표준소요량']) * 100,
    0
)

# 이상치 제거 (로스율 -50% ~ 200%)
df = df[(df['calcLossRate'] >= -50) & (df['calcLossRate'] <= 200)].copy()
# 표준소요량 3000 이하 제거
df = df[df['표준소요량'] > 3000].copy()
print(f"이상치 제거 후: {len(df):,}")

# ============ 3. 피처 엔지니어링 ============
print("피처 엔지니어링 중...")

# 카테고리 → 숫자 매핑
type_map = {}
for i, t in enumerate(df['관리유형'].unique()):
    type_map[t] = i

machine_map = {}
for i, m in enumerate(df['작업장'].unique()):
    machine_map[m] = i

df['type_encoded'] = df['관리유형'].map(type_map)
df['machine_encoded'] = df['작업장'].map(machine_map)

# 벌크별 통계
bulk_stats = df.groupby('벌크코드').agg(
    avg_loss=('calcLossRate', 'mean'),
    std_loss=('calcLossRate', 'std'),
    count=('calcLossRate', 'count'),
    median_loss=('calcLossRate', 'median'),
    avg_std_per_unit=('stdPerUnit', 'mean')
).reset_index()
bulk_stats['std_loss'] = bulk_stats['std_loss'].fillna(0)

df = df.merge(bulk_stats[['벌크코드', 'avg_loss', 'std_loss', 'median_loss', 'avg_std_per_unit']], on='벌크코드', how='left')

# 피처 선택
features = ['실적수량', 'type_encoded', 'machine_encoded', 'avg_loss', 'std_loss', 'median_loss', 'avg_std_per_unit']
target = 'calcLossRate'

df_model = df[features + [target, '벌크코드']].dropna()
print(f"학습 데이터: {len(df_model):,}")

# ============ 4. 모델 학습 ============
print("모델 학습 중...")

X = df_model[features].values
y = df_model[target].values

X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

model = GradientBoostingRegressor(
    n_estimators=200,
    max_depth=5,
    learning_rate=0.1,
    min_samples_leaf=10,
    random_state=42
)
model.fit(X_train, y_train)

y_pred = model.predict(X_test)
mae = mean_absolute_error(y_test, y_pred)
print(f"테스트 MAE: {mae:.2f}%")

# ============ 5. 벌크별 예측 결과 생성 ============
print("예측 결과 생성 중...")

qty_ranges = [3000, 5000, 10000, 20000, 30000, 50000, 100000]
bulk_lookup = {}

for bulk_code, group in df_model.groupby('벌크코드'):
    avg_loss = group['avg_loss'].iloc[0]
    cnt = len(group)

    # 벌크별 대표 피처 (최빈값)
    row = group.iloc[0]
    type_enc = group['type_encoded'].mode().iloc[0] if len(group['type_encoded'].mode()) > 0 else row['type_encoded']
    machine_enc = group['machine_encoded'].mode().iloc[0] if len(group['machine_encoded'].mode()) > 0 else row['machine_encoded']
    avg_l = row['avg_loss']
    std_l = row['std_loss']
    med_l = row['median_loss']
    avg_spu = row['avg_std_per_unit']

    preds = {}
    for qty in qty_ranges:
        feat = np.array([[qty, type_enc, machine_enc, avg_l, std_l, med_l, avg_spu]])
        pred = model.predict(feat)[0]
        preds[str(qty)] = round(max(pred, 0), 2)

    bulk_lookup[bulk_code] = {
        "avg": round(avg_loss, 2),
        "cnt": cnt,
        "pred": preds
    }

# ============ 6. 저장 ============
output = {
    "bulkLookup": bulk_lookup,
    "modelInfo": {
        "algorithm": "GradientBoosting",
        "dataCount": len(df_model),
        "testMAE": round(mae, 2),
        "qtyRanges": qty_ranges
    }
}

output_dir = "/home/djlee/cosmax/project2"

# JSON
with open(os.path.join(output_dir, "ml_predictions.json"), "w", encoding="utf-8") as f:
    json.dump(output, f, ensure_ascii=False)

# JS (브라우저용)
with open(os.path.join(output_dir, "ml_predictions.js"), "w", encoding="utf-8") as f:
    f.write("var ML_PREDICTIONS = ")
    json.dump(output, f, ensure_ascii=False)
    f.write(";")

# 매핑 저장
mappings = {
    "type_map": type_map,
    "machine_map": machine_map
}
with open(os.path.join(output_dir, "ml_mappings.json"), "w", encoding="utf-8") as f:
    json.dump(mappings, f, ensure_ascii=False, indent=2)

print(f"\n===== 학습 완료 =====")
print(f"벌크 수: {len(bulk_lookup):,}")
print(f"학습 데이터: {len(df_model):,}")
print(f"테스트 MAE: {mae:.2f}%")
print(f"저장 완료: ml_predictions.js, ml_predictions.json, ml_mappings.json")
