import json
import pandas as pd

EXCEL_PATH = "혼인상태별 인구구성_20251117102829.xlsx"
OUTPUT_JSON = "marriage_data.json"

df = pd.read_excel(EXCEL_PATH)

# 27번째 행에 2005 / 2010 / 2015 / 2020 가로로 들어있음
row_year_header = df.iloc[27]

years_positions = {}
for col_name, val in row_year_header.items():
    if isinstance(val, (int, float, str)) and str(val).strip() in ("2005", "2010", "2015", "2020"):
        years_positions[int(val)] = col_name

years = sorted(years_positions.keys())
year_to_base_idx = {year: df.columns.get_loc(col_name) for year, col_name in years_positions.items()}

status_names = ["미혼", "유배우", "사별", "이혼"]
gender_map_kr_to_short = {"남자": "M", "여자": "F"}
genders_short = ["M", "F"]

# 연령대 라벨 (어느 연도나 동일 블록 구조)
first_year = years[0]
base_idx = year_to_base_idx[first_year]
cols = df.columns[base_idx : base_idx + 15]
age_labels = list(df.loc[28, cols][1:])  # 첫 번째는 '합계'라서 제외

# 전체 구조: percentages[year][M/F][status][age] = 값(%)
percentages = {}

for year in years:
    base_idx = year_to_base_idx[year]
    cols = df.columns[base_idx : base_idx + 15]
    percentages[str(year)] = {g: {} for g in genders_short}

    current_gender_kr = None
    # 실제 데이터는 29~43행에 있음 (전체/남자/여자 × 상태)
    for ridx in range(29, 44):
        gender_val = df.iloc[ridx, 0]
        if isinstance(gender_val, str) and gender_val in ["전체", "남자", "여자"]:
            current_gender_kr = gender_val

        status_val = df.iloc[ridx, 1]
        # 우리는 남자, 여자만 사용
        if current_gender_kr in gender_map_kr_to_short and status_val in status_names:
            g_short = gender_map_kr_to_short[current_gender_kr]
            row = df.loc[ridx, cols]
            vals = list(row[1:])  # 합계 열은 건너뛰고 연령대 14개

            percentages[str(year)].setdefault(g_short, {})
            percentages[str(year)][g_short].setdefault(status_val, {})

            for age_label, pct in zip(age_labels, vals):
                pct_f = float(pct) if pd.notna(pct) else 0.0
                percentages[str(year)][g_short][status_val][age_label] = pct_f

# 최종 JSON 구조
out = {
    "ages": age_labels,
    "years": years,
    "genders": ["M", "F"],
    "statuses": status_names,
    "percentages": percentages,
}

with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)

print(f"saved → {OUTPUT_JSON}")
