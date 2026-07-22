# -*- coding: utf-8 -*-
"""
학교안전사고 데이터 전처리 및 EDA(빈도/패턴 분석)
- 사고데이터: 발생 빈도/패턴 분석용
- 보상데이터: 보상금 총액 산출 및 중대성 3단계 타겟 생성

사용법:
1. 아래 ACCIDENT_PATH, COMPENSATION_PATH를 본인 파일 경로로 수정
2. python preprocessing_eda.py 실행
3. output/ 폴더에 그래프 png들과 요약 통계 csv가 저장됨
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

# ------------------------------------------------------------------
# 0. 설정
# ------------------------------------------------------------------
ACCIDENT_PATH = "사고데이터.xlsx"       # 본인 파일 경로로 수정
COMPENSATION_PATH = "보상데이터.xlsx"   # 본인 파일 경로로 수정
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 한글 폰트 (Windows: Malgun Gothic / Mac: AppleGothic / Linux: NanumGothic)
plt.rcParams["font.family"] = "Malgun Gothic"  # 본인 OS에 맞게 수정
plt.rcParams["axes.unicode_minus"] = False
sns.set_style("whitegrid")

학교급_순서 = ["초등학교", "중학교", "고등학교"]
요일_순서 = ["월", "화", "수", "목", "금", "토", "일"]


# ------------------------------------------------------------------
# 1. 사고데이터 전처리
# ------------------------------------------------------------------
def load_and_clean_accident(path):
    # 실제 데이터가 3번째 시트부터 5개(연도별 등)로 나뉘어 있는 경우 처리
    # 1~2번째 시트는 표지/데이터설명 등 메타정보이므로 제외
    xls = pd.ExcelFile(path)
    print("전체 시트 목록:", xls.sheet_names)

    target_sheets = xls.sheet_names[2:]  # 3번째 시트부터 끝까지
    print("데이터로 읽을 시트:", target_sheets)

    df_list = []
    for sheet in target_sheets:
        sheet_df = pd.read_excel(path, sheet_name=sheet)
        sheet_df["원본시트"] = sheet  # 어느 시트(연도 등)에서 왔는지 추적
        df_list.append(sheet_df)
    df = pd.concat(df_list, ignore_index=True)

    print("=== 사고데이터 원본 정보 ===")
    print(df.shape)
    print(df.isnull().sum())

    # 결측치: 범주형은 '미상'으로, 시각 등 수치는 중앙값 대체 (컬럼 상황 보고 조정)
    cat_cols = ["지역", "학교급", "사고요일", "사고장소", "사고형태"]
    for c in cat_cols:
        if c in df.columns:
            df[c] = df[c].fillna("미상")

    # 사고연월 -> 연/월/계절 파생
    if "사고연월" in df.columns:
        df["사고연월"] = pd.to_datetime(df["사고연월"], errors="coerce")
        df["연도"] = df["사고연월"].dt.year
        df["월"] = df["사고연월"].dt.month

        def to_season(m):
            if m in [3, 4, 5]:
                return "봄"
            elif m in [6, 7, 8]:
                return "여름"
            elif m in [9, 10, 11]:
                return "가을"
            else:
                return "겨울"
        df["계절"] = df["월"].apply(to_season)

    # 사고발생시각: "쉬는시간", "체육" 등 이미 시간대/활동 구간을 나타내는
    # 텍스트 카테고리이므로 구간화(pd.cut) 없이 그대로 시간대로 사용
    if "사고발생시각" in df.columns:
        df["시간대"] = df["사고발생시각"].fillna("미상")

    # 학교급 순서 지정 (그래프용 categorical order)
    if "학교급" in df.columns:
        df["학교급"] = pd.Categorical(df["학교급"], categories=학교급_순서, ordered=True)
    if "사고요일" in df.columns:
        df["사고요일"] = pd.Categorical(df["사고요일"], categories=요일_순서, ordered=True)

    return df


# ------------------------------------------------------------------
# 2. 보상데이터 전처리 (보상금 총액 + 3단계 타겟)
# ------------------------------------------------------------------
COMP_COLS = ["요양급여", "장해급여", "간병급여", "유족급여", "장례비", "위로금", "보전비용"]

def load_and_clean_compensation(path):
    # 사고데이터와 동일하게 1~2번째 시트는 표지/설명, 3번째 시트부터 실제 데이터
    xls = pd.ExcelFile(path)
    print("전체 시트 목록:", xls.sheet_names)

    target_sheets = xls.sheet_names[2:]
    print("데이터로 읽을 시트:", target_sheets)

    df_list = []
    for sheet in target_sheets:
        sheet_df = pd.read_excel(path, sheet_name=sheet)
        sheet_df["원본시트"] = sheet
        df_list.append(sheet_df)
    df = pd.concat(df_list, ignore_index=True)

    print("=== 보상데이터 원본 정보 ===")
    print(df.shape)
    print(df.isnull().sum())

    # 보상금 항목 결측 -> 0으로 (지급 안 된 항목으로 간주)
    for c in COMP_COLS:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # 보상금 총액
    existing_comp_cols = [c for c in COMP_COLS if c in df.columns]
    df["보상금총액"] = df[existing_comp_cols].sum(axis=1)

    if "학교급" in df.columns:
        df["학교급"] = pd.Categorical(df["학교급"], categories=학교급_순서, ordered=True)

    return df, existing_comp_cols


def make_severity_target(df, method="quantile"):
    """
    보상금총액을 저/중/고 3단계로 변환.
    method:
      - "quantile": 전체 분포 3분위(0~33%, 33~66%, 66~100%)로 구간화
      - "custom": 0원=저 / 0초과~중앙값=중 / 중앙값초과=고  (0원 비중이 클 때 권장)
    """
    total = df["보상금총액"]

    if method == "quantile":
        df["중대성등급"] = pd.qcut(total, q=3, labels=["저", "중", "고"], duplicates="drop")
    else:
        median_nonzero = total[total > 0].median()
        def classify(x):
            if x == 0:
                return "저"
            elif x <= median_nonzero:
                return "중"
            else:
                return "고"
        df["중대성등급"] = total.apply(classify)

    print("\n=== 보상금총액 기술통계 ===")
    print(total.describe())
    print("\n=== 중대성등급 분포 ===")
    print(df["중대성등급"].value_counts())

    return df


# ------------------------------------------------------------------
# 3. 시각화 - 사고데이터 (빈도/패턴)
# ------------------------------------------------------------------
def plot_accident_eda(df):
    # 3-1. 학교급별 사고 건수
    plt.figure(figsize=(6, 4))
    sns.countplot(data=df, x="학교급", order=학교급_순서, palette="Blues_d")
    plt.title("학교급별 사고 건수")
    plt.tight_layout()
    plt.savefig(f"{OUTPUT_DIR}/01_학교급별_건수.png", dpi=150)
    plt.close()

    # 3-2. 요일별 사고 건수
    if "사고요일" in df.columns:
        plt.figure(figsize=(7, 4))
        sns.countplot(data=df, x="사고요일", order=요일_순서, palette="Greens_d")
        plt.title("요일별 사고 건수")
        plt.tight_layout()
        plt.savefig(f"{OUTPUT_DIR}/02_요일별_건수.png", dpi=150)
        plt.close()

    # 3-3. 시간대별 사고 건수
    if "시간대" in df.columns:
        plt.figure(figsize=(8, 4))
        sns.countplot(data=df, x="시간대", palette="Oranges_d")
        plt.title("시간대별 사고 건수")
        plt.tight_layout()
        plt.savefig(f"{OUTPUT_DIR}/03_시간대별_건수.png", dpi=150)
        plt.close()

    # 3-4. 계절별 사고 건수
    if "계절" in df.columns:
        plt.figure(figsize=(6, 4))
        sns.countplot(data=df, x="계절", order=["봄", "여름", "가을", "겨울"], palette="Purples_d")
        plt.title("계절별 사고 건수")
        plt.tight_layout()
        plt.savefig(f"{OUTPUT_DIR}/04_계절별_건수.png", dpi=150)
        plt.close()

    # 3-5. 학교급 x 요일 히트맵 (핵심: 위험요인 교차분석)
    if "사고요일" in df.columns:
        pivot = pd.crosstab(df["학교급"], df["사고요일"])
        pivot = pivot.reindex(index=학교급_순서, columns=요일_순서).fillna(0).astype(int)
        plt.figure(figsize=(8, 4))
        sns.heatmap(pivot, annot=True, fmt="d", cmap="YlOrRd")
        plt.title("학교급 x 요일 사고건수 히트맵")
        plt.tight_layout()
        plt.savefig(f"{OUTPUT_DIR}/05_학교급x요일_히트맵.png", dpi=150)
        plt.close()

    # 3-6. 학교급 x 시간대 히트맵
    if "시간대" in df.columns:
        pivot2 = pd.crosstab(df["학교급"], df["시간대"])
        pivot2 = pivot2.reindex(index=학교급_순서).fillna(0).astype(int)
        plt.figure(figsize=(8, 4))
        sns.heatmap(pivot2, annot=True, fmt="d", cmap="YlOrRd")
        plt.title("학교급 x 시간대 사고건수 히트맵")
        plt.tight_layout()
        plt.savefig(f"{OUTPUT_DIR}/06_학교급x시간대_히트맵.png", dpi=150)
        plt.close()

    # 3-7. 학교급 x 사고형태 히트맵
    if "사고형태" in df.columns:
        pivot3 = pd.crosstab(df["학교급"], df["사고형태"])
        pivot3 = pivot3.reindex(index=학교급_순서).fillna(0).astype(int)
        plt.figure(figsize=(9, 4))
        sns.heatmap(pivot3, annot=True, fmt="d", cmap="YlOrRd")
        plt.title("학교급 x 사고형태 사고건수 히트맵")
        plt.tight_layout()
        plt.savefig(f"{OUTPUT_DIR}/07_학교급x사고형태_히트맵.png", dpi=150)
        plt.close()

    # 3-8. 학교급 x 사고장소 히트맵
    if "사고장소" in df.columns:
        pivot4 = pd.crosstab(df["학교급"], df["사고장소"])
        pivot4 = pivot4.reindex(index=학교급_순서).fillna(0).astype(int)
        plt.figure(figsize=(9, 4))
        sns.heatmap(pivot4, annot=True, fmt="d", cmap="YlOrRd")
        plt.title("학교급 x 사고장소 사고건수 히트맵")
        plt.tight_layout()
        plt.savefig(f"{OUTPUT_DIR}/08_학교급x사고장소_히트맵.png", dpi=150)
        plt.close()

    print(f"\n사고데이터 그래프 {8}개 저장 완료 -> {OUTPUT_DIR}/")


# ------------------------------------------------------------------
# 4. 시각화 - 보상데이터 (중대성)
# ------------------------------------------------------------------
def plot_compensation_eda(df):
    # 4-1. 보상금총액 분포 히스토그램 (0원 많으면 로그스케일도 같이)
    plt.figure(figsize=(7, 4))
    sns.histplot(df["보상금총액"], bins=50, kde=False)
    plt.title("보상금총액 분포")
    plt.tight_layout()
    plt.savefig(f"{OUTPUT_DIR}/09_보상금총액_분포.png", dpi=150)
    plt.close()

    # 4-2. 학교급별 보상금총액 박스플롯
    plt.figure(figsize=(7, 4))
    sns.boxplot(data=df, x="학교급", y="보상금총액", order=학교급_순서)
    plt.title("학교급별 보상금총액 분포")
    plt.tight_layout()
    plt.savefig(f"{OUTPUT_DIR}/10_학교급별_보상금_박스플롯.png", dpi=150)
    plt.close()

    # 4-3. 학교급 x 중대성등급 히트맵
    if "중대성등급" in df.columns:
        pivot = pd.crosstab(df["학교급"], df["중대성등급"])
        pivot = pivot.reindex(index=학교급_순서, columns=["저", "중", "고"]).fillna(0).astype(int)
        plt.figure(figsize=(6, 4))
        sns.heatmap(pivot, annot=True, fmt="d", cmap="Reds")
        plt.title("학교급 x 중대성등급 히트맵")
        plt.tight_layout()
        plt.savefig(f"{OUTPUT_DIR}/11_학교급x중대성등급_히트맵.png", dpi=150)
        plt.close()

    print(f"보상데이터 그래프 3개 저장 완료 -> {OUTPUT_DIR}/")


# ------------------------------------------------------------------
# 5. 실행
# ------------------------------------------------------------------
if __name__ == "__main__":
    # 사고데이터
    df_acc = load_and_clean_accident(ACCIDENT_PATH)
    plot_accident_eda(df_acc)
    df_acc.to_csv(f"{OUTPUT_DIR}/사고데이터_전처리완료.csv", index=False, encoding="utf-8-sig")

    # 보상데이터
    df_comp, comp_cols = load_and_clean_compensation(COMPENSATION_PATH)
    df_comp = make_severity_target(df_comp, method="quantile")  # 0원 비중 크면 "custom"으로 변경
    plot_compensation_eda(df_comp)
    df_comp.to_csv(f"{OUTPUT_DIR}/보상데이터_전처리완료.csv", index=False, encoding="utf-8-sig")

    print("\n전체 완료. output/ 폴더 확인.")