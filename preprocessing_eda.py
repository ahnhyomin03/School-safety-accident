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
import numpy as np
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

# 한글 폰트 설정 (Windows 기준: 맑은 고딕)
import matplotlib.font_manager as fm
import os as _os

WINDOWS_FONT_PATHS = {
    "Malgun Gothic": r"C:\Windows\Fonts\malgun.ttf",
    "MalgunGothic-Bold": r"C:\Windows\Fonts\malgunbd.ttf",
    "NanumGothic": r"C:\Windows\Fonts\NanumGothic.ttf",
    "Gulim": r"C:\Windows\Fonts\gulim.ttc",
    "Batang": r"C:\Windows\Fonts\batang.ttc",
}

selected_font = None

# 1) 폰트 파일 경로를 직접 fontManager에 등록 시도 (Windows 시스템 폰트 캐시 누락 대비)
for font_name, font_path in WINDOWS_FONT_PATHS.items():
    if _os.path.exists(font_path):
        try:
            fm.fontManager.addfont(font_path)
            prop = fm.FontProperties(fname=font_path)
            selected_font = prop.get_name()
            break
        except Exception:
            continue

# 2) 그래도 못 찾으면 기존 fontManager 목록에서 검색
if not selected_font:
    available_fonts = {f.name for f in fm.fontManager.ttflist}
    for font_name in ["Malgun Gothic", "NanumGothic", "Gulim", "Batang"]:
        if font_name in available_fonts:
            selected_font = font_name
            break

if selected_font:
    # font.family를 직접 지정하는 대신, sans-serif 폰트 목록 맨 앞에 추가하는 방식이
    # 더 안정적으로 적용됨 (matplotlib 버전에 따라 font.family 직접 지정이 무시되는 경우 있음)
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = [selected_font] + plt.rcParams.get("font.sans-serif", [])
    print(f"한글 폰트 설정: {selected_font}")

    # matplotlib 폰트 캐시가 오래되어 새로 등록한 폰트를 못 읽는 경우 대비, 캐시 재생성
    try:
        fm._load_fontmanager(try_read_cache=False)
    except Exception:
        pass
else:
    print("경고: 한글 폰트를 찾지 못했습니다. 그래프의 한글이 깨질 수 있습니다.")
    print("나눔고딕(https://hangeul.naver.com/font) 등을 설치한 후 다시 실행해주세요.")

plt.rcParams["axes.unicode_minus"] = False

# 그래프 숫자(축 눈금, 범례 등) 폰트 크기 축소
plt.rcParams["xtick.labelsize"] = 7
plt.rcParams["ytick.labelsize"] = 7
plt.rcParams["axes.labelsize"] = 9
plt.rcParams["legend.fontsize"] = 8
plt.rcParams["figure.titlesize"] = 12
sns.set_style("whitegrid")

학교급_순서 = ["초등학교", "중학교", "고등학교"]
요일_순서 = ["월", "화", "수", "목", "금", "토", "일"]
TOP_N = 8  # 범주 많은 컬럼에서 상위 몇 개까지 개별표시할지 (나머지는 '기타')


def group_top_n(series, n=TOP_N, other_label="기타"):
    """빈도 상위 n개 범주만 남기고 나머지는 '기타'로 묶음"""
    top_categories = series.value_counts().nlargest(n).index
    return series.where(series.isin(top_categories), other_label)


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

    # 사고발생시각: "HH:MM" 형식 5분단위 값이라 고유값이 많으므로
    # 몇 개의 시간대 구간으로 묶어서 사용
    if "사고발생시각" in df.columns:
        # "HH:MM" 문자열 -> 시(hour) 추출. 형식이 다르면(예: datetime.time, 정수) 유연하게 처리
        parsed_time = pd.to_datetime(
            df["사고발생시각"].astype(str), format="%H:%M", errors="coerce"
        )
        # 위 형식으로 안 되면 일반 파서로 재시도
        if parsed_time.isna().all():
            parsed_time = pd.to_datetime(df["사고발생시각"], errors="coerce")

        hour = parsed_time.dt.hour

        def hour_to_period(h):
            if pd.isna(h):
                return "미상"
            elif h < 8:
                return "등교전"
            elif h < 9:
                return "등교시간"
            elif h < 12:
                return "오전수업"
            elif h < 13:
                return "점심시간"
            elif h < 17:
                return "오후수업"
            else:
                return "하교이후"

        df["시간대"] = hour.apply(hour_to_period)

    # 학교급 순서 지정 (그래프용 categorical order)
    if "학교급" in df.columns:
        df["학교급"] = pd.Categorical(df["학교급"], categories=학교급_순서, ordered=True)
    if "사고요일" in df.columns:
        df["사고요일"] = pd.Categorical(df["사고요일"], categories=요일_순서, ordered=True)

    # 범주가 너무 많은 컬럼(사고형태, 사고부위, 사고당시활동)은
    # 빈도 상위 TOP_N개만 개별표시하고 나머지는 '기타'로 묶은 컬럼을 별도 생성
    # (원본 컬럼은 그대로 두고 "_그룹" 컬럼을 시각화에 사용)
    for col in ["사고형태", "사고부위", "사고당시활동", "사고장소"]:
        if col in df.columns:
            df[f"{col}_그룹"] = group_top_n(df[col])
            n_unique = df[col].nunique()
            print(f"{col}: 고유값 {n_unique}개 -> 상위 {TOP_N}개 + 기타로 그룹핑")

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

    # 시간 컬럼(사고시간 또는 사고발생시각) -> 시간대 구간화 (5분단위 등 고유값 많은 경우 대비)
    time_col = None
    for candidate in ["사고시간", "사고발생시각"]:
        if candidate in df.columns:
            time_col = candidate
            break

    if time_col:
        parsed_time = pd.to_datetime(df[time_col].astype(str), format="%H:%M", errors="coerce")
        if parsed_time.isna().all():
            parsed_time = pd.to_datetime(df[time_col], errors="coerce")
        hour = parsed_time.dt.hour

        def hour_to_period(h):
            if pd.isna(h):
                return "미상"
            elif h < 8:
                return "등교전"
            elif h < 9:
                return "등교시간"
            elif h < 12:
                return "오전수업"
            elif h < 13:
                return "점심시간"
            elif h < 17:
                return "오후수업"
            else:
                return "하교이후"

        df["시간대"] = hour.apply(hour_to_period)

    # 범주가 많은 컬럼은 상위 TOP_N개 + 기타로 그룹핑 (독립변수 후보로 사용)
    for col in ["사고장소", "사고부위", "사고형태", "사고당시활동"]:
        if col in df.columns:
            df[f"{col}_그룹"] = group_top_n(df[col])
            print(f"{col}: 고유값 {df[col].nunique()}개 -> 상위 {TOP_N}개 + 기타로 그룹핑")

    return df, existing_comp_cols


def make_severity_5tier(df):
    """
    보상금 항목의 '종류'를 기준으로 5단계 중대성 등급 부여.
    금액이 아니라 보상 항목의 성격(사망/장해/간병/일반치료)을 우선 기준으로 삼아서,
    사망 관련 보상(유족급여·장례비)이 발생하면 금액과 무관하게 최고 등급(5)으로 분류.

    5_사망       : 유족급여 또는 장례비 발생 (최우선, 금액 무관)
    4_중대(장해)  : 장해급여 발생
    3_중등도(간병): 간병급여 발생
    2_경도       : 요양급여+위로금+보전비용 합이 중앙값 이상
    1_경미       : 요양급여+위로금+보전비용 합이 중앙값 미만 (0원 포함)
    """
    for c in ["유족급여", "장례비", "장해급여", "간병급여", "요양급여", "위로금", "보전비용"]:
        if c not in df.columns:
            df[c] = 0

    # 1/2단계 구분 임계값: 상위 등급(사망/장해/간병) 없는 행들의 (요양급여+위로금+보전비용) 중앙값
    base_amount = df["요양급여"] + df["위로금"] + df["보전비용"]
    base_mask = (df["유족급여"] == 0) & (df["장례비"] == 0) & (df["장해급여"] == 0) & (df["간병급여"] == 0)
    threshold = base_amount[base_mask].median()
    if pd.isna(threshold):
        threshold = 0
    print(f"\n1단계/2단계 구분 임계값(요양급여+위로금+보전비용 중앙값): {threshold}")

    def classify(row):
        if row["유족급여"] > 0 or row["장례비"] > 0:
            return "5_사망"
        elif row["장해급여"] > 0:
            return "4_중대(장해)"
        elif row["간병급여"] > 0:
            return "3_중등도(간병)"
        else:
            amount = row["요양급여"] + row["위로금"] + row["보전비용"]
            return "2_경도" if amount >= threshold else "1_경미"

    df["중대성등급"] = df.apply(classify, axis=1)

    print("\n=== 보상금총액 기술통계 ===")
    print(df["보상금총액"].describe())
    print("\n=== 중대성등급(5단계) 분포 ===")
    print(df["중대성등급"].value_counts().sort_index())

    return df


중대성_순서 = ["1_경미", "2_경도", "3_중등도(간병)", "4_중대(장해)", "5_사망"]


# ------------------------------------------------------------------
# 중대성등급과 독립변수(범주형) 간 연관성 분석 (Cramér's V)
# ------------------------------------------------------------------
def cramers_v(x, y):
    """범주형-범주형 변수 간 연관성 강도 (0~1). 피어슨 상관계수의 범주형 버전."""
    from scipy.stats import chi2_contingency
    confusion = pd.crosstab(x, y)
    chi2 = chi2_contingency(confusion)[0]
    n = confusion.sum().sum()
    r, k = confusion.shape
    if n == 0 or r < 2 or k < 2:
        return np.nan
    phi2 = chi2 / n
    phi2corr = max(0, phi2 - ((k - 1) * (r - 1)) / (n - 1))
    rcorr = r - ((r - 1) ** 2) / (n - 1)
    kcorr = k - ((k - 1) ** 2) / (n - 1)
    denom = min((kcorr - 1), (rcorr - 1))
    if denom <= 0:
        return np.nan
    return np.sqrt(phi2corr / denom)


def analyze_severity_correlations(df, target="중대성등급"):
    """중대성등급(종속변수)과 다른 범주형 독립변수들 간의 Cramér's V 연관성 계산"""
    candidate_cols = [
        "지역", "학교급", "사고자구분", "사고자학년", "사고자성별",
        "시간대", "사고장소_그룹", "사고부위_그룹", "사고형태_그룹", "사고당시활동_그룹",
    ]
    predictor_cols = [c for c in candidate_cols if c in df.columns and c != target]

    results = {}
    for col in predictor_cols:
        v = cramers_v(df[target], df[col])
        results[col] = v

    result_df = pd.Series(results).dropna().sort_values(ascending=False)
    print("\n=== 중대성등급과 각 변수의 Cramér's V (연관성 강도, 0~1) ===")
    print(result_df)

    return result_df


def plot_severity_correlations(result_df):
    plt.figure(figsize=(8, max(4, len(result_df) * 0.5)))
    sns.barplot(x=result_df.values, y=result_df.index, palette="viridis", orient="h")
    plt.xlabel("Cramer's V (연관성 강도)")
    plt.title("중대성등급과 독립변수 간 연관성 (Cramer's V)")
    plt.xlim(0, 1)
    plt.tight_layout()
    plt.savefig(f"{OUTPUT_DIR}/14_중대성_연관성_Cramers_V.png", dpi=150)
    plt.close()

    # 연관성 상위 3개 변수는 교차 히트맵도 별도 생성
    return result_df.head(3).index.tolist()


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
        sns.heatmap(pivot, annot=True, fmt="d", cmap="YlOrRd", annot_kws={"size": 6})
        plt.title("학교급 x 요일 사고건수 히트맵")
        plt.tight_layout()
        plt.savefig(f"{OUTPUT_DIR}/05_학교급x요일_히트맵.png", dpi=150)
        plt.close()

    # 3-6. 학교급 x 시간대 히트맵
    시간대_순서 = ["등교전", "등교시간", "오전수업", "점심시간", "오후수업", "하교이후", "미상"]
    if "시간대" in df.columns:
        pivot2 = pd.crosstab(df["학교급"], df["시간대"])
        existing_order = [t for t in 시간대_순서 if t in pivot2.columns]
        pivot2 = pivot2.reindex(index=학교급_순서, columns=existing_order).fillna(0).astype(int)
        plt.figure(figsize=(8, 4))
        sns.heatmap(pivot2, annot=True, fmt="d", cmap="YlOrRd", annot_kws={"size": 6})
        plt.title("학교급 x 시간대 사고건수 히트맵")
        plt.tight_layout()
        plt.savefig(f"{OUTPUT_DIR}/06_학교급x시간대_히트맵.png", dpi=150)
        plt.close()

    # 3-7. 학교급 x 사고형태 히트맵 (상위 8개 + 기타)
    if "사고형태_그룹" in df.columns:
        pivot3 = pd.crosstab(df["학교급"], df["사고형태_그룹"])
        pivot3 = pivot3.reindex(index=학교급_순서).fillna(0).astype(int)
        plt.figure(figsize=(10, 4))
        sns.heatmap(pivot3, annot=True, fmt="d", cmap="YlOrRd", annot_kws={"size": 6})
        plt.title(f"학교급 x 사고형태 사고건수 히트맵 (상위 {TOP_N}개 + 기타)")
        plt.tight_layout()
        plt.savefig(f"{OUTPUT_DIR}/07_학교급x사고형태_히트맵.png", dpi=150)
        plt.close()

    # 3-8. 학교급 x 사고장소 히트맵 (상위 8개 + 기타)
    if "사고장소_그룹" in df.columns:
        pivot4 = pd.crosstab(df["학교급"], df["사고장소_그룹"])
        pivot4 = pivot4.reindex(index=학교급_순서).fillna(0).astype(int)
        plt.figure(figsize=(10, 4))
        sns.heatmap(pivot4, annot=True, fmt="d", cmap="YlOrRd", annot_kws={"size": 6})
        plt.title(f"학교급 x 사고장소 사고건수 히트맵 (상위 {TOP_N}개 + 기타)")
        plt.tight_layout()
        plt.savefig(f"{OUTPUT_DIR}/08_학교급x사고장소_히트맵.png", dpi=150)
        plt.close()

    # 3-9. 학교급 x 사고부위 히트맵 (상위 8개 + 기타)
    if "사고부위_그룹" in df.columns:
        pivot5 = pd.crosstab(df["학교급"], df["사고부위_그룹"])
        pivot5 = pivot5.reindex(index=학교급_순서).fillna(0).astype(int)
        plt.figure(figsize=(10, 4))
        sns.heatmap(pivot5, annot=True, fmt="d", cmap="YlOrRd", annot_kws={"size": 6})
        plt.title(f"학교급 x 사고부위 사고건수 히트맵 (상위 {TOP_N}개 + 기타)")
        plt.tight_layout()
        plt.savefig(f"{OUTPUT_DIR}/09_학교급x사고부위_히트맵.png", dpi=150)
        plt.close()

    # 3-10. 학교급 x 사고당시활동 히트맵 (상위 8개 + 기타)
    if "사고당시활동_그룹" in df.columns:
        pivot6 = pd.crosstab(df["학교급"], df["사고당시활동_그룹"])
        pivot6 = pivot6.reindex(index=학교급_순서).fillna(0).astype(int)
        plt.figure(figsize=(10, 4))
        sns.heatmap(pivot6, annot=True, fmt="d", cmap="YlOrRd", annot_kws={"size": 6})
        plt.title(f"학교급 x 사고당시활동 사고건수 히트맵 (상위 {TOP_N}개 + 기타)")
        plt.tight_layout()
        plt.savefig(f"{OUTPUT_DIR}/10_학교급x사고당시활동_히트맵.png", dpi=150)
        plt.close()

    print(f"\n사고데이터 그래프 저장 완료 -> {OUTPUT_DIR}/")


# ------------------------------------------------------------------
# 4. 시각화 - 보상데이터 (중대성)
# ------------------------------------------------------------------
def plot_compensation_eda(df):
    # 4-1. 보상금총액 분포 히스토그램 (0원 많으면 로그스케일도 같이)
    plt.figure(figsize=(7, 4))
    sns.histplot(df["보상금총액"], bins=50, kde=False)
    plt.title("보상금총액 분포")
    plt.tight_layout()
    plt.savefig(f"{OUTPUT_DIR}/11_보상금총액_분포.png", dpi=150)
    plt.close()

    # 4-2. 학교급별 보상금총액 박스플롯
    plt.figure(figsize=(7, 4))
    sns.boxplot(data=df, x="학교급", y="보상금총액", order=학교급_순서)
    plt.title("학교급별 보상금총액 분포")
    plt.tight_layout()
    plt.savefig(f"{OUTPUT_DIR}/12_학교급별_보상금_박스플롯.png", dpi=150)
    plt.close()

    # 4-3. 학교급 x 중대성등급 히트맵 (5단계)
    if "중대성등급" in df.columns:
        pivot = pd.crosstab(df["학교급"], df["중대성등급"])
        existing_order = [t for t in 중대성_순서 if t in pivot.columns]
        pivot = pivot.reindex(index=학교급_순서, columns=existing_order).fillna(0).astype(int)
        plt.figure(figsize=(8, 4))
        sns.heatmap(pivot, annot=True, fmt="d", cmap="Reds", annot_kws={"size": 6})
        plt.title("학교급 x 중대성등급(5단계) 히트맵")
        plt.tight_layout()
        plt.savefig(f"{OUTPUT_DIR}/13_학교급x중대성등급_히트맵.png", dpi=150)
        plt.close()

    print(f"보상데이터 그래프 저장 완료 -> {OUTPUT_DIR}/")


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
    df_comp = make_severity_5tier(df_comp)
    plot_compensation_eda(df_comp)

    # 중대성등급과 독립변수 간 연관성(Cramer's V) 분석
    corr_result = analyze_severity_correlations(df_comp)
    if len(corr_result) > 0:
        top3 = plot_severity_correlations(corr_result)

        # 연관성 상위 3개 변수는 교차 히트맵도 생성
        for i, col in enumerate(top3, start=1):
            pivot = pd.crosstab(df_comp[col], df_comp["중대성등급"])
            existing_order = [t for t in 중대성_순서 if t in pivot.columns]
            pivot = pivot.reindex(columns=existing_order).fillna(0).astype(int)
            plt.figure(figsize=(8, max(3, len(pivot) * 0.5)))
            sns.heatmap(pivot, annot=True, fmt="d", cmap="Reds", annot_kws={"size": 6})
            plt.title(f"{col} x 중대성등급 히트맵 (연관성 {i}위)")
            plt.tight_layout()
            plt.savefig(f"{OUTPUT_DIR}/1{4+i}_{col}x중대성등급_히트맵.png", dpi=150)
            plt.close()

        corr_result.to_csv(f"{OUTPUT_DIR}/중대성_연관성_CramersV.csv", encoding="utf-8-sig")

    df_comp.to_csv(f"{OUTPUT_DIR}/보상데이터_전처리완료.csv", index=False, encoding="utf-8-sig")

    print("\n전체 완료. output/ 폴더 확인.")
