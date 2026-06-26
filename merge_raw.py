"""
루저스클럽 채널별 로우파일 → 통합 매출자료 변환 스크립트

사용법:
  1. 이 파일이 있는 폴더에 'raw' 폴더를 만드세요
  2. raw 폴더에 아래 파일명으로 각 채널 로우파일을 넣으세요:
       무신사.xls          ← 무신사 판매자센터 주문 내역
       무신사글로벌.xls    ← 무신사글로벌 주문 내역
       자사몰.csv          ← 자사몰(cafe24) 주문 내역 CSV
       29cm.xlsx           ← 29CM 주문상품 내역
  3. 터미널에서 실행:
       py merge_raw.py
  4. data/루저스클럽_통합매출.xlsx 파일이 생성됩니다

* 기존 통합 파일(루저스클럽_매출자료_V*.xlsx)이 있으면 자동으로 합산합니다
"""

import pandas as pd
import numpy as np
from pathlib import Path
from lxml import etree
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

BASE_DIR = Path(__file__).parent
RAW_DIR = BASE_DIR / "raw"
DATA_DIR = BASE_DIR / "data" if (BASE_DIR / "data").exists() else BASE_DIR.parent / "매출자료"
OUTPUT_FILE = DATA_DIR / "루저스클럽_통합매출.xlsx"

# 취소/반품으로 간주할 키워드
CANCEL_KEYWORDS = ['취소', '반품', '교환완료', '클레임완료']


def is_cancelled(status_val):
    if pd.isna(status_val): return False
    s = str(status_val)
    return any(k in s for k in CANCEL_KEYWORDS)


def read_spreadsheet_xml(filepath):
    """무신사/무신사글로벌 .xls (SpreadsheetML XML) 파싱"""
    with open(filepath, 'rb') as f:
        content = f.read()
    if content.startswith(b'\xef\xbb\xbf'):
        content = content[3:]
    content = content.replace(b'<?mso-application progid="Excel.Sheet"?>', b'')
    root = etree.fromstring(content)
    ns = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}
    ws = root.find('.//ss:Worksheet', ns)
    rows = []
    for row in ws.findall('.//ss:Row', ns):
        row_data = []
        for cell in row.findall('ss:Cell', ns):
            data = cell.find('ss:Data', ns)
            row_data.append(data.text if data is not None else '')
        rows.append(row_data)
    if not rows:
        return pd.DataFrame()
    header = rows[0]
    return pd.DataFrame(rows[1:], columns=header)


def load_musinsa(filepath):
    """무신사 주문 내역 → 표준 포맷"""
    df = read_spreadsheet_xml(filepath)
    if df.empty:
        print("  ⚠️  무신사: 데이터 없음")
        return pd.DataFrame()

    # 취소/반품 제거
    if '주문상태' in df.columns:
        df = df[~df['주문상태'].apply(is_cancelled)]
    if '클레임상태' in df.columns:
        df = df[~df['클레임상태'].apply(lambda x: '반품' in str(x) or '교환' in str(x))]

    # 숫자 변환
    for col in ['수량', '정상가', '판매가', '매출금액']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df = df[df['수량'] > 0].reset_index(drop=True)

    out = pd.DataFrame({
        '판매처': '무신사',
        '판매일자': pd.to_datetime(df['주문일시'].str[:10], errors='coerce'),
        '상품명': df['상품명'].values,
        '상품코드': df.get('스타일넘버', pd.Series([''] * len(df))).values,
        '수량': df['수량'].astype(int).values,
        '택가': df['정상가'].values,
        '판매금액': df['매출금액'].values,
    })

    print(f"  ✅ 무신사: {len(out):,}건 로드")
    return out


def load_musinsa_global(filepath):
    """무신사글로벌 주문 내역 → 표준 포맷"""
    df = read_spreadsheet_xml(filepath)
    if df.empty:
        print("  ⚠️  무신사글로벌: 데이터 없음")
        return pd.DataFrame()

    # 첫 빈 컬럼 제거
    if df.columns[0] == '':
        df = df.iloc[:, 1:]

    # 취소/반품 제거
    if '주문상태' in df.columns:
        df = df[~df['주문상태'].apply(is_cancelled)]

    # 숫자 변환
    for col in ['수량', '판매가', '구매금액', '결제금액(원화)', '결제환율']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df = df[df['수량'] > 0].reset_index(drop=True)

    # 국내 택가 역산: (외화판매가 × 환율) / 1.2
    exchange = df['결제환율'].replace(0, np.nan)
    foreign_price = df['판매가']

    out = pd.DataFrame({
        '판매처': '무신사글로벌',
        '판매일자': pd.to_datetime(df['주문일시'].str[:10], errors='coerce'),
        '상품명': df['상품명'].values,
        '상품코드': df.get('스타일넘버', pd.Series([''] * len(df))).values,
        '수량': df['수량'].astype(int).values,
        '택가': (foreign_price * exchange / 1.2).round(0).values,
        '판매금액': df['결제금액(원화)'].values,
    })

    print(f"  ✅ 무신사글로벌: {len(out):,}건 로드")
    return out


def load_jasam(filepath):
    """자사몰 CSV → 표준 포맷"""
    try:
        df = pd.read_csv(filepath, encoding='utf-8-sig')
    except Exception:
        df = pd.read_csv(filepath, encoding='cp949')

    if df.empty:
        print("  ⚠️  자사몰: 데이터 없음")
        return pd.DataFrame()

    # 취소/반품 제거
    if '주문 상태' in df.columns:
        df = df[~df['주문 상태'].apply(is_cancelled)]

    # 숫자 변환
    for col in ['수량', '판매가', '상품구매금액(KRW)']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df = df[df['수량'] > 0].reset_index(drop=True)

    상품명_col = df.get('상품명(한국어 쇼핑몰)', df.get('상품명', pd.Series([''] * len(df))))
    out = pd.DataFrame({
        '판매처': '자사몰',
        '판매일자': pd.to_datetime(df['주문일시'].str[:10], errors='coerce'),
        '상품명': 상품명_col.values if hasattr(상품명_col, 'values') else 상품명_col,
        '상품코드': df.get('자체 상품코드', pd.Series([''] * len(df))).values,
        '수량': df['수량'].astype(int).values,
        '택가': df['판매가'].values,
        '판매금액': df['상품구매금액(KRW)'].values,
    })

    print(f"  ✅ 자사몰: {len(out):,}건 로드")
    return out


def load_29cm(filepath):
    """29CM 주문상품 내역 → 표준 포맷"""
    try:
        df = pd.read_excel(filepath)
    except Exception as e:
        print(f"  ⚠️  29CM 파일 오류: {e}")
        return pd.DataFrame()

    if df.empty:
        print("  ⚠️  29CM: 데이터 없음")
        return pd.DataFrame()

    # 취소/반품 제거
    if '주문상태' in df.columns:
        df = df[~df['주문상태'].apply(is_cancelled)]

    # 숫자 변환
    for col in ['수량', '판매가', '실 판매액', '판매액']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df = df[df['수량'] > 0].reset_index(drop=True)

    style_code = df['옵션코드'].values if '옵션코드' in df.columns else [''] * len(df)
    판매금액_col = '실 판매액' if '실 판매액' in df.columns else '판매액'

    out = pd.DataFrame({
        '판매처': '29CM',
        '판매일자': pd.to_datetime(df['주문일시'].str[:10], errors='coerce'),
        '상품명': df['상품명'].values,
        '상품코드': style_code,
        '수량': df['수량'].astype(int).values,
        '택가': df['판매가'].values,
        '판매금액': df[판매금액_col].values,
    })

    print(f"  ✅ 29CM: {len(out):,}건 로드")
    return out


def add_date_columns(df):
    """날짜 파생 컬럼 추가"""
    df['판매일자'] = pd.to_datetime(df['판매일자'], errors='coerce')
    df['판매년도'] = df['판매일자'].dt.year
    df['월'] = df['판매일자'].dt.month
    df['주차'] = df['판매일자'].dt.isocalendar().week.astype(int)
    return df


def load_existing_data():
    """기존 통합 파일 로드"""
    files = sorted(DATA_DIR.glob("루저스클럽_매출자료_V*.xlsx"), key=lambda x: x.name, reverse=True)
    if not files:
        return pd.DataFrame()
    f = files[0]
    df = pd.read_excel(f)
    df = df.rename(columns={'판매처': '판매처'})
    print(f"  ✅ 기존 파일 ({f.name}): {len(df):,}건 로드")
    return df


def merge_all():
    print("\n🔄 루저스클럽 매출자료 통합 시작")
    print(f"   로우파일 폴더: {RAW_DIR}")
    print(f"   출력 파일: {OUTPUT_FILE}\n")

    parts = []

    # 기존 통합 파일 (베이스)
    print("📂 기존 통합 파일 로드...")
    existing = load_existing_data()
    if not existing.empty:
        # 최소 컬럼 정규화
        if '판매처' in existing.columns and '판매일자' in existing.columns:
            parts.append(existing)

    # 로우파일 로드
    if RAW_DIR.exists():
        print("\n📂 채널별 로우파일 로드...")

        ms_file = RAW_DIR / "무신사.xls"
        if ms_file.exists():
            try:
                parts.append(load_musinsa(str(ms_file)))
            except Exception as e:
                print(f"  ❌ 무신사 오류: {e}")

        gl_file = RAW_DIR / "무신사글로벌.xls"
        if gl_file.exists():
            try:
                parts.append(load_musinsa_global(str(gl_file)))
            except Exception as e:
                print(f"  ❌ 무신사글로벌 오류: {e}")

        js_file = RAW_DIR / "자사몰.csv"
        if js_file.exists():
            try:
                parts.append(load_jasam(str(js_file)))
            except Exception as e:
                print(f"  ❌ 자사몰 오류: {e}")

        cm_file = RAW_DIR / "29cm.xlsx"
        if cm_file.exists():
            try:
                parts.append(load_29cm(str(cm_file)))
            except Exception as e:
                print(f"  ❌ 29CM 오류: {e}")
    else:
        print(f"\n⚠️  raw 폴더가 없습니다. {RAW_DIR} 폴더를 만들고 파일을 넣어주세요.")

    if not parts:
        print("\n❌ 로드된 데이터가 없습니다.")
        return

    # 기존 데이터 마지막 날짜 파악 → 그 이후 raw 데이터만 추가
    if not existing.empty and '판매일자' in existing.columns:
        existing['판매일자'] = pd.to_datetime(existing['판매일자'], errors='coerce')
        cutoff_date = existing['판매일자'].max()
        print(f"\n📅 기존 데이터 마지막 날짜: {cutoff_date.strftime('%Y-%m-%d')}")
        print(f"   → 이후 날짜의 로우파일 데이터만 추가합니다\n")
    else:
        cutoff_date = None

    # raw 파일 부분만 날짜 필터 적용
    raw_parts = parts[1:] if not existing.empty else parts  # 첫번째는 기존 파일
    raw_combined_parts = []
    for p in raw_parts:
        if p.empty: continue
        p = p.reset_index(drop=True)
        p['판매일자'] = pd.to_datetime(p['판매일자'], errors='coerce')
        ch_name = p['판매처'].dropna().iloc[0] if not p['판매처'].dropna().empty else '알수없음'
        if cutoff_date is not None:
            p_new = p[p['판매일자'] > cutoff_date]
            if len(p_new) > 0:
                print(f"  + 신규 {ch_name}: {len(p_new):,}건 추가 (cutoff 이후)")
            else:
                print(f"  - {ch_name}: 새 데이터 없음 (이미 포함됨)")
        else:
            p_new = p.copy()
        raw_combined_parts.append(p_new)

    # 최종 병합
    all_parts = []
    if not existing.empty:
        all_parts.append(existing)
    all_parts.extend([p for p in raw_combined_parts if not p.empty])

    if not all_parts:
        print("\n no data to merge.")
        return

    final = pd.concat(all_parts, ignore_index=True)
    final = add_date_columns(final)

    # 컬럼 순서 정리
    cols = ['판매일자', '판매년도', '월', '주차', '판매처', '상품명', '상품코드', '수량', '택가', '판매금액']
    for c in cols:
        if c not in final.columns:
            final[c] = ''
    final = final[cols]
    final = final.sort_values('판매일자').reset_index(drop=True)

    # 저장
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    final.to_excel(str(OUTPUT_FILE), index=False)

    print(f"Done! total {len(final)} rows -> {OUTPUT_FILE}")
    ch_counts = final['판매처'].value_counts().to_dict()
    print(f"Channels: {ch_counts}")



if __name__ == "__main__":
    merge_all()
