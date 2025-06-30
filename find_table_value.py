import pandas as pd
import json
import re

def normalize_col(col):
    # 소문자, 공백/특수문자 제거
    return re.sub(r'[^a-zA-Z0-9가-힣]', '', str(col)).lower()

def find_best_column(columns, candidates):
    # candidates: 튜플 후보들의 리스트
    for candidate in candidates:
        if candidate in columns:
            return candidate
    
    # 정확한 매치가 없으면 부분 매치 시도
    for candidate in candidates:
        candidate_upper = normalize_col(candidate[0]) if candidate[0] else ""
        candidate_lower = normalize_col(candidate[1]) if candidate[1] else ""
        
        for col in columns:
            if isinstance(col, tuple):
                col_upper = normalize_col(col[0]) if col[0] else ""
                col_lower = normalize_col(col[1]) if col[1] else ""
                
                # 상위 헤더가 매치되고, 하위 헤더가 비어있거나 매치되는 경우
                if candidate_upper and candidate_upper in col_upper:
                    if not candidate_lower or candidate_lower in col_lower:
                        return col
            else:
                col_norm = normalize_col(col)
                if candidate_upper and candidate_upper in col_norm:
                    return col
    return None

def safe_get(row, col):
    if col is None:
        return ""
    
    try:
        # MultiIndex 컬럼인 경우
        if isinstance(col, tuple):
            if col in row.index:
                val = row[col]
            else:
                return ""
        else:
            if col in row.index:
                val = row[col]
            else:
                return ""
        
        # Series나 DataFrame이 반환된 경우 첫 번째 값만 가져오기
        if isinstance(val, pd.Series):
            val = val.iloc[0] if not val.empty else ""
        elif isinstance(val, pd.DataFrame):
            val = val.iloc[0, 0] if not val.empty else ""
        
        # NaN 체크 및 문자열 변환
        if pd.isna(val):
            return ""
        
        return str(val).strip()
    except (KeyError, IndexError, AttributeError):
        return ""

def find_table_value(df):
    # 1. 정확한 헤더 시작 위치 찾기 (첫 번째 컬럼이 "CASE No."인 행)
    header_start_idx = None
    for idx, row in df.iterrows():
        first_col = str(row.iloc[0]).strip()
        if "CASE NO" in first_col.upper() or "CASE NO." in first_col.upper():
            header_start_idx = idx
            break
    
    if header_start_idx is None:
        return []
    
    base_idx = header_start_idx

    # 2. 상위/하위 헤더 병합 (정확한 2줄 헤더 구조)
    header1 = df.iloc[base_idx].fillna("").astype(str)
    header2 = df.iloc[base_idx + 1].fillna("").astype(str) if base_idx + 1 < len(df) else pd.Series([""]*len(header1))
    
    # 헤더 병합 시 빈 값 처리
    merged_headers = []
    for a, b in zip(header1, header2):
        a_clean = a.strip()
        b_clean = b.strip()
        if b_clean and a_clean:
            merged_headers.append((a_clean, b_clean))
        elif a_clean:
            merged_headers.append((a_clean, ""))
        else:
            merged_headers.append(("", b_clean))
    print("병합된 헤더:", merged_headers)
    
    df_data = df.iloc[base_idx + 2:].copy()
    df_data.columns = pd.MultiIndex.from_tuples(merged_headers, names=["upper", "lower"])
    df_data.reset_index(drop=True, inplace=True)

    # 3. key map 정의 (실제 엑셀 헤더와 정확히 매칭)
    key_map = {
        "case_no": [("CASE No", "")],
        "package.style": [("Package", "Style")],
        "description.contract_no": [("Description", "Contract No．")],
        "description.por_no": [("", "POR No.")],
        "description.eng_model": [("", "Eng.Model")],
        "description.company_serial": [("", "弊社工番")],
        "description.drw_no": [("", "Drw．No．")],
        "description.parts_name": [("", "Parts Of Name")],
        "description.qty": [("", "Q'ty")],
        "description.price": [("", "Price(￥)")],
        "description.amount": [("", "Amount(￥)")],
        "n_w.kgs": [("N/W", "(kgs)")],
        "g_w.kgs": [("G/W", "(kgs)")],
        "dimension.l": [("Dimension(ｃｍ）", "Ｌ")],
        "dimension.w": [("", "Ｗ")],
        "dimension.h": [("", "Ｈ")],
        "mment.m3": [("M'ment", "(m3)")]
    }

    columns = list(df_data.columns)
    print("실제 columns:", columns)
    col_case_no = find_best_column(columns, key_map["case_no"])
    if col_case_no is None:
        return []

    # 4. 유효 행만 필터 (CASE No.가 있는 행만)
    if isinstance(col_case_no, tuple):
        col_series = df_data.loc[:, [col_case_no]].squeeze()
    else:
        col_series = df_data[col_case_no]
    if isinstance(col_series, pd.DataFrame):
        col_series = col_series.iloc[:, 0]
    col_series = col_series.astype(str).str.strip()

    # "SUB TOTAL", "TOTAL" 등 합계 행 제외
    valid_mask = (col_series != "") & (col_series.str.lower() != "nan")
    invalid_keywords = ["sub total", "subtotal", "t o t a l", "total"]
    for kw in invalid_keywords:
        valid_mask &= ~col_series.str.lower().str.contains(kw)
    valid_rows = df_data[valid_mask.values]

    # 5. row to dict
    result = []
    for _, row in valid_rows.iterrows():
        item = {}
        item["case_no"] = safe_get(row, col_case_no)
        col_package_style = find_best_column(columns, key_map["package.style"])
        item["package"] = {"style": safe_get(row, col_package_style)}

        desc = {}
        for subkey in [
            "contract_no", "por_no", "eng_model", "company_serial", "drw_no", "parts_name", "qty", "price", "amount"
        ]:
            map_key = f"description.{subkey}"
            col = find_best_column(columns, key_map[map_key])
            desc[subkey] = safe_get(row, col)
        item["description"] = desc

        col_nw = find_best_column(columns, key_map["n_w.kgs"])
        item["n_w"] = {"kgs": safe_get(row, col_nw)}
        col_gw = find_best_column(columns, key_map["g_w.kgs"])
        item["g_w"] = {"kgs": safe_get(row, col_gw)}

        col_l = find_best_column(columns, key_map["dimension.l"])
        col_w = find_best_column(columns, key_map["dimension.w"])
        col_h = find_best_column(columns, key_map["dimension.h"])
        item["dimension"] = {
            "l": safe_get(row, col_l),
            "w": safe_get(row, col_w),
            "h": safe_get(row, col_h),
        }

        col_m3 = find_best_column(columns, key_map["mment.m3"])
        item["mment"] = {"m3": safe_get(row, col_m3)}

        result.append(item)
    return result

if __name__ == "__main__":
    file_path = "/Users/zionchoi/Desktop/test_pdf/HHI24-152_20241010.xlsx"
    try:
        # 엑셀 파일을 DataFrame으로 읽기
        xls = pd.ExcelFile(file_path)
        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name, header=None)
            result = find_table_value(df)
            if result:
                print(json.dumps(result, ensure_ascii=False, indent=2))
                break
        else:
            print(json.dumps([], ensure_ascii=False, indent=2))
    except FileNotFoundError as e:
        print(f"오류: {e}")
    except Exception as e:
        print(f"예상치 못한 오류가 발생했습니다: {e}")
