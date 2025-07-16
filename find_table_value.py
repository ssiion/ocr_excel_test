import pandas as pd
import os
import json

def find_case_no_header(df, keyword="Case No."):
    for idx, row in df.iterrows():
        for col_idx, cell in enumerate(row):
            if keyword.lower() in str(cell).strip().lower():
                return idx, col_idx
    return None, None

def extract_multiline_header(df, header_row_idx, header_col_idx, header_above=0, header_below=0):
    header_start = max(0, header_row_idx - header_above)
    header_end = header_row_idx + header_below
    header_rows = header_end - header_start + 1
    header_df = df.iloc[header_start:header_end+1, :]
    headers = []
    n_cols = header_df.shape[1]
    for col in range(header_col_idx, n_cols):
        col_cells = [str(header_df.iloc[i, col]).strip() for i in range(header_rows)]
        merged = " ".join([c for c in col_cells if c and c.lower() != 'nan'])
        if merged:
            headers.append(merged)
    return headers

def normalize_headers(headers):
    seen = {}
    result = []
    for h in headers:
        h_clean = h.strip()
        if not h_clean:
            h_clean = "Unnamed"
        count = seen.get(h_clean, 0)
        if count > 0:
            h_clean = f"{h_clean}_{count+1}"
        seen[h_clean] = count + 1
        result.append(h_clean)
    return result

def extract_table_rows(df, data_start_row, header_col_idx, n_cols, height=None):
    end_row = len(df) if height is None else data_start_row + height
    table = []
    for idx in range(data_start_row, min(end_row, len(df))):
        row = df.iloc[idx]
        
        # ▶ row 전체를 탐색해서 우측 끝까지 실제 값이 있는 최대 열 위치 계산
        last_non_empty_col = header_col_idx + n_cols  # 기본값
        for col_idx in range(header_col_idx + n_cols, len(row)):
            cell_val = str(row.iloc[col_idx]).strip().lower()
            if cell_val and cell_val != 'nan':
                last_non_empty_col = col_idx + 1  # 실제 데이터 존재 시 확장
        
        # ▶ 필요한 만큼 열 확장해서 가져오기
        row_slice = row.iloc[header_col_idx:last_non_empty_col]
        row_values = [str(cell).strip() if str(cell).strip().lower() != 'nan' else '' for cell in row_slice]
        print(f"[extract_table_rows] idx={idx}, row_values={row_values}")
        table.append(row_values)
    return table


def group_data_rows(rows, group_size=2):
    groups = []
    for i in range(0, len(rows), group_size):
        group = rows[i:i+group_size]
        if len(group) < group_size:
            continue

        # 첫 번째 줄 기준으로 실제 데이터가 있는지 확인
        first_row_non_empty = any(cell.strip() for j, cell in enumerate(group[0][1:]))  # Case No. 제외

        if not first_row_non_empty:
            continue  # 전부 비어 있으면 이 그룹 스킵

        merged = []
        num_cols = max(len(r) for r in group)
        for col in range(num_cols):
            values = [group[row_idx][col].strip() if col < len(group[row_idx]) else '' for row_idx in range(group_size)]
            non_empty_values = [v for v in values if v]
            if non_empty_values:
                merged.append(non_empty_values[0])
            else:
                merged.append('')
        groups.append(merged)

    return groups


def extract_table_with_dynamic_header(file_path, keyword, header_above=1, header_below=0, height=None, group_size=2):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
    
    xls = pd.ExcelFile(file_path)
    for sheet_name in xls.sheet_names:
        df = xls.parse(sheet_name).astype(str).fillna('')
        header_row_idx, header_col_idx = find_case_no_header(df, keyword)
        if header_row_idx is None:
            continue

        # 헤더 추출 및 정제
        raw_headers = extract_multiline_header(df, header_row_idx, header_col_idx, header_above, header_below)
        headers = normalize_headers(raw_headers)
        n_cols = len(headers)

        # 데이터 추출 및 그룹 병합
        data_start_row = header_row_idx + header_below + 1
        table = extract_table_rows(df, data_start_row, header_col_idx, n_cols, height)
        grouped_rows = group_data_rows(table, group_size=group_size)

        # 헤더-데이터 매핑 (첫 번째 헤더가 비어있으면 제외)
        result = []
        for row in grouped_rows:
            row_dict = {h: v for h, v in zip(headers, row)}
            if row_dict.get(headers[0], '').strip():
                result.append(row_dict)
        return result

    return []

# ✅ 실행 부분
if __name__ == "__main__":
    file_path = "/Users/zionchoi/Desktop/test_pdf/PEBE00026818-BLC-25210(AIR).xlsx"
    header_above = 1
    header_below = 0
    start_keyword = "Case No."  # 헤더 시작 키워드
    group_size = 2        # 데이터는 2줄씩 병합
    height = None         # 추출할 데이터 높이 (None이면 끝까지)

    result = extract_table_with_dynamic_header(
        file_path=file_path,
        keyword=start_keyword,
        header_above=header_above,
        header_below=header_below,
        height=height,
        group_size=group_size
    )

    print(json.dumps(result, ensure_ascii=False, indent=2))
