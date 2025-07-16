import pandas as pd
import os
import json

def find_case_no_header(df, keyword="Case No."):
    for idx, row in df.iterrows():
        for col_idx, cell in enumerate(row):
            if keyword.lower() in str(cell).strip().lower():
                return idx, col_idx
    return None, None

def extract_multiline_header_with_indices(df, header_row_idx, header_col_idx, header_above=0, header_below=0):
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
            headers.append((merged, col))
    return headers

def get_header_ranges(headers_with_indices, total_cols):
    ranges = []
    for i, (h, start) in enumerate(headers_with_indices):
        end = headers_with_indices[i+1][1] if i+1 < len(headers_with_indices) else total_cols
        ranges.append((h, start, end))
    return ranges

def extract_table_rows(df, data_start_row, header_col_idx, n_cols, height=None):
    end_row = len(df) if height is None else data_start_row + height
    table = []
    for idx in range(data_start_row, min(end_row, len(df))):
        row = df.iloc[idx]
        row_values = [str(cell).strip() if str(cell).strip().lower() != 'nan' else '' for cell in row]
        print(f"[extract_table_rows] idx={idx}, row_values={row_values}")
        table.append(row_values)
    return table

def group_data_rows_by_ranges(rows, group_size, header_ranges):
    groups = []
    for i in range(0, len(rows), group_size):
        group = rows[i:i+group_size]
        if len(group) < group_size:
            continue
        merged = []
        for h, start, end in header_ranges:
            values = []
            for row in group:
                for idx in range(start, end):
                    if idx < len(row):
                        v = row[idx].strip()
                        if v:
                            values.append(v)
            merged_val = ', '.join(values)
            merged.append(merged_val)
        groups.append(merged)
    return groups

def extract_table_with_dynamic_header(file_path, keyword, header_above=1, header_below=0, height=None, group_size=2, header_ranges=None):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
    
    xls = pd.ExcelFile(file_path)
    for sheet_name in xls.sheet_names:
        df = xls.parse(sheet_name).astype(str).fillna('')
        header_row_idx, header_col_idx = find_case_no_header(df, keyword)
        if header_row_idx is None:
            continue

        if header_ranges is not None:
            # 사용자가 직접 header_ranges를 지정한 경우
            header_names = [h for h, _, _ in header_ranges]
        else:
            # 자동 계산
            headers_with_indices = extract_multiline_header_with_indices(df, header_row_idx, header_col_idx, header_above, header_below)
            header_names = [h for h, _ in headers_with_indices]
            header_ranges = get_header_ranges(headers_with_indices, df.shape[1])

        # 데이터 추출
        data_start_row = header_row_idx + header_below + 1
        table = extract_table_rows(df, data_start_row, header_col_idx, len(header_names), height)
        grouped_rows = group_data_rows_by_ranges(table, group_size, header_ranges)

        # 헤더-데이터 매핑 (빈 값은 제외)
        result = []
        for row in grouped_rows:
            row_dict = {h: v for h, v in zip(header_names, row) if v}
            if row_dict.get(header_names[0], '').strip():
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

    # 직접 header_ranges 지정 예시 (헤더명, 시작 인덱스, 끝 인덱스)
    header_ranges = [
        ("Case No.", 0, 1),
        ("Parts No.", 4, 8),
        ("Description", 11, 20),
        ("Q'ty", 20, 26),
        ("Net Weight", 26, 29),
        ("Gross Weight", 29, 32),
        ("Measurement", 32, 40)
    ]

    result = extract_table_with_dynamic_header(
        file_path=file_path,
        keyword=start_keyword,
        header_above=header_above,
        header_below=header_below,
        height=height,
        group_size=group_size,
        header_ranges=header_ranges
    )

    print(json.dumps(result, ensure_ascii=False, indent=2))
