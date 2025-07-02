import pandas as pd
import os
import json

def find_all_header_locations(df, keywords):
    found = []
    for idx, row in df.iterrows():
        for col_idx, cell in enumerate(row):
            cell_str = str(cell).lower()
            for keyword in keywords:
                if keyword in cell_str:
                    found.append((keyword, idx, col_idx))
    return found

def extract_box_column(df, header_row_idx, col_idx):
    lines = []
    started = False
    for idx in range(header_row_idx + 1, len(df)):
        cell_value = str(df.iloc[idx, col_idx]).strip()
        if not started:
            if cell_value and cell_value.lower() not in ["nan", "-", ""]:
                started = True
                lines.append(cell_value)
            continue
        if not cell_value or cell_value.lower() in ["nan", "-", ""]:
            break
        lines.append(cell_value)
    return lines

def extract_inline_below_header(df, header_row_idx):
    values = []
    if header_row_idx + 1 < len(df):
        row = df.iloc[header_row_idx + 1]
        for cell in row:
            cell_value = str(cell).strip()
            if cell_value and cell_value.lower() not in ["nan", "-", ""]:
                values.append(cell_value)
    return values

def extract_row_right_of_header(df, header_row_idx, header_col_idx, offset=1, width=None):
    """
    헤더 아래 행에서 (header_col_idx+offset) ~ (header_col_idx+offset+width)까지 모든 열 추출
    - offset: 몇 칸 오른쪽부터 시작할지 (default=1, 즉 바로 옆)
    - width: 몇 칸을 추출할지 (None이면 끝까지)
    """
    values = []
    start_col = header_col_idx + offset
    end_col = df.shape[1] if width is None else start_col + width
    if header_row_idx + 1 < len(df):
        row = df.iloc[header_row_idx + 1, start_col:end_col]
        for cell in row:
            cell_value = str(cell).strip()
            if cell_value and cell_value.lower() not in ["nan", "-", ""]:
                values.append(cell_value)
    return values

def extract_row_right_of_header_same_row(df, header_row_idx, header_col_idx, offset=1, width=None):
    """
    헤더와 같은 행에서 header_col_idx+offset부터 width만큼 오른쪽 데이터 추출
    """
    values = []
    start_col = header_col_idx + offset
    end_col = df.shape[1] if width is None else start_col + width
    row = df.iloc[header_row_idx, start_col:end_col]
    for cell in row:
        cell_value = str(cell).strip()
        if cell_value and cell_value.lower() not in ["nan", "-", ""]:
            values.append(cell_value)
    return values

def extract_multi_targets(file_path, targets):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
    xls = pd.ExcelFile(file_path)
    result = {}
    for sheet_name in xls.sheet_names:
        df = xls.parse(sheet_name).astype(str)
        info = {}
        for key, conf in targets.items():
            found_locs = find_all_header_locations(df, conf['keywords'])
            all_values = []
            for _, row_idx, col_idx in found_locs:
                mode = conf.get("mode", "column")
                if mode == "column":
                    values = extract_box_column(df, row_idx, col_idx)
                elif mode == "inline":
                    values = extract_inline_below_header(df, row_idx)
                elif mode == "row":
                    offset = conf.get("offset", 1)
                    width = conf.get("width", None)
                    values = extract_row_right_of_header(df, row_idx, col_idx, offset=offset, width=width)
                elif mode == "row_same":
                    offset = conf.get("offset", 1)
                    width = conf.get("width", None)
                    values = extract_row_right_of_header_same_row(df, row_idx, col_idx, offset=offset, width=width)
                else:
                    values = []
                for v in values:
                    if v not in all_values:
                        all_values.append(v)
            info[key] = all_values
        result[sheet_name] = info
        break
    return result

if __name__ == "__main__":
    file_path = "/Users/zionchoi/Desktop/test_pdf/PIJ-24-566(547349).xlsx"
    targets = {
        "shipper": {
            "keywords": ["shipper", "shipper/exporter", "exporter"],
            "mode": "column"
        },
        "consignee": {
            "keywords": ["consignee", "consignee/importer", "consigee"],
            "mode": "column"
        },
        "depatrure": {
            "keywords": ["depatrure"],
            "mode": "row_same",
            "offset": 1,     # CODE NO.가 notify 오른쪽 첫 칸이면 1
            "width": 3       # 3칸만 긁고 싶으면 3, 끝까지는 None
        },
        "invoice_no": {
            "keywords": ["invoice no"],
            "mode": "row_same",
            "offset": 1,
            "width": 2
        },
        "notify_all": {
            "keywords": ["notify", "notify party"],
            "mode": "column"  # header 아래 한 줄 전체 값 추출
        }
    }
    result = extract_multi_targets(file_path, targets)
    info = next(iter(result.values()))
    print(json.dumps(info, ensure_ascii=False, indent=2))
