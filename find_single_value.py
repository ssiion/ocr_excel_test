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

# column mode
def extract_box_column(df, header_row_idx, col_idx, offset=0, x=1, y=None):
    lines = []
    started = False
    start_col = col_idx + offset
    end_col = df.shape[1] if x is None else start_col + x
    start_row = header_row_idx + 1
    end_row = len(df) if y is None else start_row + y
    for idx in range(start_row, min(end_row, len(df))):
        row = df.iloc[idx, start_col:end_col]
        row_has_value = False
        for cell in row:
            cell_value = str(cell).strip()
            if cell_value and cell_value.lower() not in ["nan", "-", ""]:
                row_has_value = True
                lines.append(cell_value)
        if not started and row_has_value:
            started = True
        if started and not row_has_value:
            break
    return lines

# inline mode (x, y는 무시)
def extract_inline_below_header(df, header_row_idx, x=None, y=None):
    values = []
    if header_row_idx + 1 < len(df):
        row = df.iloc[header_row_idx + 1]
        for cell in row:
            cell_value = str(cell).strip()
            if cell_value and cell_value.lower() not in ["nan", "-", ""]:
                values.append(cell_value)
    return values

# row mode (x: 오른쪽 열 개수, y: 아래 행 개수)
def extract_row_right_of_header(df, header_row_idx, header_col_idx, offset=1, x=1, y=None):
    values = []
    start_col = header_col_idx + offset
    end_col = df.shape[1] if x is None else start_col + x
    start_row = header_row_idx + 1
    end_row = len(df) if y is None else start_row + y
    for idx in range(start_row, min(end_row, len(df))):
        row = df.iloc[idx, start_col:end_col]
        for cell in row:
            cell_value = str(cell).strip()
            if not cell_value or cell_value.lower() in ["nan", "-", ""]:
                return values
            values.append(cell_value)
    return values

# row_same mode (x: 오른쪽 열 개수, y: 같은 행만, y>1이면 같은 행부터 y개 행까지)
def extract_row_right_of_header_same_row(df, header_row_idx, header_col_idx, offset=1, x=1, y=None):
    values = []
    start_col = header_col_idx + offset
    end_col = df.shape[1] if x is None else start_col + x
    start_row = header_row_idx
    end_row = len(df) if y is None else start_row + y
    for idx in range(start_row, min(end_row, len(df))):
        row = df.iloc[idx, start_col:end_col]
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
                offset = conf.get("offset", 0)
                x = conf.get("x", 1)
                y = conf.get("y", None)
                if mode == "column":
                    values = extract_box_column(df, row_idx, col_idx, offset=offset, x=x, y=y)
                elif mode == "inline":
                    values = extract_inline_below_header(df, row_idx, x=x, y=y)
                elif mode == "row":
                    values = extract_row_right_of_header(df, row_idx, col_idx, offset=offset, x=x, y=y)
                elif mode == "row_same":
                    values = extract_row_right_of_header_same_row(df, row_idx, col_idx, offset=offset, x=x, y=y)
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
    # file_path = "/Users/zionchoi/Desktop/test_pdf/example_excel.xlsx"
    file_path = "/Users/zionchoi/Desktop/test_pdf/6019  250623  SOLID(HAK)WOOYOUNGMI AIR 日東.xlsx"

    targets = {
        "row_mode": {
            "keywords": ["row mode"],
            "mode": "row",
            "offset": 1,
            "x": 1,
            "y": 4
        },
        "row_mode_same": {
            "keywords": ["row_same mode"],
            "mode": "row_same",
            "offset": 1,
            "x": 1,
            "y": 2
        },
        "column_mode": {
            "keywords": ["column mode"],
            "mode": "column",
            "offset": 1,
            "x": 1,
            "y": 4
        },
        "inline_mode": {
            "keywords": ["inline mode"],
            "mode": "inline"
        },
        "shipper": {
            "keywords": ["shipper", "shipper/exporter", "exporter"],
            "mode": "column",
            "offset": 1,
            "x": 1,
            "y": 4
        },
        "consignee": {
            "keywords": ["consignee", "consignee/importer", "consigee"],
            "mode": "column",
            "offset": 1,
            "x": 1,
            "y": 4
        },
        "depatrure": {
            "keywords": ["depatrure"],
            "mode": "row_same",
            "offset": 1,     # CODE NO.가 notify 오른쪽 첫 칸이면 1
            "x": 3,          # 3칸만 긁고 싶으면 3, 끝까지는 None
            "y": 1
        },
        "invoice_no": {
            "keywords": ["invoice no", "inv no"],
            "mode": "row_same",
            "offset": 1,
            "x": 2,
            "y": 1
        },
        "notify_all": {
            "keywords": ["notify", "notify party"],
            "mode": "column"  # header 아래 한 줄 전체 값 추출
        }
    }
    result = extract_multi_targets(file_path, targets)
    info = next(iter(result.values()))
    print(json.dumps(info, ensure_ascii=False, indent=2))
