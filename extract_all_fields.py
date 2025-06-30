import pandas as pd
import os
import json
import warnings

# applymap 경고 무시
warnings.filterwarnings('ignore', category=FutureWarning, message='.*applymap.*')

def extract_all_fields(df):
    """
    엑셀 데이터에서 다양한 필드들을 추출합니다.
    - Shipper/Consignee: 여러 줄 정보
    - 기타 필드: 키워드 옆 셀의 값
    """
    # applymap 대신 map 사용
    df_str = df.astype(str).fillna("").map(lambda x: x.strip().lower())

    # 1. Shipper / Consignee 추출 (여러 줄)
    def extract_multiline(keyword):
        for idx, row in df.iterrows():
            joined = " ".join(row.fillna("").astype(str).str.lower().tolist())
            if keyword in joined:
                lines = []

                # 1. 같은 행에서 키워드 오른쪽 셀들의 정보 추출
                row_cells = [cell if str(cell).lower() != "nan" else "" for cell in row.fillna("").tolist()]
                # 키워드가 있는 셀의 인덱스를 찾기
                keyword_col_idx = None
                for col_idx, cell in enumerate(row_cells):
                    if keyword in str(cell).lower():
                        keyword_col_idx = col_idx
                        break

                # 키워드 오른쪽 셀들의 정보 추출 (최대 3칸까지)
                if keyword_col_idx is not None:
                    for col_offset in range(1, 4):  # 오른쪽 1~3칸
                        if keyword_col_idx + col_offset < len(row_cells):
                            cell_value = str(row_cells[keyword_col_idx + col_offset]).strip()
                            if cell_value and cell_value.lower() != "nan" and keyword not in cell_value.lower():
                                lines.append(cell_value)

                # 2. 키워드 아래 행들의 정보 추출
                for offset in range(1, 5):
                    if idx + offset >= len(df): break
                    row_cells = [cell if str(cell).lower() != "nan" else "" for cell in df.iloc[idx + offset].fillna("").tolist()]
                    line = " ".join(row_cells).strip()
                    if line and keyword not in line.lower():
                        lines.append(line)
                    elif not line:
                        break
                return lines
        return []

    shipper_info = extract_multiline("shipper")
    consignee_info = extract_multiline("consignee")

    # 2. 나머지 필드 추출 (한 셀 옆)
    keyword_map = {
        "invoice_no": ["invoice no", "invoice number", "inv no", "請求書番号", "invoice no."],
        "payment": ["payment", "支払い", "terms", "payment term"],
        "freight": ["freight", "運賃", "shipping method"],
        "airport": ["airport", "空港", "成田空港"],
        "invoice_date": ["invoice date", "弊社出荷日", "出荷日"],
        "arrival_date": ["arrival date", "御社搬入日", "到着日"]
    }

    simple_fields = {}
    # 1. 키워드 위치 저장
    keyword_positions = {}
    for field, keywords in keyword_map.items():
        for keyword in keywords:
            for idx, row in df.iterrows():
                for col_name in df.columns:
                    cell_str = str(row[col_name]).lower().strip().replace("：", "").replace(":", "")
                    if keyword in cell_str:
                        keyword_positions[field] = (idx, df.columns.get_loc(col_name))
                        break
                if field in keyword_positions:
                    break
            if field in keyword_positions:
                break

    # 2. 위치 기반 오른쪽 값 추출
    for field, (row_idx, col_idx) in keyword_positions.items():
        row_values = df.iloc[row_idx].tolist()
        for offset in range(col_idx + 1, len(row_values)):
            value = str(row_values[offset]).strip()
            if value.lower() != "nan" and value:
                # '; ' 제거
                value = value.replace('; ', '').replace('；', '').strip()
                if value:  # 제거 후에도 값이 있으면 저장
                    simple_fields[field] = value
                    break

    return {
        "shipper": shipper_info,
        "consignee": consignee_info,
        **simple_fields
    }

def extract_from_excel(file_path):
    """
    엑셀 파일에서 모든 필드를 추출합니다.
    """
    # 파일 존재 여부 확인
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
    
    # ExcelFile 객체로 읽기
    xls = pd.ExcelFile(file_path)
    result = {}

    for sheet_name in xls.sheet_names:
        # 모든 셀을 문자열로 처리
        df = xls.parse(sheet_name)
        extracted_data = extract_all_fields(df)
        
        # 데이터가 있는 경우에만 결과에 추가
        if any(extracted_data.values()):
            result[sheet_name] = extracted_data
            break  # 가장 먼저 찾은 시트에서 멈춤

    return result

if __name__ == "__main__":
    file_path = "/Users/zionchoi/Desktop/test_pdf/HHIENG25-036_20250612.xlsx"
    
    try:
        result = extract_from_excel(file_path)
        
        if not result:
            print(json.dumps({
                "shipper": [],
                "consignee": [],
                "invoice_no": "",
                "payment": "",
                "freight": "",
                "airport": "",
                "invoice_date": "",
                "arrival_date": ""
            }, ensure_ascii=False, indent=2))
        else:
            # 첫 번째 시트의 결과만 사용
            info = next(iter(result.values()))
            print(json.dumps(info, ensure_ascii=False, indent=2))
    
    except FileNotFoundError as e:
        print(f"오류: {e}")
    except Exception as e:
        print(f"예상치 못한 오류가 발생했습니다: {e}") 