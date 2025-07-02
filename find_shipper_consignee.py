import pandas as pd
import os
import json

def extract_single_value(file_path):
    """
    주어진 엑셀 파일에서 Shipper와 Consignee 정보를 추출합니다.
    :param file_path: 엑셀 파일 경로
    :return: dict 형태로 {'shipper': [...], 'consignee': [...]} 반환
    """
    # 파일 존재 여부 확인
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
    
    # ExcelFile 객체로 읽기 (시트 이름에 접근하기 위해)
    xls = pd.ExcelFile(file_path)
    result = {}

    for sheet_name in xls.sheet_names:
        # 모든 셀을 문자열로 처리
        df = xls.parse(sheet_name).astype(str) 
        shipper, consignee = extract_shipper_consignee(df)
        if shipper or consignee:
            result[sheet_name] = {'shipper': shipper, 'consignee': consignee}
            break # 가장 먼저 찾은 시트에서 멈춤

    return result

def extract_shipper_consignee(df):
    """
    엑셀 데이터에서 Shipper와 Consignee 정보를 추출합니다.
    """
    # 1. 텍스트로 찾기
    def find_index(keywords):
        for idx, row in df.iterrows():
            joined = " ".join(row.fillna("").tolist()).lower()
            for keyword in keywords:
                if keyword in joined:
                    return idx
        return None

    # 1-1. 우측 + 우측 포함 하단에 데이터가 있는 경우
    def get_next_lines(start_idx):
        lines = []
        
        # 1. 같은 행에서 키워드 오른쪽 셀들의 정보 추출
        if start_idx is not None:
            row_cells = [cell if cell.lower() != "nan" else "" for cell in df.iloc[start_idx].fillna("").tolist()]
            # 키워드가 있는 셀의 인덱스를 찾기
            keyword_col_idx = None
            for col_idx, cell in enumerate(row_cells):
                if "shipper" in cell.lower() or "consignee" in cell.lower():
                    keyword_col_idx = col_idx
                    break
            
            # 키워드 오른쪽 셀들의 정보 추출 (최대 3칸까지)
            if keyword_col_idx is not None:
                for col_offset in range(1, 4):  # 오른쪽 1~3칸
                    if keyword_col_idx + col_offset < len(row_cells):
                        cell_value = row_cells[keyword_col_idx + col_offset].strip()
                        if cell_value and not any(k in cell_value.lower() for k in ["shipper", "consignee"]):
                            lines.append(cell_value)
        
        for offset in range(1, 5):
            if start_idx + offset < len(df):
                # 각 셀의 'nan'을 공백으로 변환
                row_cells = [cell if cell.lower() != "nan" else "" for cell in df.iloc[start_idx + offset].fillna("").tolist()]
                line = " ".join(row_cells).strip()
                # 완전히 빈 줄은 제외
                if line and not any(k in line.lower() for k in ["shipper", "consignee"]):
                    lines.append(line)
                elif not line:
                    break
        return lines

    # 1-2. 하단 + 우측에 데이터가 있는 경우

    # 헤더 키워드 설정
    shipper_header_keywords = ["shipper", "shipper/exporter"]
    consignee_header_keywords = ["consignee"]

    # 헤더 찾기
    shipper_idx = find_index(shipper_header_keywords)
    consignee_idx = find_index(consignee_header_keywords)

    shipper_info = get_next_lines(shipper_idx) if shipper_idx is not None else []
    consignee_info = get_next_lines(consignee_idx) if consignee_idx is not None else []

    return shipper_info, consignee_info

if __name__ == "__main__":
    file_path = "/Users/zionchoi/Desktop/test_pdf/HHIENG25-036_20250612.xlsx"
    result = extract_single_value(file_path)

    if not result:
        print(json.dumps({"Shipper": [], "Consignee": []}, ensure_ascii=False, indent=2))
    else:
        info = next(iter(result.values()))
        print(json.dumps({
            "Shipper": info['shipper'],
            "Consignee": info['consignee']
        }, ensure_ascii=False, indent=2))