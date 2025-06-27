import requests
import uuid
import time
import json
from pdf2image import convert_from_path
import os
import shutil

# Clova OCR API 설정
api_url = 'https://8t3q98q5p4.apigw.ntruss.com/custom/v1/43241/4332772734bad9042b8d3b16ced05e86995eb0deddc51a2b60bd558c497bcc97/general'
secret_key = 'bWJMSWtWSm9SWmdIa2Z5UkFock5JTWp6S1Bpdm1VYkE='

# PDF 파일 경로
pdf_file = '/Users/zionchoi/Desktop/test_pdf/EU2506-0217 SIMMTECH 6-13申告.pdf'

total_start = time.time()

def convert_pdf_to_images(pdf_path, output_dir='temp_images'):
    """PDF를 이미지로 변환"""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # PDF를 이미지로 변환
    images = convert_from_path(pdf_path)
    image_paths = []
    
    for i, image in enumerate(images):
        image_path = os.path.join(output_dir, f'page_{i+1}.jpg')
        image.save(image_path, 'JPEG')
        image_paths.append(image_path)
    
    return image_paths

# PDF를 이미지로 변환
print("PDF를 이미지로 변환 중...")
pdf2img_start = time.time()
image_paths = convert_pdf_to_images(pdf_file)
pdf2img_end = time.time()
print(f"PDF 변환 소요 시간: {pdf2img_end - pdf2img_start:.2f}초")

all_results = []
for i, image_path in enumerate(image_paths):
    print(f"페이지 {i+1} OCR 처리 중...")
    ocr_start = time.time()
    request_json = {
        'images': [
            {
                'format': 'jpg',
                'name': 'demo'
            }
        ],
        'requestId': str(uuid.uuid4()),
        'version': 'V2',
        'timestamp': int(round(time.time() * 1000)),
        'lang': 'ko, ja'
    }
    payload = {'message': json.dumps(request_json).encode('UTF-8')}
    files = [
        ('file', open(image_path, 'rb'))
    ]
    headers = {
        'X-OCR-SECRET': secret_key
    }
    response = requests.request("POST", api_url, headers=headers, data=payload, files=files)
    print(response.text)
    all_results.append(response.json())
    ocr_end = time.time()
    print(f"페이지 {i+1} OCR 소요 시간: {ocr_end - ocr_start:.2f}초")

# 결과를 파일로 저장 (원본 JSON)
with open('clova_result.json', 'w', encoding='utf-8') as f:
    json.dump(all_results, f, ensure_ascii=False, indent=2)

# 임시 이미지 파일 정리
temp_dir = 'temp_images'
if os.path.exists(temp_dir):
    shutil.rmtree(temp_dir)

total_end = time.time()
print(f"전체 소요 시간: {total_end - total_start:.2f}초")
print("OCR 완료! 결과가 clova_result.json에 저장되었습니다.")