import requests
import json
import time
 
api_key = "up_JRsVNVkblzi60ZHnaFajLOo7nm5pG"
filename = "/Users/zionchoi/Desktop/test_pdf/EU2506-0217 SIMMTECH 6-13申告.pdf"
 
url = "https://api.upstage.ai/v1/document-digitization"
headers = {"Authorization": f"Bearer {api_key}"}

total_start = time.time()

file_open_start = time.time()
files = {"document": open(filename, "rb")}
file_open_end = time.time()
print(f"파일 열기 소요 시간: {file_open_end - file_open_start:.2f}초")

data = {"model": "ocr"}
ocr_start = time.time()
response = requests.post(url, headers=headers, files=files, data=data)
ocr_end = time.time()
print(f"OCR API 요청 소요 시간: {ocr_end - ocr_start:.2f}초")
 
print(response.json())

# Save full JSON
with open("result.json", "w", encoding="utf-8") as f:
    json.dump(response.json(), f, ensure_ascii=False, indent=2)

# Save extracted text only
result = response.json()
text = result.get("text")
if text:
    with open("result.txt", "w", encoding="utf-8") as f:
        f.write(text)

total_end = time.time()
print(f"전체 소요 시간: {total_end - total_start:.2f}초")