# 폴더와 파일 다루기
'''
import os
import shutil
import csv
from openpyxl import load_workbook
'''

'''
# if not os.path.exists("C:/Sihwan/Book"):
#   print("폴더 없음")
# else:
#   print("폴더 있음")

# lists = os.listdir("C:/Sihwan/code/excel")
# print(lists)

# 폴더이름 변경
# if os.path.exists("C:/Sihwan/code/excel"):
#   os.rename("C:/Sihwan/code/excel", "C:/Sihwan/code/Newexcel") #원본폴더이름, 변경될 폴더 이름

#폴더 복사
# path_from = "C:/Sihwan/code/Newexcel"
# path_to = "C:/Newexcel"
# if not os.path.exists(path_to):
#   shutil.copytree(path_from, path_to)


#파일 내용 읽기.
# 파일 내용을 기록,수정,추가를 하면 .close()로 닫아야 한다.
file = open("example.txt", "r", encoding="utf-8")
# r:읽기모드  w:기록하기(저장) | a: 내용추가(수정) | x: 해당하는 파일이 없으면 만든다.(덮어씌우기)
content = file.read()
file.close()
print(content)

# with는 자동으로 .close()가 된다. (윗 코드랑 같은 의미)
with open("example.txt", "w", encoding="utf-8") as file:
  file.write("홍길동\n안녕하세요")


with open("example.txt", "r", encoding="utf-8") as file:
  line1 = file.readline()
  line2 = file.readline()

print(line1, line2)

# 공백 제거 -> 전부 출력
with open("example.txt", "r", encoding="utf-8") as file:
  line = file.readline()

  while line:
    print(line.strip()) # .strip() : 공백을 전부 제거
    line = file.readline()

# readlines의 s 같은건 가급적 사용 금지!!
with open("example.txt", "r", encoding="utf-8") as file:
  line = file.readlines()
  print(line)

with open("example.txt", "r", encoding="utf-8") as file:
  line = file.readline()
  print(line)

with open("example.csv", "w", encoding="cp949", newline="") as file:
  csv_writer = csv.writer(file)
  csv_writer.writerow(["이름","나이","직업"])
  csv_writer.writerow(["홍길동","29","취준생"])
  csv_writer.writerow(["박시환","30","직장인"])
  csv_writer.writerow(["희야","34","직장인"])
  csv_writer.writerow(["날좀","25","프리"])
  csv_writer.writerow(["바라봐","30","직장인"])
'''

'''엑셀 연동
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "수강생 정보"

# ws["A1"] = "이철수"
# wb.save("수강생 리스트.xlsx")
# wb.close()

column = ["번호", "이름", "과목"]
ws.append(column)
row = [[1,"이철수","수학"],[2, "빛나리", "영어"],[1,"홍길동","수학"]]
for data in row:
  ws.append(data)
# row = [1, "이철수", "수학"]
# ws.append(row)

# 시트 자동 생성
# wb.create_sheet("중간 평가")
# wb.create_sheet("기말 평가")
wb.save("수강생_리스트.xlsx")
wb.close()

'''

'''
wb = load_workbook(filename="월별구매고객리스트.xlsx")
ws = wb["10월"]

new_rows = list(ws.rows)[2:]

for row in new_rows:
  row_values = [cell.value for cell in row]
  print(row_values)
'''

