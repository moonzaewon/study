import openpyxl

status_counter = {}

with open('access.log.2017-10-13', 'r') as fp:
    for line in fp:
        stat = line.split()[8]
        if stat in status_counter:
            status_counter[stat] += 1
        else:
            status_counter[stat] = 1
            
with open('statics.txt', 'w')as fp:
    for key, value in status_counter.items():
        print(f"{key} : {value}", file=fp)

print("통계를 출력하였습니다.")

wb = openpyxl.Workbook()    # wb는 엑셀파일 데이터
ws = wb.active              # 엑셀파일에는 기본시트를 얻어내기
for item in status_counter.items():
    ws.append(item)         # 시트의 행으로 기록한다. (리스트, 튜플 데이터)
count = len(status_counter)
ws.append(['요청합계', f'=SUM(B1:B{count}'])

wb.save('statics.xlsx')     # wb 데이터를 파일로 저장
wb.close()                  # 엑셀파일 닫기
print('엑셀파일로 통계를 출력하였습니다.')