import openpyxl
import os
import locale
from os import listdir
from datetime import datetime, timedelta

print('*'*50)
print("출결관리 프로그램v2.0 \n장애발생 시 삼양초 '이은섭'으로 메세지 주세요. :)")
print('*'*50)

BASE_DIR = os.path.dirname(os.path.realpath(__file__))
LIST_DIR = os.path.join(BASE_DIR, "roll_list/")
locale.setlocale(locale.LC_ALL, '')

list_files = listdir("roll_list")

days = []

for f in list_files:
    # 파일 읽기
    # 실행 중 생성되는 임시파일 무시
    if not '.xlsx' in f:
        continue

    thisDay = f.split("-")[-1].rstrip('.xlsx')
    thisDay = datetime.strptime(thisDay, "%m월%d일")
    thisDay = thisDay.replace(year=2020)
    days.append(thisDay.strftime("%m%d"))
    thisDay = thisDay.strftime("%m월%d일(%a)")

    print('[처리중인 파일]')
    print(f)
    read_f = os.path.join(LIST_DIR, f)
    wb = openpyxl.load_workbook(read_f, read_only=True)

    sheets = wb.sheetnames
    sheet = wb[sheets[0]]

    # 학생 수, 과목 수
    count_students = sheet.max_row - 1
    print('학생 수:', count_students)
    count_subjects = sheet.max_column - 1
    print('과목 수:', count_subjects)

    # 학생 명렬표
    student_names = []
    for i in range(count_students):
        student_names.append(sheet.cell(row=i + 2, column=1).value)

    # 과목 리스트
    subjects = []
    for i in range(count_subjects):
        subject = sheet.cell(row=1, column=i + 2).value
        subject = subject.split("-")[1]
        subjects.append(subject)

    print(student_names)
    print(subjects)

    # 데이터 가공
    data = [['이름', '과목', '이수 시간']]
    for student in range(2, count_students + 2):
        for subject in range(2, count_subjects + 2):
            s_name = sheet.cell(row=student, column=1).value
            subj = subjects[subject - 2]
            # 이수 시간은 datetime 으로 저장
            checked_time = sheet.cell(
                row=student, column=subject).value
            if checked_time == None:
                checked_time = '결석'
            else:
                checked_time = checked_time.rstrip()
                checked_time = datetime.strptime(checked_time, "%m/%d %H:%M")
                # zero leading datetime
                checked_time = checked_time.strftime("%m/%d %H:%M")
                # print(checked_time)
            data.append([s_name, subj, checked_time])

    # 이수 시각 기준으로 재정렬
    reordered = [['이름', '과목', '이수시간', '소요시간']]
    temp = []
    for i in range(len(data)):
        if i == 0:

            continue
        if not i % count_subjects == 0:
            temp.append(data[i])
        else:
            temp.append(data[i])
            # 이수 시간 기준으로 정렬

            temp.sort(key=lambda temp: temp[2])
            # 소요시간 계산을 위한 temp2
            temp2 = [temp[0]] + temp[:-1]

            for i in range(len(temp)):
                if temp[i][2] != '결석':
                    elapsed = datetime.strptime(
                        temp[i][2], "%m/%d %H:%M") - datetime.strptime(temp2[i][2], "%m/%d %H:%M")

                    if elapsed < timedelta(seconds=1):
                        temp[i].append('시작')
                        continue
                    elif elapsed < timedelta(hours=24):
                        dt = datetime(2020, 1, 1, 0, 0, 0) + elapsed
                        elapsed = f'({dt.time()})'
                    else:
                        elapsed = '(다른 날 수강)'
                else:
                    elapsed = '-'
                temp[i].append(elapsed)
            for item in temp:
                reordered.append(item)
            temp = []

    result_wb = openpyxl.Workbook()
    result_ws = result_wb.active
    result_ws.title = thisDay
    for row in reordered:

        result_ws.append(row)

# 엑셀 파일에 쓰기
school_infos = f.split('-')[:-1]
file_prefix = '_'.join(school_infos)
file_period = f'{days[0]}~{days[-1]}'
RESULT_DIR = os.path.join(BASE_DIR, "results")
result_filename = os.path.join(
    RESULT_DIR, f'(출결){file_prefix}_{file_period}.xlsx')
result_wb.save(result_filename)
