import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 엑셀파일 불러오기
def load_excel():
    global student_file, grade_assignment, grade_final, grade_middle, grade_attendance
    # 학번 오름차순으로 정렬해서 저장
    student_file = pd.read_excel('Entity/학생.xlsx').sort_values('학번')
    grade_assignment = pd.read_excel('Entity/성적_과제.xlsx').sort_values('학번')
    grade_final = pd.read_excel('Entity/성적_기말.xlsx').sort_values('학번')
    grade_middle = pd.read_excel('Entity/성적_중간.xlsx').sort_values('학번')
    grade_attendance = pd.read_excel('Entity/성적_출석.xlsx').sort_values('학번')

# 엑셀 인스턴스 개수확인
def excel_row_count():
    # 엑셀 파일에 오타가 없다고 가정
    # 행의 수가 다른 엑셀 파일이 있으면 종료
    if len(student_file) != len(grade_assignment) or len(student_file) != len(grade_final) or len(student_file) != len(grade_middle) or len(student_file) != len(grade_attendance):
        print("파일의 학생 수가 일치하지 않습니다.")
        exit()
    return len(student_file)

# 평점 계산
def get_rate(attendance, grade):
    if attendance < 60: 
        rate = 'F'
        grade = str(grade) + "(출석미달)"
    else:
        if grade >= 90:
            rate = 'A'
        elif grade >= 80:
            rate = 'B'
        elif grade >= 70:
            rate = 'C'
        elif grade >= 60:
            rate = 'D'
        else:
            rate = 'F'
    return grade, rate

rating = [] # 성적을 저장할 리스트
load_excel()
student_rows = excel_row_count()

# 한줄씩 불러와서 성적 계산 후 저장(학번 오름차순)
for index in range(student_rows):
    # 이름, 학번 저장
    name = student_file.loc[index, '이름']
    num = student_file.loc[index, '학번']
    # 성적 저장(값이 null이면 0점으로 처리)
    middle = int(grade_middle.loc[index, '성적']) if not pd.isna(grade_middle.loc[index, '성적']) else 0 # 중간 점수
    final = int(grade_final.loc[index, '성적']) if not pd.isna(grade_final.loc[index, '성적']) else 0 # 기말 점수
    attendance = int(grade_attendance.loc[index, '성적']) if not pd.isna(grade_attendance.loc[index, '성적']) else 0 # 출석 점수
    asSum = int(grade_assignment.loc[index, '과제1']) if not pd.isna(grade_assignment.loc[index, '과제1']) else 0 # 과제 점수 4개 합계산
    asSum += int(grade_assignment.loc[index, '과제2']) if not pd.isna(grade_assignment.loc[index, '과제2']) else 0
    asSum += int(grade_assignment.loc[index, '과제3']) if not pd.isna(grade_assignment.loc[index, '과제3']) else 0
    asSum += int(grade_assignment.loc[index, '과제4']) if not pd.isna(grade_assignment.loc[index, '과제4']) else 0
    # 총점 계산
    grade = round(middle * 0.3 + final * 0.3 + attendance * 0.2 + asSum * 0.2, 2)
    grade, rate = get_rate(attendance, grade)
    rating.append([name, num, grade, rate])

# 0번 인덱스(이름)기준 오름차순 정렬
rating.sort(key=lambda x:x[0])
# 데이터프레임으로 변환 후 엑셀파일로 저장하고 출력
df = pd.DataFrame(rating, columns=['Name', 'Num', 'Grade', 'Rate'])
try:
    df.to_excel("Entity/최종 성적.xlsx")
except Exception as e:
    print(f"파일 저장 실패: {e}")
print(df)
# A부터 오름차순으로 차트 출력
sns.countplot(x='Rate',data=df, order=['A', 'B', 'C', 'D', 'F'])
plt.show()




    

    

