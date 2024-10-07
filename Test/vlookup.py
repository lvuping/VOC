import pandas as pd

# Excel 파일 읽기
df1 = pd.read_excel('file1.xlsx')
df2 = pd.read_excel('file2.xlsx')

# print("file1.xlsx columns:", df1.columns)
# print("file2.xlsx columns:", df2.columns)

# df1에 B 열 추가 (이미 존재하지 않는 경우에만)
if '동물' not in df1.columns:
    df1['동물'] = ''

# df1을 기준으로 df2의 '동물' 열 데이터를 가져와 매칭
df1['동물'] = df1['Phone Numbers'].map(df2.set_index('Phone Numbers')['동물'])

# '과일' 열이 존재하는 경우에만 C 열에 매칭
if '과일' in df2.columns:
    if '과일' not in df1.columns:
        df1['과일'] = ''
    df1['과일'] = df1['Phone Numbers'].map(df2.set_index('Phone Numbers')['과일'])

# 결과를 새로운 Excel 파일로 저장
df1.to_excel('result.xlsx', index=False)

# 결과 출력
print("\nResult:")
print(df1)
