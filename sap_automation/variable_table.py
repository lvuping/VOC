import streamlit as st
import pandas as pd
import pyrfc
import json

# 자격 증명 로드
with open("C:\\Users\\lvupi\\workplace\\py\\main.json") as f:
    credentials = json.load(f)

st.title('SAP Automation')

# 사용자로부터 테이블 이름 입력 받기
table_name = st.text_input('조회할 테이블 이름을 입력하세요:', 'ZTBP007')

x = st.slider('검색 Rows', 10)
st.write(x, '개 조회')

if st.button("Search Table"):
    with pyrfc.Connection(**credentials) as conn:
        # 테이블 구조 가져오기
        table_fields = conn.call('DDIF_FIELDINFO_GET',
                                 TABNAME=table_name,
                                 LANGU='E')  # 'E'는 영어를 의미합니다.

        # 필드 이름 추출
        fields = [field['FIELDNAME'] for field in table_fields['DFIES_TAB']]

        # 테이블 데이터 조회
        result = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE=table_name,
            ROWCOUNT=x,
            FIELDS=fields,
            DELIMITER=","
        )

        # 데이터 처리
        data = []
        delimiter = ","
        for line in result["DATA"]:
            raw_data = line["WA"].strip().split(delimiter)
            data.append(raw_data)

        # DataFrame 생성
        df = pd.DataFrame(data, columns=fields)

        # 데이터 출력
        st.subheader(f'{table_name} 테이블 데이터')
        st.dataframe(df)

        # 테이블 구조 출력
        st.subheader(f'{table_name} 테이블 구조')
        structure_df = pd.DataFrame(table_fields['DFIES_TAB'])
        st.dataframe(structure_df[['FIELDNAME', 'DATATYPE', 'LENG', 'OUTPUTLEN', 'DECIMALS']])