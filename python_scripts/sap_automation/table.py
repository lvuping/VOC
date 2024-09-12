import streamlit as st
import pandas as pd
import pyrfc
import json

f = open("C:\\Users\\lvupi\\workplace\\py\\main.json")
credentials = json.load(f)

st.title('SAP Automation')
st.write('테이블 ZTBP007 조회')

x = st.slider('검색 Rows', 10)  # 👈 this is a widget
st.write(x, '개 조회')

st.subheader('Raw data')

if st.button("Search Table"):
    with pyrfc.Connection(**credentials) as conn:
        fields = ["COMPANY",
                  "COUNTRY",
                  "ASC_CODE",
                  "POST_CODE",
                  "REGION",            
                  "CITY", ]

        result = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="ZTBP007",
            ROWCOUNT=x,
            FIELDS=fields,
            DELIMITER=",",
            OPTIONS=[{"TEXT": "COMPANY = 'C310'"}]
        )

    data = []
    delimiter = ","
    for line in result["DATA"]:
        raw_data = line["WA"].strip().split(delimiter)
        data.append(raw_data)

    df = pd.DataFrame(data, columns=fields)  # 여기서 columns 매개변수를 사용합니다.
    st.dataframe(df)  # st.dataframe()을 사용하여 더 인터랙티브한 테이블을 표시합니다.