import streamlit as st
import pandas as pd
import pyrfc
import json
import io
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# 자격 증명 로드
with open("C:\\Users\\lvupi\\workplace\\py\\main.json") as f:
    credentials = json.load(f)

st.title("SAP Automation")

# 사용자로부터 테이블 이름 입력 받기
table_name = st.text_input("조회할 테이블 이름을 입력하세요:", "ZTSVC109")

x = st.slider("검색 Rows", 10, 5000, 10)
st.write(x, "개 조회")

# 필드 이름 저장할 공간 초기화
if "fields" not in st.session_state:
    st.session_state.fields = []
if "selected_fields" not in st.session_state:
    st.session_state.selected_fields = []
if "conditions" not in st.session_state:
    st.session_state.conditions = {}

# Load Fields 버튼 클릭 시 필드 로드
if st.button("Load Fields"):
    try:
        with pyrfc.Connection(**credentials) as conn:
            # 테이블 구조 가져오기
            table_fields = conn.call(
                "DDIF_FIELDINFO_GET", TABNAME=table_name, LANGU="E"
            )

            # 필드 이름 추출
            st.session_state.fields = [
                field["FIELDNAME"] for field in table_fields["DFIES_TAB"]
            ]
    except pyrfc.RFCLibError as e:
        st.error(f"RFC Library Error: {e}")
    except Exception as e:
        st.error(f"An error occurred: {e}")

# 필드 선택 UI
if st.session_state.fields:
    st.session_state.selected_fields = st.multiselect(
        "출력할 필드를 선택하세요:",
        st.session_state.fields,
        default=st.session_state.fields[:5],
    )

    # 조건 입력 UI
    st.subheader("필드에 조건 추가")
    for field in st.session_state.selected_fields:
        condition = st.text_input(f"조건 ({field}):", key=f"condition_{field}")
        st.session_state.conditions[field] = condition

# Search Table 버튼 클릭 시 테이블 검색
if st.button("Search Table"):
    if st.session_state.selected_fields:
        try:
            with pyrfc.Connection(**credentials) as conn:
                # 조건을 위한 옵션 생성
                options = []
                conditions = []

                # Constructing the conditions
                for field, condition in st.session_state.conditions.items():
                    if condition:
                        conditions.append(f"{field} EQ '{condition}'")

                # Join conditions with 'AND' if there are multiple
                if conditions:
                    condition_string = " AND ".join(conditions)
                    options.append({"TEXT": condition_string})

                # 테이블 데이터 조회
                result = conn.call(
                    "RFC_READ_TABLE",
                    QUERY_TABLE=table_name,
                    ROWCOUNT=x,
                    FIELDS=[
                        {"FIELDNAME": field}
                        for field in st.session_state.selected_fields
                    ],
                    OPTIONS=options,
                    DELIMITER=",",
                )

                # 데이터 처리
                data = []
                delimiter = ","
                for line in result["DATA"]:
                    raw_data = line["WA"].strip().split(delimiter)
                    data.append(raw_data)

                # DataFrame 생성
                df = pd.DataFrame(data, columns=st.session_state.selected_fields)

                # 데이터 출력
                st.subheader(f"{table_name} 테이블 데이터")
                st.dataframe(df)

                # ERDAT 필드가 있는 경우 날짜 분포 그래프 생성
                if "ERDAT" in df.columns:
                    st.subheader("ERDAT Date Distribution Graph.")

                    # 날짜 형식 변환
                    df["ERDAT"] = pd.to_datetime(
                        df["ERDAT"], format="%Y%m%d", errors="coerce"
                    )

                    # 날짜별 빈도 계산
                    date_counts = df["ERDAT"].value_counts().sort_index()

                    # 시간축 그래프 그리기
                    plt.style.use("ggplot")  # 스타일 테마 적용
                    fig, ax = plt.subplots(figsize=(10, 6))

                    # 그래프 그리기
                    ax.plot(
                        date_counts.index,
                        date_counts.values,
                        marker="o",
                        linestyle="-",
                        color="b",
                    )
                    ax.set_title("ERDAT 별 데이터 수", fontsize=16)
                    ax.set_xlabel("날짜", fontsize=14)
                    ax.set_ylabel("빈도수", fontsize=14)
                    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m-%d"))
                    ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
                    ax.grid(
                        visible=True,
                        which="both",
                        axis="both",
                        linestyle="--",
                        linewidth=0.5,
                    )
                    plt.xticks(rotation=45)

                    st.pyplot(fig)

                # Excel 파일로 변환
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    df.to_excel(writer, sheet_name="Data", index=False)

                # 다운로드 버튼 생성
                st.download_button(
                    label="Download Excel file",
                    data=output.getvalue(),
                    file_name=f"{table_name}_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        except pyrfc.RFCLibError as e:
            st.error(f"RFC Library Error: {e}")
        except Exception as e:
            st.error(f"An error occurred: {e}")
    else:
        st.error("Please select at least one field to display.")
