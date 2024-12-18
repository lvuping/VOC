import streamlit as st
import pandas as pd


def main():
    st.title("Excel/CSV 파일 업로더")

    # 파일 업로드 위젯
    uploaded_file = st.file_uploader(
        "Excel 또는 CSV 파일을 업로드해주세요", type=["csv", "xls", "xlsx"]
    )

    if uploaded_file is not None:
        try:
            # 파일 확장자 확인
            file_extension = uploaded_file.name.split(".")[-1]

            # 파일 형식에 따라 데이터 읽기
            if file_extension == "csv":
                df = pd.read_csv(uploaded_file)
            elif file_extension == "xls":
                df = pd.read_excel(uploaded_file, engine="xlrd")
            elif file_extension == "xlsx":
                df = pd.read_excel(uploaded_file, engine="openpyxl")

            # 데이터 정보 표시
            st.subheader("데이터 미리보기")
            st.dataframe(df)

            # 데이터 기본 정보 표시
            st.subheader("데이터 기본 정보")
            st.write(f"행 수: {df.shape[0]}")
            st.write(f"열 수: {df.shape[1]}")
            st.write("컬럼명:", df.columns.tolist())

        except Exception as e:
            st.error(f"파일을 읽는 중 오류가 발생했습니다: {str(e)}")


if __name__ == "__main__":
    main()
