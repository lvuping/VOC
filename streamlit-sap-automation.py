import streamlit as st
import win32com.client
import subprocess
import time
import os

def run_vbs_script(script_path):
    subprocess.call(['cscript', script_path])

def create_vbs_script(username, password, transaction_code):
    vbs_script = f"""
    If Not IsObject(application) Then
       Set SapGuiAuto  = GetObject("SAPGUI")
       Set application = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(connection) Then
       Set connection = application.Children(0)
    End If
    If Not IsObject(session) Then
       Set session    = connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session,     "on"
       WScript.ConnectObject application, "on"
    End If

    ' SAP 로그인
    session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "{username}"
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "{password}"
    session.findById("wnd[0]").sendVKey 0

    ' 트랜잭션 코드 실행
    session.findById("wnd[0]/tbar[0]/okcd").Text = "{transaction_code}"
    session.findById("wnd[0]").sendVKey 0

    ' 여기에 추가 SAP 작업을 구현할 수 있습니다

    ' 로그아웃
    session.findById("wnd[0]/tbar[0]/btn[15]").press()
    """

    with open("sap_automation.vbs", "w") as f:
        f.write(vbs_script)

def sap_automation(username, password, transaction_code):
    create_vbs_script(username, password, transaction_code)
    run_vbs_script("sap_automation.vbs")
    os.remove("sap_automation.vbs")  # 스크립트 실행 후 삭제
    return "SAP 자동화가 완료되었습니다."

st.title("SAP 자동화 인터페이스")

username = st.text_input("SAP 사용자 이름", key="username")
password = st.text_input("SAP 비밀번호", type="password", key="password")
transaction_code = st.text_input("실행할 트랜잭션 코드", key="transaction_code")

if st.button("SAP 작업 실행"):
    if username and password and transaction_code:
        result = sap_automation(username, password, transaction_code)
        st.success(result)
    else:
        st.error("모든 필드를 입력해주세요.")

st.sidebar.header("도움말")
st.sidebar.info("""
이 애플리케이션은 SAP 자동화 작업을 수행합니다.
1. SAP 사용자 이름과 비밀번호를 입력하세요.
2. 실행하려는 트랜잭션 코드를 입력하세요.
3. "SAP 작업 실행" 버튼을 클릭하여 자동화를 시작하세요.
""")
