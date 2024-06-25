import pyrfc
from pyrfc import Connection

# SAP 연결 설정
conn = Connection(
    user='SAP_USER',
    passwd='SAP_PASSWORD',
    ashost='SAP_SERVER',
    sysnr='SYSNR',
    client='CLIENT',
    lang='EN'
)

# BDC 데이터 구성
lt_bdcdata = [
    {'PROGRAM': 'SAPL...', 'DYNPRO': '0100', 'DYNBEGIN': 'X'},
    {'FNAM': 'BDC_OKCODE', 'FVAL': '/00'},
    {'FNAM': 'RFBELN', 'FVAL': '123456'}
]

# RFC 호출
try:
    result = conn.call('RFC_CALL_TRANSACTION_USING', 
                       TCODE='VA01', 
                       MODE='E', 
                       UPDATE='S', 
                       BDCTAB=lt_bdcdata)
    
    if 'MESSAGES' in result:
        lt_messages = result['MESSAGES']
        for message in lt_messages:
            print(message)
    else:
        print("Transaction successful")

except pyrfc.CommunicationError as e:
    print(f"Communication error: {e}")
except pyrfc.LogonError as e:
    print(f"Logon error: {e}")
except Exception as e:
    print(f"An error occurred: {e}")

