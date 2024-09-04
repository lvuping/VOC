
const { Client } = require('node-rfc');

// SAP 연결 설정
const abapSystem = {
    user: 'SAP_USER',
    passwd: 'SAP_PASSWORD',
    ashost: 'SAP_SERVER',
    sysnr: 'SYSNR',
    client: 'CLIENT',
    lang: 'EN'
};

const client = new Client(abapSystem);

async function callTransactionUsing() {
    try {
        // SAP 시스템에 연결
        await client.open();

        // BDC 데이터 구성
        const lt_bdcdata = [
            { PROGRAM: 'SAPL...', DYNPRO: '0100', DYNBEGIN: 'X' },
            { FNAM: 'BDC_OKCODE', FVAL: '/00' },
            { FNAM: 'RFBELN', FVAL: '123456' }
        ];

        // RFC 호출
        const result = await client.call('RFC_CALL_TRANSACTION_USING', {
            TCODE: 'VA01',
            MODE: 'E',
            UPDATE: 'S',
            BDCTAB: lt_bdcdata
        });

        if (result.MESSAGES) {
            const lt_messages = result.MESSAGES;
            lt_messages.forEach(message => {
                console.log(message);
            });
        } else {
            console.log('Transaction successful');
        }
    } catch (err) {
        console.error(`Error: ${err.message}`);
    } finally {
        // SAP 연결 종료
        await client.close();
    }
}

callTransactionUsing();
