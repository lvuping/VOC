const noderfc = require("node-rfc");

// SAP 연결 설정
const client = new noderfc.Client({
  user: "username",
  passwd: "password",
  ashost: "hostname",
  sysnr: "00",
  client: "100",
  lang: "EN",
});

// 연결 및 RFC 호출
client.connect(function (err) {
  if (err) {
    return console.error("Error connecting to SAP:", err);
  }

  // STFC_CONNECTION 함수 호출
  client.invoke(
    "STFC_CONNECTION",
    { REQUTEXT: "Hello SAP" },
    function (err, res) {
      if (err) {
        return console.error("Error invoking RFC:", err);
      }

      // 응답 출력
      console.log("Response from SAP:");
      console.log("ECHOTEXT:", res.ECHOTEXT);
      console.log("RESPTEXT:", res.RESPTEXT);
    }
  );
});
