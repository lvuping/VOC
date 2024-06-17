
DATA: lt_bdcdata TYPE TABLE OF bdcdata,
      lt_messages TYPE TABLE OF bdcmsgcoll.

" BDC data filling example
PERFORM bdc_dynpro USING 'SAPL...'
                          '0100'.
PERFORM bdc_field  USING 'BDC_OKCODE'
                          '/00'.
PERFORM bdc_field  USING 'RFBELN'
                          '123456'.

CALL FUNCTION 'RFC_CALL_TRANSACTION_USING'
  EXPORTING
    tcode       = 'VA01'
    mode        = 'E'
    update      = 'S'
  TABLES
    bdctab      = lt_bdcdata
    messages    = lt_messages
  EXCEPTIONS
    OTHERS      = 1.

IF sy-subrc = 0.
  " 성공 처리
ELSE.
  " 오류 처리
ENDIF.
