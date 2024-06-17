DATA: lt_bdcdata TYPE TABLE OF bdcdata,
      lt_msgs TYPE TABLE OF bapiret2,
      lv_tran TYPE c LENGTH 20 VALUE 'Your_Transaction_Code'.

" Populate lt_bdcdata with your BDC recording data

CALL FUNCTION 'RFC_CALL_TRANSACTION_USING'
  EXPORTING
    tcode              = lv_tran
    mode               = 'N'
  TABLES
    bdctab             = lt_bdcdata
    messages           = lt_msgs
  EXCEPTIONS
    communication_failure = 1
    system_failure        = 2
    OTHERS               = 3.

IF sy-subrc <> 0.
  " Handle errors
ENDIF.

