
DATA: lt_entries TYPE TABLE OF rfc_db_fld.

CALL FUNCTION 'RFC_GET_TABLE_ENTRIES'
  EXPORTING
    table_name = 'TABLENAME'
  TABLES
    data       = lt_entries
  EXCEPTIONS
    OTHERS     = 1.

IF sy-subrc = 0.
  LOOP AT lt_entries INTO DATA(ls_entry).
    " ls_entry 를 통해 데이터 처리
  ENDLOOP.
ENDIF.
