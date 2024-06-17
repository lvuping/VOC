
DATA: lt_data TYPE TABLE OF RFC_DB_FLD,
      lt_fields TYPE TABLE OF RFC_DB_FLD.

CALL FUNCTION 'RFC_READ_TABLE'
  EXPORTING
    query_table = 'TABLENAME'
    delimiter   = '|'
  TABLES
    fields      = lt_fields
    data        = lt_data
  EXCEPTIONS
    OTHERS      = 1.

IF sy-subrc = 0.
  LOOP AT lt_data INTO DATA(ls_data).
    " ls_data 를 통해 데이터 처리
  ENDLOOP.
ENDIF.
