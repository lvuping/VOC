
DATA: lt_source TYPE TABLE OF string.

CALL FUNCTION 'RFC_READ_REPORT'
  EXPORTING
    program  = 'ZPROGRAM'
  TABLES
    source   = lt_source
  EXCEPTIONS
    OTHERS   = 1.

IF sy-subrc = 0.
  LOOP AT lt_source INTO DATA(ls_line).
    WRITE: / ls_line.
  ENDLOOP.
ENDIF.
