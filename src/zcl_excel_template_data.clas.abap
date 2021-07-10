CLASS zcl_excel_template_data DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    DATA mt_data TYPE zexcel_t_template_data .

    METHODS add
      IMPORTING
        !iv_sheet TYPE zexcel_sheet_title
        !iv_data  TYPE data .
  PROTECTED SECTION.
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_excel_template_data IMPLEMENTATION.


  METHOD add.
    FIELD-SYMBOLS: <fs_data> TYPE zexcel_s_template_data,
                   <fs_any> TYPE any.

    APPEND INITIAL LINE TO mt_data ASSIGNING <fs_data>.
    <fs_data>-sheet = iv_sheet.
    CREATE DATA  <fs_data>-data LIKE iv_data.

    ASSIGN <fs_data>-data->* TO <fs_any>.
    <fs_any> = iv_data.

  ENDMETHOD.
ENDCLASS.
