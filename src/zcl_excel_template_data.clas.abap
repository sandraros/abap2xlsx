CLASS zcl_excel_template_data DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES: sheet_title TYPE c LENGTH 31,
           sheet_titles TYPE STANDARD TABLE OF sheet_title WITH DEFAULT KEY,
           BEGIN OF template_data_sheet,
             sheet TYPE sheet_title,
             data  TYPE REF TO data,
           END OF template_data_sheet,
           template_data_sheets TYPE STANDARD TABLE OF template_data_sheet WITH DEFAULT KEY.

    DATA mt_data TYPE template_data_sheets .

    METHODS add
      IMPORTING
        iv_sheet TYPE sheet_title
        iv_data  TYPE data .
  PROTECTED SECTION.
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_excel_template_data IMPLEMENTATION.


  METHOD add.
    FIELD-SYMBOLS: <fs_data> TYPE template_data_sheet,
                   <fs_any>  TYPE any.

    APPEND INITIAL LINE TO mt_data ASSIGNING <fs_data>.
    <fs_data>-sheet = iv_sheet.
    CREATE DATA  <fs_data>-data LIKE iv_data.

    ASSIGN <fs_data>-data->* TO <fs_any>.
    <fs_any> = iv_data.

  ENDMETHOD.
ENDCLASS.
