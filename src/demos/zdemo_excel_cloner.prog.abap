*&---------------------------------------------------------------------*
*& Report zdemo_excel_cloner
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zdemo_excel_cloner.

CONSTANTS gc_save_file_name TYPE string VALUE ''.
INCLUDE zdemo_excel_outputopt_incl.

CLASS lcx_app DEFINITION INHERITING FROM cx_static_check.
  PUBLIC SECTION.
    METHODS constructor
      IMPORTING
        name      TYPE string
        !textid   LIKE textid OPTIONAL
        !previous LIKE previous OPTIONAL.
    METHODS get_text REDEFINITION.
    METHODS get_longtext REDEFINITION.
    DATA name TYPE string.
ENDCLASS.

CLASS lcx_app IMPLEMENTATION.
  METHOD constructor.
    super->constructor(
      EXPORTING
        textid   = textid
        previous = previous ).
    me->name = name.
  ENDMETHOD.
  METHOD get_text.
    result = |Error { name }|.
  ENDMETHOD.
  METHOD get_longtext.
    result = get_text( ).
  ENDMETHOD.
ENDCLASS.

CLASS lcl_excel_cloner DEFINITION.
  PUBLIC SECTION.

    METHODS constructor.

    METHODS clone_workbook
      IMPORTING
        io_input_excel TYPE REF TO zcl_excel
      RETURNING
        VALUE(result)  TYPE REF TO zcl_excel
      RAISING
        zcx_excel.

  PRIVATE SECTION.

    TYPES: BEGIN OF ty_mapping_of_style,
             input  TYPE REF TO zcl_excel_style,
             output TYPE REF TO zcl_excel_style,
           END OF ty_mapping_of_style,
           ty_mapping_of_styles TYPE HASHED TABLE OF ty_mapping_of_style WITH UNIQUE KEY input,
           BEGIN OF ty_mapping_of_worksheet,
             input  TYPE REF TO zcl_excel_worksheet,
             output TYPE REF TO zcl_excel_worksheet,
           END OF ty_mapping_of_worksheet,
           ty_mapping_of_worksheets TYPE HASHED TABLE OF ty_mapping_of_worksheet WITH UNIQUE KEY input,
           BEGIN OF ty_mapping_of_drawing,
             input  TYPE REF TO zcl_excel_drawing,
             output TYPE REF TO zcl_excel_drawing,
           END OF ty_mapping_of_drawing,
           ty_mapping_of_drawings TYPE HASHED TABLE OF ty_mapping_of_drawing WITH UNIQUE KEY input,
           BEGIN OF ty_mapping,
             styles     TYPE ty_mapping_of_styles,
             drawings   TYPE ty_mapping_of_drawings,
             worksheets TYPE ty_mapping_of_worksheets,
           END OF ty_mapping.

    METHODS clone_worksheet
      IMPORTING
        io_input_worksheet TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(result)      TYPE REF TO zcl_excel_worksheet
      RAISING
        zcx_excel.

    METHODS clone_style
      IMPORTING
        io_style      TYPE REF TO zcl_excel_style
      RETURNING
        VALUE(result) TYPE REF TO zcl_excel_style.

    METHODS clone_drawing
      IMPORTING
        io_input_drawing TYPE REF TO zcl_excel_drawing
      RETURNING
        VALUE(result)    TYPE REF TO zcl_excel_drawing
      RAISING
        zcx_excel.

    DATA:
      output_workbook         TYPE REF TO zcl_excel,
      input_workbook          TYPE REF TO zcl_excel,
      mapping                 TYPE ty_mapping,
      "! Used to remove the default worksheet in the output workbook.
      count_worksheets_cloned TYPE i.

ENDCLASS.

CLASS lcl_app DEFINITION.
  PUBLIC SECTION.

    METHODS pai
      IMPORTING
        template TYPE string
        ucomm    TYPE sscrfields-ucomm
      RAISING
        zcx_excel.

    DATA: input_workbook TYPE REF TO zcl_excel READ-ONLY.

  PRIVATE SECTION.

    METHODS upload_binary_file
      IMPORTING
        path          TYPE csequence
      RETURNING
        VALUE(result) TYPE xstring.

    METHODS write_bin_file
      IMPORTING
        path     TYPE csequence
        contents TYPE xstring.

ENDCLASS.

CLASS lcl_excel_cloner IMPLEMENTATION.

  METHOD constructor.

  ENDMETHOD.

  METHOD clone_workbook.

    input_workbook = io_input_excel.
    output_workbook = NEW zcl_excel( ).

    DATA(lo_input_worksheets_iterator) = input_workbook->get_worksheets_iterator( ).

    " STYLES
    DATA(lo_style_iterator) = input_workbook->get_styles_iterator( ).
    WHILE lo_style_iterator->has_next( ).
      DATA(lo_style4) = CAST zcl_excel_style( lo_style_iterator->get_next( ) ).
      clone_style( io_style = lo_style4 ).
    ENDWHILE.

    " IMAGE DRAWINGS
    DATA(input_drawings_iterator) = input_workbook->get_drawings_iterator( zcl_excel_drawing=>type_image ).
    WHILE input_drawings_iterator->has_next( ).
      DATA(input_drawing) = CAST zcl_excel_drawing( input_drawings_iterator->get_next( ) ).
      DATA(output_drawing) = input_workbook->add_new_drawing( ).
      INSERT VALUE ty_mapping_of_drawing( input = input_drawing output = output_drawing ) INTO TABLE mapping-drawings.
    ENDWHILE.

    " WORKSHEETS
    WHILE lo_input_worksheets_iterator->has_next( ).
      DATA(lo_input_worksheet) = CAST zcl_excel_worksheet( lo_input_worksheets_iterator->get_next( ) ).
      clone_worksheet( io_input_worksheet = lo_input_worksheet ).
    ENDWHILE.

    result = output_workbook.

  ENDMETHOD.

  METHOD clone_worksheet.

    DATA: column_start   TYPE zexcel_cell_column_alpha,
          column_end     TYPE zexcel_cell_column_alpha,
          row_start      TYPE zexcel_cell_row,
          row_end        TYPE zexcel_cell_row,
          row            TYPE zexcel_cell_row,
          column_int     TYPE i,
          column_end_int TYPE i,
          output_drawing TYPE REF TO zcl_excel_drawing,
          lo_style       TYPE REF TO zcl_excel_style,
          value          TYPE zexcel_cell_value,
          formula        TYPE zexcel_cell_formula.

    DATA(output_worksheet) = output_workbook->get_worksheet_by_name( io_input_worksheet->get_title( ) ).
    IF output_worksheet IS BOUND.
      result = output_worksheet.
    ELSE.
      IF count_worksheets_cloned = 0.
        DATA(initial_output_worksheet) = output_workbook->get_worksheet_by_index( 1 ).
      ENDIF.
      result = output_workbook->add_new_worksheet( io_input_worksheet->get_title( ) ).
      IF count_worksheets_cloned = 0.
        output_workbook->delete_worksheet( initial_output_worksheet ).
      ENDIF.
    ENDIF.
    ADD 1 TO count_worksheets_cloned.

    " TABLES

    DATA(tables_iterator) = io_input_worksheet->get_tables_iterator( ).
    WHILE tables_iterator->has_next( ).
      DATA(table) = CAST zcl_excel_table( tables_iterator->get_next( ) ).

      DATA(components) = VALUE abap_component_tab(
              FOR <field> IN table->fieldcat
                ( name = <field>-fieldname
                  type = cl_abap_elemdescr=>get_string( ) ) ).
      DATA(rtti_table) = cl_abap_tabledescr=>get( p_line_type = cl_abap_structdescr=>get( components ) ).
      DATA dref_table TYPE REF TO data.
      FIELD-SYMBOLS <table> TYPE STANDARD TABLE.
      CREATE DATA dref_table TYPE HANDLE rtti_table.
      ASSIGN dref_table->* TO <table>.

*io_input_worksheet->get_
      result->bind_table(
        EXPORTING
          ip_table            = <table>
          it_field_catalog    = table->fieldcat
          is_table_settings   = table->settings
*          iv_default_descr    =
*          iv_no_line_if_empty = ABAP_FALSE
*        IMPORTING
*          es_table_settings   =
      ).
*        CATCH zcx_excel.    "
    ENDWHILE.

    " RANGES

    DATA(ranges_iterator) = result->get_ranges_iterator( ).
    WHILE ranges_iterator->has_next( ).
      DATA(range) = CAST zcl_excel_range( ranges_iterator->get_next( ) ).
*      range->
    ENDWHILE.


    " ROWS/COLUMNS/CELLS

    zcl_excel_common=>convert_range2column_a_row(
      EXPORTING
        i_range        = COND #( LET dim = io_input_worksheet->get_dimension_range( ) IN WHEN dim CA ':' THEN dim ELSE |{ dim }:{ dim }| )
      IMPORTING
        e_column_start = column_start
        e_column_end   = column_end
        e_row_start    = row_start
        e_row_end      = row_end ).
    column_int = zcl_excel_common=>convert_column2int( column_start ).
    column_end_int = zcl_excel_common=>convert_column2int( column_end ).

    " ROWS
    row = row_start.
    WHILE row <= row_end.
      result->get_row( row )->set_row_height( io_input_worksheet->get_row( row )->get_row_height( ) ).
      row = row + 1.
    ENDWHILE.

    " COLUMNS
    WHILE column_int <= column_end_int.
      result->get_column( column_int )->set_width( io_input_worksheet->get_column( column_int )->get_width( ) ).
      column_int = column_int + 1.
    ENDWHILE.

    " CELLS
    row = row_start.
    WHILE row <= row_end.

      column_int = zcl_excel_common=>convert_column2int( column_start ).
      column_end_int = zcl_excel_common=>convert_column2int( column_end ).
      WHILE column_int <= column_end_int.

        io_input_worksheet->get_cell(
          EXPORTING
            ip_column  = column_int
            ip_row     = row
          IMPORTING
            ep_value   = value
            ep_style   = lo_style
            ep_formula = formula ).

        IF value IS NOT INITIAL OR formula IS NOT INITIAL OR lo_style IS BOUND.
          result->set_cell(
              ip_column  = column_int
              ip_row     = row
              ip_value   = value
              ip_formula = formula
              ip_style   = COND #( WHEN lo_style IS BOUND THEN mapping-styles[ input = lo_style ]-output->get_guid( ) ) ).
        ENDIF.

        column_int = column_int + 1.
      ENDWHILE.

      row = row + 1.
    ENDWHILE.

    " DRAWINGS
    DATA(input_drawings_iterator) = io_input_worksheet->get_drawings_iterator( zcl_excel_drawing=>type_image ).
    WHILE input_drawings_iterator->has_next( ).
      DATA(input_drawing) = CAST zcl_excel_drawing( input_drawings_iterator->get_next( ) ).
      output_drawing = clone_drawing( io_input_drawing = input_drawing ).
      result->add_drawing( ip_drawing = output_drawing ).
    ENDWHILE.

    INSERT VALUE ty_mapping_of_worksheet( input = io_input_worksheet output = result ) INTO TABLE mapping-worksheets.

  ENDMETHOD.

  METHOD clone_style.

    result = output_workbook->add_new_style( ).
    result->alignment     = io_style->alignment.
    result->font          = io_style->font.
    result->borders       = io_style->borders.
    result->fill          = io_style->fill.
    result->number_format = io_style->number_format.
    result->protection    = io_style->protection.

    INSERT VALUE ty_mapping_of_style( input = io_style output = result ) INTO TABLE mapping-styles.

  ENDMETHOD.


  METHOD clone_drawing.

    result = mapping-drawings[ input = io_input_drawing ]-output.

    result->set_position( ip_from_row = io_input_drawing->get_from_row( ) + 1
                          ip_from_col = zcl_excel_common=>convert_column2alpha( io_input_drawing->get_from_col( ) + 1 )
                          ip_rowoff   = io_input_drawing->get_position( )-from-row_offset
                          ip_coloff   = io_input_drawing->get_position( )-from-col_offset ).

    result->set_media( ip_media      = io_input_drawing->get_media( )
                       ip_media_type = io_input_drawing->get_media_type( )
                       ip_width      = COND #( LET xstring4 = io_input_drawing->get_media( ) IN WHEN io_input_drawing->get_media_type( ) = 'png' THEN xstring4+16(4) ELSE 192 )
                       ip_height     = COND #( LET xstring4 = io_input_drawing->get_media( ) IN WHEN io_input_drawing->get_media_type( ) = 'png' THEN xstring4+20(4) ELSE 71 ) ).

  ENDMETHOD.

ENDCLASS.

CLASS lcl_app IMPLEMENTATION.

  METHOD pai.

    CASE ucomm.
      WHEN 'ONLI'.
        input_workbook = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( upload_binary_file( template ) ).
    ENDCASE.

  ENDMETHOD.


  METHOD upload_binary_file.
    DATA: filename TYPE string,
          length   TYPE i,
          lines    TYPE solix_tab.
    filename = path.
    cl_gui_frontend_services=>gui_upload(
      EXPORTING
        filename   = filename
        filetype   = 'BIN'
      IMPORTING
        filelength = length
      CHANGING
        data_tab   = lines
      EXCEPTIONS
        OTHERS     = 1 ).
    IF sy-subrc <> 0.
      MESSAGE 'Error while uploading' TYPE 'I' DISPLAY LIKE 'E'.
    ENDIF.
    result = cl_bcs_convert=>solix_to_xstring(
                     it_solix = lines
                     iv_size  = length ).
  ENDMETHOD.


  METHOD write_bin_file.

    DATA lt_xstring TYPE TABLE OF x255.
    DATA l_length TYPE i.
    DATA l_file_name TYPE string.

    l_file_name = path.

    CALL METHOD cl_swf_utl_convert_xstring=>xstring_to_table
      EXPORTING
        i_stream = contents
      IMPORTING
        e_table  = lt_xstring
      EXCEPTIONS
        OTHERS   = 3.

    l_length = xstrlen( contents ).

    CALL METHOD cl_gui_frontend_services=>gui_download
      EXPORTING
        bin_filesize = l_length
        filename     = l_file_name
        filetype     = 'BIN'
      CHANGING
        data_tab     = lt_xstring
      EXCEPTIONS
        OTHERS       = 3.
    IF sy-subrc <> 0.
      MESSAGE 'Error while downloading' TYPE 'I' DISPLAY LIKE 'E'.
    ENDIF.

  ENDMETHOD.

ENDCLASS.

FORM popup_f4
  USING
    tabname    TYPE tabname
    fieldname  TYPE fieldname
    display    TYPE abap_bool
  CHANGING
    returncode TYPE char1
    value      TYPE clike.

  CASE |{ tabname }-{ fieldname }|.
    WHEN 'ADMI_FILES-FILENAME'.
      PERFORM popup_f4_filename CHANGING value.
  ENDCASE.

ENDFORM.

FORM popup_f4_filename
  CHANGING
    value TYPE clike.

  DATA: lt_filetable     TYPE filetable,
        default_filename TYPE string,
        l_rc             TYPE i,
        l_action         TYPE i.
  FIELD-SYMBOLS <ls_file> TYPE file_table.

  default_filename = value.
  CALL METHOD cl_gui_frontend_services=>file_open_dialog
    EXPORTING
      window_title            = 'Select EXCEL file to be recreated by abap2xlsx'
      default_filename        = default_filename
    CHANGING
      file_table              = lt_filetable
      rc                      = l_rc
      user_action             = l_action
    EXCEPTIONS
      file_open_dialog_failed = 1
      cntl_error              = 2
      error_no_gui            = 3
      not_supported_by_gui    = 4
      OTHERS                  = 5.
  IF sy-subrc = 0 AND l_action = cl_gui_frontend_services=>action_ok.
    READ TABLE lt_filetable INDEX 1 ASSIGNING <ls_file>.
    IF sy-subrc = 0.
      value = <ls_file>-filename.
    ENDIF.
  ENDIF.

ENDFORM.

FORM popup_okcode
  TABLES
    fields     STRUCTURE sval
  USING
    ok_code    TYPE clike
  CHANGING
    error      STRUCTURE svale
    show_popup TYPE char1.
ENDFORM.

FORM demo_show_get_parameters
  TABLES
    parameters STRUCTURE rsparamsl_255
  CHANGING
    cancel TYPE abap_bool.

  DATA: lt_field     TYPE TABLE OF sval,
        l_returncode TYPE flag,
        parameter    TYPE rsparamsl_255.
  FIELD-SYMBOLS:
        <ls_field>    TYPE sval.

  CLEAR parameters.
  cancel = abap_false.

  REFRESH lt_field.
  APPEND INITIAL LINE TO lt_field ASSIGNING <ls_field>.
  <ls_field>-tabname = 'ADMI_FILES'.
  <ls_field>-fieldname = 'FILENAME'.
  CALL FUNCTION 'POPUP_GET_VALUES_USER_BUTTONS'
    EXPORTING
      f4_formname       = 'POPUP_F4'
      f4_programname    = sy-repid
      formname          = 'POPUP_OKCODE'
      programname       = sy-repid
      popup_title       = 'Enter name of file (same directory)'
      ok_pushbuttontext = 'Select'
    IMPORTING
      returncode        = l_returncode
    TABLES
      fields            = lt_field
    EXCEPTIONS
      error_in_fields   = 1
      OTHERS            = 2.
  CASE l_returncode.
    WHEN space. " Enter
      READ TABLE lt_field WITH KEY fieldname = 'FILENAME' ASSIGNING <ls_field>.
      ASSERT sy-subrc = 0.
      CLEAR parameter.
      parameter-selname = 'TEMPLATE'.
      parameter-kind = 'P'.
      parameter-low = <ls_field>-value.
      APPEND parameter TO parameters.
    WHEN 'A'. " Abort
      cancel = abap_true.
  ENDCASE.
ENDFORM.

FORM dynp_values_read USING fieldname CHANGING value.
  DATA lt_dynpfield TYPE TABLE OF dynpread.
  DATA ls_dynpfield TYPE dynpread.
  REFRESH lt_dynpfield.
  ls_dynpfield-fieldname = fieldname.
  APPEND ls_dynpfield TO lt_dynpfield.
  CALL FUNCTION 'DYNP_VALUES_READ'
    EXPORTING
      dyname     = sy-repid
      dynumb     = sy-dynnr
    TABLES
      dynpfields = lt_dynpfield
    EXCEPTIONS
      OTHERS     = 9.
  IF sy-subrc = 0.
    READ TABLE lt_dynpfield WITH KEY fieldname = fieldname INTO ls_dynpfield.
    IF sy-subrc = 0.
      value = ls_dynpfield-fieldvalue.
    ENDIF.
  ENDIF.
ENDFORM.

TABLES sscrfields.

PARAMETERS template TYPE string LOWER CASE.

LOAD-OF-PROGRAM.
  DATA(app) = NEW lcl_app( ).

AT SELECTION-SCREEN ON VALUE-REQUEST FOR template.
  PERFORM dynp_values_read USING 'TEMPLATE' CHANGING template.
  PERFORM popup_f4_filename CHANGING template.

AT SELECTION-SCREEN.
  TRY.
      app->pai( ucomm = sscrfields-ucomm template = template )."output_path = output ).
    CATCH cx_root INTO DATA(lx_root).
      MESSAGE lx_root TYPE 'E'.
  ENDTRY.
  ASSERT 1 = 1.

START-OF-SELECTION.
  TRY.
      DATA(lo_excel) = NEW lcl_excel_cloner( )->clone_workbook( app->input_workbook ).
      lcl_output=>output( cl_excel = lo_excel ).
    CATCH cx_root INTO DATA(lx_root).
      MESSAGE lx_root TYPE 'E'.
  ENDTRY.
  ASSERT 1 = 1.
