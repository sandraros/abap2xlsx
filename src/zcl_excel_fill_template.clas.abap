CLASS zcl_excel_fill_template DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    DATA mt_sheet TYPE zexcel_template_t_sheet_title READ-ONLY.
    DATA mt_range TYPE zexcel_template_t_range READ-ONLY.
    DATA mt_var TYPE zexcel_template_t_var READ-ONLY.
    DATA mo_excel TYPE REF TO zcl_excel READ-ONLY.

    CLASS-METHODS create
      IMPORTING
        !io_excel                 TYPE REF TO zcl_excel
      RETURNING
        VALUE(eo_template_filler) TYPE REF TO zcl_excel_fill_template.

    METHODS fill_sheet
      IMPORTING
        !iv_data TYPE zexcel_s_template_data .

  PROTECTED SECTION.
  PRIVATE SECTION.

    METHODS fill_range
      IMPORTING
        !iv_sheet              TYPE zexcel_sheet_title
        !iv_parent             TYPE zexcel_cell_row
        !iv_data               TYPE data
        VALUE(iv_range_length) TYPE zexcel_cell_row
        !io_sheet              TYPE REF TO zcl_excel_worksheet
      CHANGING
        !ct_cells              TYPE zexcel_template_t_cell_data
        !cv_diff               TYPE zexcel_cell_row
        !ct_merged_cells       TYPE zcl_excel_worksheet=>mty_ts_merge .
    METHODS get_range .
    METHODS validate_range
      IMPORTING
        !io_range TYPE REF TO zcl_excel_range .
    METHODS discard_overlapped .
    METHODS sign_range .
    METHODS find_var .

ENDCLASS.



CLASS zcl_excel_fill_template IMPLEMENTATION.


  METHOD create.

    CREATE OBJECT eo_template_filler .

    eo_template_filler->mo_excel = io_excel.
    eo_template_filler->get_range( ).
    eo_template_filler->discard_overlapped( ).
    eo_template_filler->sign_range( ).
    eo_template_filler->find_var( ).

  ENDMETHOD.

  METHOD discard_overlapped.
    DATA:
       lt_range TYPE zexcel_template_t_range.
    FIELD-SYMBOLS:
      <fs_range> TYPE zexcel_template_s_range,
      <fs_tmp>   TYPE zexcel_template_s_range.

    SORT mt_range BY sheet  start  stop.

    LOOP AT mt_range ASSIGNING <fs_range>.

      LOOP AT mt_range ASSIGNING <fs_tmp> WHERE sheet =  <fs_range>-sheet
                                            AND name  <> <fs_range>-name
                                            AND stop  >= <fs_range>-start
                                            AND start <  <fs_range>-start
                                            AND stop  <  <fs_range>-stop.
        EXIT.
      ENDLOOP.

      IF sy-subrc NE 0.
        APPEND <fs_range> TO lt_range.
      ENDIF.

    ENDLOOP.

    mt_range = lt_range.

    SORT mt_range BY sheet  start  stop DESCENDING.

  ENDMETHOD.


  METHOD fill_range.

    DATA: tmp_cells_template        TYPE zexcel_template_t_cell_data,
          lt_cells_result           TYPE zexcel_template_t_cell_data,
          tmp_cells                 TYPE zexcel_template_t_cell_data,
          ls_cell                   TYPE zexcel_s_cell_data,

          tmp_merged_cells_template TYPE zcl_excel_worksheet=>mty_ts_merge,
          lt_merged_cells_result    TYPE zcl_excel_worksheet=>mty_ts_merge,
          lt_merged_cells_half      TYPE zcl_excel_worksheet=>mty_ts_merge,
          tmp_merged_cells          TYPE zcl_excel_worksheet=>mty_ts_merge,
          ls_merged_cell            LIKE LINE OF tmp_merged_cells,

          lv_start                  TYPE i,
          lv_stop                   TYPE i,

          col_str                   TYPE string,

          result_tab                TYPE match_result_tab,
          lv_search                 TYPE string,
          lv_var_name               TYPE string,
          lo_value_new              TYPE REF TO data,
          lo_addit                  TYPE REF TO cl_abap_elemdescr,
          lv_str                    TYPE string,
          lv_value_type             TYPE c.

    FIELD-SYMBOLS:
      <fs_table>   TYPE ANY TABLE,
      <fs_line>    TYPE any,
      <fs_range>   TYPE zexcel_template_s_range,
      <fs_cell>    TYPE zexcel_s_cell_data,
      <fs_numeric> TYPE numeric,
      <fs_result>  TYPE match_result,
      <fs_var>     TYPE any,
                   <fs_typekind_int8> TYPE abap_typekind.


    cv_diff = cv_diff +  iv_range_length .

    lv_start = 1.


* recursive fill nested range

    LOOP AT mt_range ASSIGNING <fs_range> WHERE sheet = iv_sheet
                                            AND parent = iv_parent.


      lv_stop = <fs_range>-start - 1.

*      update cells before any range

      LOOP AT ct_cells INTO ls_cell  WHERE cell_row >= lv_start AND cell_row <= lv_stop .
        ls_cell-cell_row =  ls_cell-cell_row + cv_diff.
        col_str = zcl_excel_common=>convert_column2alpha( ls_cell-cell_column ).

        ls_cell-cell_coords = ls_cell-cell_row.
        CONCATENATE col_str ls_cell-cell_coords INTO ls_cell-cell_coords.
        CONDENSE ls_cell-cell_coords NO-GAPS.

        APPEND ls_cell TO lt_cells_result.
      ENDLOOP.



*      update merged cells before range

      LOOP AT ct_merged_cells INTO ls_merged_cell WHERE row_from >=  lv_start AND row_to <= lv_stop.
        ls_merged_cell-row_from = ls_merged_cell-row_from + cv_diff.
        ls_merged_cell-row_to = ls_merged_cell-row_to + cv_diff.

        APPEND ls_merged_cell TO lt_merged_cells_result.

      ENDLOOP.



      lv_start = <fs_range>-stop + 1.



      CLEAR:
       tmp_cells_template,
       tmp_merged_cells_template.


*copy cell template
      LOOP AT ct_cells INTO ls_cell WHERE cell_row >= <fs_range>-start AND cell_row <= <fs_range>-stop.
        APPEND ls_cell TO tmp_cells_template.
      ENDLOOP.

      LOOP AT ct_merged_cells INTO ls_merged_cell WHERE row_from >= <fs_range>-start AND row_to <= <fs_range>-stop.
        APPEND ls_merged_cell TO tmp_merged_cells_template.
      ENDLOOP.


      ASSIGN COMPONENT <fs_range>-name OF STRUCTURE iv_data TO <fs_table>.
      CHECK sy-subrc = 0.

      cv_diff = cv_diff - <fs_range>-length.

*merge each line of data table with template
      LOOP AT <fs_table> ASSIGNING <fs_line>.
*        make local copy
        tmp_cells = tmp_cells_template.
        tmp_merged_cells = tmp_merged_cells_template.

*fill data

        fill_range(
          EXPORTING
            io_sheet        = io_sheet
            iv_sheet        = iv_sheet
            iv_parent       = <fs_range>-id
            iv_data         = <fs_line>
            iv_range_length = <fs_range>-length
          CHANGING
            ct_cells        = tmp_cells
            ct_merged_cells = tmp_merged_cells
            cv_diff         = cv_diff ).

*collect data

        APPEND LINES OF tmp_cells TO lt_cells_result.
        APPEND LINES OF tmp_merged_cells TO lt_merged_cells_result.

      ENDLOOP.

    ENDLOOP.


    IF <fs_range> IS ASSIGNED.

      LOOP AT ct_cells INTO ls_cell WHERE cell_row > <fs_range>-stop .
        ls_cell-cell_row =  ls_cell-cell_row + cv_diff.
        col_str = zcl_excel_common=>convert_column2alpha( ls_cell-cell_column ).

        ls_cell-cell_coords = ls_cell-cell_row.
        CONCATENATE col_str ls_cell-cell_coords INTO ls_cell-cell_coords.
        CONDENSE ls_cell-cell_coords NO-GAPS.

        APPEND ls_cell TO lt_cells_result.
      ENDLOOP.

      ct_cells = lt_cells_result.

      LOOP AT ct_merged_cells INTO ls_merged_cell WHERE row_from > <fs_range>-stop.
        ls_merged_cell-row_from = ls_merged_cell-row_from + cv_diff.
        ls_merged_cell-row_to = ls_merged_cell-row_to + cv_diff.

        APPEND ls_merged_cell TO lt_merged_cells_result.
      ENDLOOP.

      ct_merged_cells = lt_merged_cells_result.

    ELSE.


      LOOP AT ct_cells ASSIGNING <fs_cell>.
        <fs_cell>-cell_row =  <fs_cell>-cell_row + cv_diff.
        col_str = zcl_excel_common=>convert_column2alpha( <fs_cell>-cell_column ).

        <fs_cell>-cell_coords = <fs_cell>-cell_row.
        CONCATENATE col_str <fs_cell>-cell_coords INTO <fs_cell>-cell_coords.
        CONDENSE <fs_cell>-cell_coords NO-GAPS.
      ENDLOOP.

      LOOP AT ct_merged_cells INTO ls_merged_cell .
        ls_merged_cell-row_from = ls_merged_cell-row_from + cv_diff.
        ls_merged_cell-row_to = ls_merged_cell-row_to + cv_diff.

        APPEND ls_merged_cell TO lt_merged_cells_result.
      ENDLOOP.

      ct_merged_cells = lt_merged_cells_result.

    ENDIF.


*check if variables in this range
    READ TABLE mt_var TRANSPORTING NO FIELDS WITH KEY sheet = iv_sheet parent = iv_parent.

    IF sy-subrc = 0.

        ASSIGN ('CL_ABAP_TYPEDESCR=>TYPEKIND_INT8') TO <fs_typekind_int8>.
        IF sy-subrc <> 0.
          ASSIGN space TO <fs_typekind_int8>. "not used as typekind!
        ENDIF.


*      replace variables of current range with data
      LOOP AT ct_cells ASSIGNING <fs_cell>.


        REFRESH result_tab.

        FIND ALL OCCURRENCES OF REGEX '\[[^\]]*\]' IN <fs_cell>-cell_value  RESULTS result_tab.

        SORT result_tab BY offset DESCENDING .

        LOOP AT result_tab ASSIGNING <fs_result>.
          lv_search = <fs_cell>-cell_value+<fs_result>-offset(<fs_result>-length).
          lv_var_name = lv_search.

          TRANSLATE lv_var_name TO UPPER CASE.
          TRANSLATE lv_var_name USING '[ ] '.
          CONDENSE lv_var_name .

          ASSIGN COMPONENT lv_var_name OF STRUCTURE iv_data TO <fs_var>.
          CHECK sy-subrc = 0.

          lv_str = <fs_var>.

          REPLACE ALL OCCURRENCES OF lv_search IN <fs_cell>-cell_value  WITH lv_str.

        ENDLOOP.

        CHECK  lines( result_tab )  = 1.

        CHECK  <fs_cell>-cell_value CO '1234567890. '.


        DESCRIBE FIELD <fs_var> TYPE lv_value_type.

        CHECK  lv_value_type = cl_abap_typedescr=>typekind_int
            OR lv_value_type = 'P'
            OR lv_value_type = 's'
            OR lv_value_type = <fs_typekind_int8>
            OR lv_value_type = 'a'
            OR lv_value_type = 'e'
            OR lv_value_type = 'F'.

        CASE lv_value_type.
          WHEN cl_abap_typedescr=>typekind_int OR cl_abap_typedescr=>typekind_int1 OR cl_abap_typedescr=>typekind_int2
            OR <fs_typekind_int8>. "Allow INT8 types columns
            lo_addit = cl_abap_elemdescr=>get_i( ).
            CREATE DATA lo_value_new TYPE HANDLE lo_addit.
            ASSIGN lo_value_new->* TO <fs_numeric>.
            IF sy-subrc = 0.
              <fs_numeric> = <fs_var>.
              <fs_cell>-cell_value = zcl_excel_common=>number_to_excel_string( ip_value = <fs_numeric> ).
              CLEAR <fs_cell>-data_type.
            ENDIF.

          WHEN cl_abap_typedescr=>typekind_float OR cl_abap_typedescr=>typekind_packed.
            lo_addit = cl_abap_elemdescr=>get_f( ).
            CREATE DATA lo_value_new TYPE HANDLE lo_addit.
            ASSIGN lo_value_new->* TO <fs_numeric>.
            IF sy-subrc = 0.
              <fs_numeric> = <fs_var>.
              <fs_cell>-cell_value = zcl_excel_common=>number_to_excel_string( ip_value = <fs_numeric> ).
              CLEAR <fs_cell>-data_type.
            ENDIF.
        ENDCASE.

      ENDLOOP.
    ENDIF.


  ENDMETHOD.


  METHOD fill_sheet.

    DATA: lo_worksheet     TYPE REF TO zcl_excel_worksheet,
          lt_sheet_content TYPE  zexcel_template_t_cell_data,
          lt_merged_cells  TYPE zcl_excel_worksheet=>mty_ts_merge,
          lv_dif           TYPE i.

    FIELD-SYMBOLS:
      <any_data>         TYPE any,
      <fs_sheet_content> TYPE zexcel_s_cell_data,
      <fs_merged_cell>   TYPE zcl_excel_worksheet=>mty_merge.


    lo_worksheet = mo_excel->get_worksheet_by_name( iv_data-sheet ).

    lt_sheet_content = lo_worksheet->sheet_content.
    lt_merged_cells = lo_worksheet->mt_merged_cells.

    ASSIGN iv_data-data->* TO <any_data>.

    fill_range(
      EXPORTING
        io_sheet        = lo_worksheet
        iv_range_length = 0
        iv_sheet        = iv_data-sheet
        iv_parent       = 0
        iv_data         = <any_data>
      CHANGING
        ct_cells        = lt_sheet_content
        ct_merged_cells = lt_merged_cells
        cv_diff         = lv_dif  ).


    REFRESH  lo_worksheet->sheet_content.

    LOOP AT lt_sheet_content ASSIGNING <fs_sheet_content>.
      INSERT <fs_sheet_content> INTO TABLE lo_worksheet->sheet_content.
    ENDLOOP.

    DATA(merged_cells) = lo_worksheet->get_merge( ).
    LOOP AT merged_cells ASSIGNING FIELD-SYMBOL(<merged_cell>).
      zcl_excel_common=>convert_range2column_a_row(
        EXPORTING
          i_range        = <merged_cell>
        IMPORTING
          e_column_start = DATA(column_start)
*    e_column_end   =
          e_row_start    = DATA(row_start)
*    e_row_end      =
*    e_sheet        =
      ).
*  CATCH zcx_excel.    "
      lo_worksheet->delete_merge( ip_cell_column = column_start ip_cell_row = row_start ).
    ENDLOOP.
*    REFRESH  lo_worksheet->mt_merged_cells.

    LOOP AT lt_merged_cells ASSIGNING <fs_merged_cell>.
      lo_worksheet->set_merge(
*      EXPORTING
          ip_column_start = <fs_merged_cell>-col_from
          ip_column_end   = <fs_merged_cell>-col_to
          ip_row          = <fs_merged_cell>-row_from
          ip_row_to       = <fs_merged_cell>-row_to ).
*      CATCH zcx_excel.    "
*      INSERT <fs_merged_cell> INTO TABLE lo_worksheet->mt_merged_cells.
    ENDLOOP.



  ENDMETHOD.


  METHOD find_var.

    DATA: row            TYPE i,
          column         TYPE i,
          col_str        TYPE string,
          value          TYPE string,

          lo_worksheet   TYPE REF TO zcl_excel_worksheet,
          ls_var         TYPE zexcel_template_s_var,

          highest_column TYPE zexcel_cell_column,
          highest_row    TYPE int4,

          result_tab     TYPE match_result_tab,
          lv_search      TYPE string,
          lv_replace     TYPE string.

    FIELD-SYMBOLS:
      <fs_result> TYPE match_result,
      <fs_range>  TYPE zexcel_template_s_range,
      <fs_sheet>  TYPE zexcel_template_sheet_title.


    LOOP AT mt_sheet ASSIGNING <fs_sheet>.

      lo_worksheet ?= mo_excel->get_worksheet_by_name(  <fs_sheet> ).
      row = 1.
      column = 1.

      highest_column = lo_worksheet->get_highest_column( ).
      highest_row    = lo_worksheet->get_highest_row( ).

      WHILE row <= highest_row.
        WHILE column <= highest_column.
          col_str = zcl_excel_common=>convert_column2alpha( column ).
          lo_worksheet->get_cell(
            EXPORTING
              ip_column = col_str
              ip_row    = row
            IMPORTING
              ep_value = value
          ).


          FIND ALL OCCURRENCES OF REGEX '\[[^\]]*\]' IN value RESULTS result_tab.

          LOOP AT result_tab ASSIGNING <fs_result>.
            lv_search = value+<fs_result>-offset(<fs_result>-length).
            lv_replace = lv_search.

            TRANSLATE lv_replace TO UPPER CASE.

            CLEAR ls_var.

            ls_var-sheet = <fs_sheet>.
            ls_var-name = lv_replace.
            TRANSLATE ls_var-name USING '[ ] '.
            CONDENSE ls_var-name .

            LOOP AT mt_range ASSIGNING <fs_range> WHERE sheet = <fs_sheet>
                                                    AND start <= row
                                                    AND stop >= row.
              ls_var-parent = <fs_range>-id.
              EXIT.
            ENDLOOP.

            READ TABLE mt_var TRANSPORTING NO FIELDS WITH KEY sheet = ls_var-sheet name = ls_var-name parent = ls_var-parent.
            IF sy-subrc NE 0.
              APPEND ls_var TO mt_var.
            ENDIF.


          ENDLOOP.
          column = column + 1.
        ENDWHILE.
        column = 1.
        row = row + 1.
      ENDWHILE.
    ENDLOOP.

    SORT mt_range BY id .
  ENDMETHOD.


  METHOD get_range.

    DATA:
      lo_worksheets_iterator TYPE REF TO cl_object_collection_iterator,
      lo_worksheet           TYPE REF TO zcl_excel_worksheet,
      lo_range_iterator      TYPE REF TO cl_object_collection_iterator,
      lo_range               TYPE REF TO zcl_excel_range.


    lo_worksheets_iterator = mo_excel->get_worksheets_iterator( ).


    WHILE lo_worksheets_iterator->has_next( ) = 'X'.
      lo_worksheet ?= lo_worksheets_iterator->get_next( ).
      APPEND lo_worksheet->get_title( ) TO mt_sheet.
    ENDWHILE.

    lo_range_iterator = mo_excel->get_ranges_iterator( ).

    WHILE lo_range_iterator->has_next(  ) = 'X'.
      lo_range ?= lo_range_iterator->get_next( ).
      validate_range( lo_range ).
    ENDWHILE.

  ENDMETHOD.


  METHOD sign_range.

    DATA: lv_tabix TYPE i.
    FIELD-SYMBOLS:
      <fs_range>     TYPE zexcel_template_s_range,
      <fs_range_tmp> TYPE zexcel_template_s_range.

    LOOP AT mt_range ASSIGNING <fs_range>.
      <fs_range>-id = sy-tabix.
    ENDLOOP.

    LOOP AT mt_range ASSIGNING  <fs_range>.
      lv_tabix = sy-tabix + 1.

      LOOP AT mt_range ASSIGNING  <fs_range_tmp>
                                  FROM lv_tabix
                                  WHERE sheet = <fs_range>-sheet.

        IF <fs_range_tmp>-start >= <fs_range>-start AND <fs_range_tmp>-stop <= <fs_range>-stop.
          <fs_range_tmp>-parent = <fs_range>-id.
        ENDIF.

      ENDLOOP.

    ENDLOOP.

    SORT mt_range BY id DESCENDING.
  ENDMETHOD.


  METHOD validate_range.

    DATA: name     TYPE string,
          value    TYPE string,
          start    TYPE string,
          stop     TYPE string,
          lv_sheet TYPE string,
          lv_value TYPE string,
          lt_str   TYPE TABLE OF string,
          lv_start TYPE string,
          lv_stop  TYPE string.

    FIELD-SYMBOLS: <fs_range> TYPE zexcel_template_s_range.


    name = io_range->name.
    TRANSLATE name TO UPPER CASE.
    value = io_range->get_value( ).

    SPLIT value AT '!' INTO lv_sheet lv_value.

    SPLIT lv_value AT ':' INTO start stop.

    SPLIT start AT '$' INTO TABLE lt_str.

    IF lines( lt_str ) > 2.
      TRY.
          zcl_excel_common=>convert_range2column_a_row(
            EXPORTING
              i_range        = value
            IMPORTING
              e_column_start = DATA(column_start)
              e_column_end   = DATA(column_end)
              e_row_start    = DATA(row_start)
              e_row_end      = DATA(row_end)
          ).
        CATCH zcx_excel.    "
          RETURN.
      ENDTRY.
      IF column_start = 'A' AND column_end = 'XFD'.
        lv_start = |{ row_start }|.
        lv_stop = |{ row_end }|.
        CLEAR lt_str.
        CLEAR stop.
      ELSE.
        RETURN.
      ENDIF.
    ENDIF.

    READ TABLE lt_str INTO lv_start INDEX 2.

    IF lv_start CO '0123456789'.
      APPEND INITIAL LINE TO mt_range ASSIGNING <fs_range>.
      <fs_range>-sheet = lv_sheet.
      <fs_range>-name = name.
      <fs_range>-start = lv_start.

      SPLIT stop AT '$' INTO TABLE lt_str.
      READ TABLE lt_str INTO lv_stop INDEX 2.
      <fs_range>-stop = lv_stop.

      <fs_range>-length = <fs_range>-stop - <fs_range>-start + 1.

    ENDIF.

  ENDMETHOD.
ENDCLASS.
