#include <ruby.h>
#include "xlsxwriter.h"
#include "workbook.h"
#include "workbook_properties.h"
#include "worksheet.h"

VALUE mXlsxWriter;
VALUE cWorkbook;
VALUE cWorksheet;
VALUE mXlsxFormat;
VALUE cWorkbookProperties;

void Init_xlsxwriter() {
  mXlsxWriter = rb_define_module("XlsxWriter");
  cWorkbook = rb_define_class_under(mXlsxWriter, "Workbook", rb_cObject);
  cWorksheet = rb_define_class_under(mXlsxWriter, "Worksheet", rb_cObject);
  mXlsxFormat = rb_define_module_under(mXlsxWriter, "Format");
  cWorkbookProperties = rb_define_class_under(cWorkbook, "Properties", rb_cObject);

  rb_define_alloc_func(cWorkbook, workbook_alloc);
  rb_define_singleton_method(cWorkbook, "new", workbook_new_, -1);
  rb_define_alias(rb_singleton_class(cWorkbook), "open", "new");
  rb_define_method(cWorkbook, "initialize", workbook_init, -1);
  rb_define_method(cWorkbook, "close", workbook_release, 0);
  rb_define_method(cWorkbook, "add_worksheet", workbook_add_worksheet_, -1);
  rb_define_method(cWorkbook, "add_format", workbook_add_format_, 2);
  // rb_define_method(cWorkbook, "add_chart", workbook_add_format_, 2);
  rb_define_method(cWorkbook, "set_default_xf_indices", workbook_set_default_xf_indices_, 0);
  rb_define_method(cWorkbook, "properties", workbook_properties_, 0);
  rb_define_method(cWorkbook, "define_name", workbook_define_name_, 2);
  rb_define_method(cWorkbook, "validate_worksheet_name", workbook_validate_worksheet_name_, 1);

  rb_define_attr(cWorkbook, "font_sizes", 1, 0);


  rb_define_alloc_func(cWorksheet, worksheet_alloc);
  rb_define_method(cWorksheet, "initialize", worksheet_init, -1);
  rb_define_method(cWorksheet, "free", worksheet_release, 0);
  rb_define_method(cWorksheet, "write_string", worksheet_write_string_, -1);
  rb_define_method(cWorksheet, "write_number", worksheet_write_number_, -1);
  rb_define_method(cWorksheet, "write_formula", worksheet_write_formula_, -1);
  rb_define_method(cWorksheet, "write_array_formula", worksheet_write_array_formula_, -1);
  rb_define_method(cWorksheet, "write_datetime", worksheet_write_datetime_, -1);
  rb_define_method(cWorksheet, "write_url", worksheet_write_url_, -1);
  rb_define_method(cWorksheet, "write_boolean", worksheet_write_boolean_, -1);
  rb_define_method(cWorksheet, "write_blank", worksheet_write_blank_, -1);
  rb_define_method(cWorksheet, "write_formula_num", worksheet_write_formula_num_, -1);
  rb_define_method(cWorksheet, "set_row", worksheet_set_row_, 2);
  rb_define_method(cWorksheet, "set_column", worksheet_set_column_, 3);
  rb_define_method(cWorksheet, "insert_image", worksheet_insert_image_, -1);
  rb_define_method(cWorksheet, "insert_chart", worksheet_insert_chart_, -1);
  rb_define_method(cWorksheet, "merge_range", worksheet_merge_range_, -1);
  rb_define_method(cWorksheet, "autofilter", worksheet_autofilter_, -1);
  rb_define_method(cWorksheet, "activate", worksheet_activate_, 0);
  rb_define_method(cWorksheet, "select", worksheet_select_, 0);
  rb_define_method(cWorksheet, "hide", worksheet_hide_, 0);
  rb_define_method(cWorksheet, "set_first_sheet", worksheet_set_first_sheet_, 0);
  rb_define_method(cWorksheet, "freeze_panes", worksheet_freeze_panes_, -1);
  rb_define_method(cWorksheet, "split_panes", worksheet_split_panes_, 2);
  rb_define_method(cWorksheet, "set_selection", worksheet_set_selection_, -1);
  rb_define_method(cWorksheet, "set_landscape", worksheet_set_landscape_, 0);
  rb_define_method(cWorksheet, "set_portrait", worksheet_set_portrait_, 0);
  rb_define_method(cWorksheet, "set_page_view", worksheet_set_page_view_, 0);
  rb_define_method(cWorksheet, "paper=", worksheet_set_paper_, 1);
  rb_define_method(cWorksheet, "set_margins", worksheet_set_margins_, 4);
  rb_define_method(cWorksheet, "set_header", worksheet_set_header_, 1);
  rb_define_method(cWorksheet, "set_footer", worksheet_set_footer_, 1);
  rb_define_method(cWorksheet, "h_pagebreaks=", worksheet_set_h_pagebreaks_, 1);
  rb_define_method(cWorksheet, "v_pagebreaks=", worksheet_set_v_pagebreaks_, 1);
  rb_define_method(cWorksheet, "print_across", worksheet_print_across_, 0);
  rb_define_method(cWorksheet, "zoom=", worksheet_set_zoom_, 1);
  rb_define_method(cWorksheet, "gridlines=", worksheet_gridlines_, 1);
  rb_define_method(cWorksheet, "center_horizontally", worksheet_center_horizontally_, 0);
  rb_define_method(cWorksheet, "center_vertically", worksheet_center_vertically_, 0);
  rb_define_method(cWorksheet, "print_row_col_headers", worksheet_print_row_col_headers_, 0);
  rb_define_method(cWorksheet, "repeat_rows", worksheet_repeat_rows_, 2);
  rb_define_method(cWorksheet, "repeat_columns", worksheet_repeat_columns_, 2);
  rb_define_method(cWorksheet, "print_area", worksheet_print_area_, -1);
  rb_define_method(cWorksheet, "fit_to_pages", worksheet_fit_to_pages_, 2);
  rb_define_method(cWorksheet, "start_page=", worksheet_set_start_page_, 1);
  rb_define_method(cWorksheet, "print_scale=", worksheet_set_print_scale_, 1);
  rb_define_method(cWorksheet, "right_to_left", worksheet_right_to_left_, 0);
  rb_define_method(cWorksheet, "hide_zero", worksheet_hide_zero_, 0);
  rb_define_method(cWorksheet, "tab_color=", worksheet_set_tab_color_, 1);
  rb_define_method(cWorksheet, "protect", worksheet_protect_, -1);
  rb_define_method(cWorksheet, "set_default_row", worksheet_set_default_row_, 2);

  rb_define_method(cWorksheet, "vertical_dpi", worksheet_get_vertical_dpi_, 0);
  rb_define_method(cWorksheet, "vertical_dpi=", worksheet_set_vertical_dpi_, 1);

  rb_define_method(cWorkbookProperties, "initialize", workbook_properties_init_, 1);
  rb_define_method(cWorkbookProperties, "[]=", workbook_properties_set_, 2);
#define DEF_PROP_HANDLER(prop) rb_define_method(cWorkbookProperties, #prop "=", workbook_properties_set_dir_, 1);
  DEF_PROP_HANDLER(title);
  DEF_PROP_HANDLER(subject);
  DEF_PROP_HANDLER(author);
  DEF_PROP_HANDLER(manager);
  DEF_PROP_HANDLER(company);
  DEF_PROP_HANDLER(category);
  DEF_PROP_HANDLER(keywords);
  DEF_PROP_HANDLER(comments);
  DEF_PROP_HANDLER(status);
  DEF_PROP_HANDLER(hyperlink_base);
#undef DEF_PROP_HANDLER


#define MAP_LXW_FMT_CONST(name) rb_define_const(mXlsxFormat, #name, INT2NUM(LXW_##name))
  MAP_LXW_FMT_CONST(COLOR_BLACK);
  MAP_LXW_FMT_CONST(COLOR_BLUE);
  MAP_LXW_FMT_CONST(COLOR_BROWN);
  MAP_LXW_FMT_CONST(COLOR_CYAN);
  MAP_LXW_FMT_CONST(COLOR_GRAY);
  MAP_LXW_FMT_CONST(COLOR_GREEN);
  MAP_LXW_FMT_CONST(COLOR_LIME);
  MAP_LXW_FMT_CONST(COLOR_MAGENTA);
  MAP_LXW_FMT_CONST(COLOR_NAVY);
  MAP_LXW_FMT_CONST(COLOR_ORANGE);
  MAP_LXW_FMT_CONST(COLOR_PINK);
  MAP_LXW_FMT_CONST(COLOR_PURPLE);
  MAP_LXW_FMT_CONST(COLOR_RED);
  MAP_LXW_FMT_CONST(COLOR_SILVER);
  MAP_LXW_FMT_CONST(COLOR_WHITE);
  MAP_LXW_FMT_CONST(COLOR_YELLOW);

  MAP_LXW_FMT_CONST(UNDERLINE_SINGLE);
  MAP_LXW_FMT_CONST(UNDERLINE_DOUBLE);
  MAP_LXW_FMT_CONST(UNDERLINE_SINGLE_ACCOUNTING);
  MAP_LXW_FMT_CONST(UNDERLINE_DOUBLE_ACCOUNTING);

  MAP_LXW_FMT_CONST(FONT_SUPERSCRIPT);
  MAP_LXW_FMT_CONST(FONT_SUBSCRIPT);

  MAP_LXW_FMT_CONST(ALIGN_LEFT);
  MAP_LXW_FMT_CONST(ALIGN_CENTER);
  MAP_LXW_FMT_CONST(ALIGN_RIGHT);
  MAP_LXW_FMT_CONST(ALIGN_FILL);
  MAP_LXW_FMT_CONST(ALIGN_JUSTIFY);
  MAP_LXW_FMT_CONST(ALIGN_CENTER_ACROSS);
  MAP_LXW_FMT_CONST(ALIGN_DISTRIBUTED);

  MAP_LXW_FMT_CONST(ALIGN_VERTICAL_TOP);
  MAP_LXW_FMT_CONST(ALIGN_VERTICAL_BOTTOM);
  MAP_LXW_FMT_CONST(ALIGN_VERTICAL_CENTER);
  MAP_LXW_FMT_CONST(ALIGN_VERTICAL_JUSTIFY);
  MAP_LXW_FMT_CONST(ALIGN_VERTICAL_DISTRIBUTED);

  MAP_LXW_FMT_CONST(PATTERN_SOLID);
  MAP_LXW_FMT_CONST(PATTERN_MEDIUM_GRAY);
  MAP_LXW_FMT_CONST(PATTERN_DARK_GRAY);
  MAP_LXW_FMT_CONST(PATTERN_LIGHT_GRAY);
  MAP_LXW_FMT_CONST(PATTERN_DARK_HORIZONTAL);
  MAP_LXW_FMT_CONST(PATTERN_DARK_VERTICAL);
  MAP_LXW_FMT_CONST(PATTERN_DARK_DOWN);
  MAP_LXW_FMT_CONST(PATTERN_DARK_UP);
  MAP_LXW_FMT_CONST(PATTERN_DARK_GRID);
  MAP_LXW_FMT_CONST(PATTERN_DARK_TRELLIS);
  MAP_LXW_FMT_CONST(PATTERN_LIGHT_HORIZONTAL);
  MAP_LXW_FMT_CONST(PATTERN_LIGHT_VERTICAL);
  MAP_LXW_FMT_CONST(PATTERN_LIGHT_DOWN);
  MAP_LXW_FMT_CONST(PATTERN_LIGHT_UP);
  MAP_LXW_FMT_CONST(PATTERN_LIGHT_GRID);
  MAP_LXW_FMT_CONST(PATTERN_LIGHT_TRELLIS);
  MAP_LXW_FMT_CONST(PATTERN_GRAY_125);
  MAP_LXW_FMT_CONST(PATTERN_GRAY_0625);

  MAP_LXW_FMT_CONST(BORDER_THIN);
  MAP_LXW_FMT_CONST(BORDER_MEDIUM);
  MAP_LXW_FMT_CONST(BORDER_DASHED);
  MAP_LXW_FMT_CONST(BORDER_DOTTED);
  MAP_LXW_FMT_CONST(BORDER_THICK);
  MAP_LXW_FMT_CONST(BORDER_DOUBLE);
  MAP_LXW_FMT_CONST(BORDER_HAIR);
  MAP_LXW_FMT_CONST(BORDER_MEDIUM_DASHED);
  MAP_LXW_FMT_CONST(BORDER_DASH_DOT);
  MAP_LXW_FMT_CONST(BORDER_MEDIUM_DASH_DOT);
  MAP_LXW_FMT_CONST(BORDER_DASH_DOT_DOT);
  MAP_LXW_FMT_CONST(BORDER_MEDIUM_DASH_DOT_DOT);
  MAP_LXW_FMT_CONST(BORDER_SLANT_DASH_DOT);

  MAP_LXW_FMT_CONST(DIAGONAL_BORDER_UP);
  MAP_LXW_FMT_CONST(DIAGONAL_BORDER_DOWN);
  MAP_LXW_FMT_CONST(DIAGONAL_BORDER_UP_DOWN);
#undef MAP_LXW_FMT_CONST

#define MAP_LXW_WH_CONST(name, val_name) rb_define_const(cWorksheet, #name, INT2NUM(LXW_##val_name))
  MAP_LXW_WH_CONST(DEF_COL_WIDTH, DEF_COL_WIDTH);
  MAP_LXW_WH_CONST(DEF_ROW_HEIGHT, DEF_ROW_HEIGHT);

  MAP_LXW_WH_CONST(GRIDLINES_HIDE_ALL, HIDE_ALL_GRIDLINES);
  MAP_LXW_WH_CONST(GRIDLINES_SHOW_SCREEN, SHOW_SCREEN_GRIDLINES);
  MAP_LXW_WH_CONST(GRIDLINES_SHOW_PRINT, SHOW_PRINT_GRIDLINES);
  MAP_LXW_WH_CONST(GRIDLINES_SHOW_ALL, SHOW_ALL_GRIDLINES);
#undef MAP_LXW_WH_CONST
}
