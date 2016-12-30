#include <ruby.h>
#include "xlsxwriter.h"

#ifndef __WORKSHEET__
#define __WORKSHEET__

struct worksheet {
  lxw_worksheet *worksheet;
};

VALUE worksheet_alloc(VALUE klass);
VALUE worksheet_init(int argc, VALUE *argv, VALUE self);
VALUE worksheet_release(VALUE self);
void worksheet_free(void *);
VALUE worksheet_write_string_(VALUE self, VALUE row, VALUE col, VALUE value, VALUE format);
VALUE worksheet_write_number_(VALUE self, VALUE row, VALUE col, VALUE value, VALUE format);
VALUE worksheet_write_formula_(VALUE self, VALUE row, VALUE col, VALUE value, VALUE format);
VALUE worksheet_write_array_formula_(VALUE self, VALUE row_from, VALUE col_from, VALUE row_to, VALUE col_to, VALUE value, VALUE format);
VALUE worksheet_write_datetime_(VALUE self, VALUE row, VALUE col, VALUE value, VALUE format);
VALUE worksheet_write_url_(int argc, VALUE *argv, VALUE self);
VALUE worksheet_write_boolean_(VALUE self, VALUE row, VALUE col, VALUE value, VALUE format);
VALUE worksheet_write_blank_(VALUE self, VALUE row, VALUE col, VALUE format);
VALUE worksheet_write_formula_num_(VALUE self, VALUE row, VALUE col, VALUE formula, VALUE format, VALUE value);
VALUE worksheet_set_row_(VALUE self, VALUE row, VALUE opts);
VALUE worksheet_set_column_(VALUE self, VALUE col_from, VALUE col_to, VALUE opts);
VALUE worksheet_insert_image_(VALUE self, VALUE row, VALUE col, VALUE fname, VALUE opts);
VALUE worksheet_insert_chart_(VALUE self, VALUE row, VALUE col, VALUE chart, VALUE opts);
VALUE worksheet_merge_range_(VALUE self, VALUE row_from, VALUE col_from,
                             VALUE row_to, VALUE col_to, VALUE value, VALUE format);
VALUE worksheet_autofilter_(VALUE self, VALUE row_from, VALUE col_from,
                            VALUE row_to, VALUE col_to);
VALUE worksheet_activate_(VALUE self);
VALUE worksheet_select_(VALUE self);
VALUE worksheet_hide_(VALUE self);
VALUE worksheet_set_first_sheet_(VALUE self);
VALUE worksheet_freeze_panes_(VALUE self, VALUE row, VALUE col);
VALUE worksheet_split_panes_(VALUE self, VALUE vertical, VALUE horizontal);
VALUE worksheet_set_selection_(VALUE self, VALUE row_from, VALUE col_from,
                               VALUE row_to, VALUE col_to);
VALUE worksheet_set_landscape_(VALUE self);
VALUE worksheet_set_portrait_(VALUE self);
VALUE worksheet_set_page_view_(VALUE self);
VALUE worksheet_set_paper_(VALUE self, VALUE paper_type);
VALUE worksheet_set_margins_(VALUE self, VALUE left, VALUE right, VALUE top, VALUE bottom);
VALUE worksheet_set_header_(VALUE self, VALUE val, VALUE opts);
VALUE worksheet_set_footer_(VALUE self, VALUE val, VALUE opts);
VALUE worksheet_set_h_pagebreaks_(VALUE self, VALUE val);
VALUE worksheet_set_v_pagebreaks_(VALUE self, VALUE val);
VALUE worksheet_print_across_(VALUE self);
VALUE worksheet_set_zoom_(VALUE self, VALUE val);
VALUE worksheet_gridlines_(VALUE self, VALUE value);
VALUE worksheet_center_horizontally_(VALUE self);
VALUE worksheet_center_vertically_(VALUE self);
VALUE worksheet_print_row_col_headers_(VALUE self);
VALUE worksheet_repeat_rows_(VALUE self, VALUE row_from, VALUE row_to);
VALUE worksheet_repeat_columns_(VALUE self, VALUE col_from, VALUE col_to);
VALUE worksheet_print_area_(VALUE self, VALUE row_from, VALUE col_FROM, VALUE row_to, VALUE col_to);
VALUE worksheet_fit_to_pages_(VALUE self, VALUE width, VALUE height);
VALUE worksheet_set_start_page_(VALUE self, VALUE start_page);
VALUE worksheet_set_print_scale_(VALUE self, VALUE scale);
VALUE worksheet_right_to_left_(VALUE self);
VALUE worksheet_hide_zero_(VALUE self);
VALUE worksheet_set_tab_color_(VALUE self, VALUE color);
VALUE worksheet_protect_(int argc, VALUE *argv, VALUE self);
VALUE worksheet_set_default_row_(VALUE self, VALUE height, VALUE hide_unused_rows);

VALUE worksheet_get_vertical_dpi_(VALUE self);
VALUE worksheet_set_vertical_dpi_(VALUE self, VALUE val);


lxw_col_t value_to_col(VALUE value);

#endif /// __WORKSHEET__
