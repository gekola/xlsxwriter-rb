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
VALUE worksheet_write_url_(VALUE self, VALUE row, VALUE col, VALUE value, VALUE format);
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
// VALUE worksheet_hide_;
// VALUE worksheet_set_first_sheet_;

VALUE worksheet_gridlines_(VALUE self, VALUE value);


lxw_col_t value_to_col(VALUE value);

#endif /// __WORKSHEET__
