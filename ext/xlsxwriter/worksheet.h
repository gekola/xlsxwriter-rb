#include <ruby.h>
#include "xlsxwriter.h"

#ifndef __WORKSHEET__
#define __WORKSHEET__

struct worksheet {
  lxw_worksheet *worksheet;
};

lxw_col_t value_to_col(VALUE value);
int extract_cell(int argc, VALUE *argv, lxw_row_t *row, lxw_col_t *col);
int extract_range(int argc, VALUE *argv, lxw_row_t *row1, lxw_col_t *col1,
                                         lxw_row_t *row2, lxw_col_t *col2);
lxw_image_options val_to_lxw_image_options(VALUE opts, char *with_options);

void init_xlsxwriter_worksheet();

extern VALUE cWorksheet;

#endif /// __WORKSHEET__
