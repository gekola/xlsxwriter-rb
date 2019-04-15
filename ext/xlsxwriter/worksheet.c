#include "chart.h"
#include "worksheet.h"
#include "workbook.h"

VALUE cWorksheet;

void worksheet_free(void *p);


VALUE
worksheet_alloc(VALUE klass)
{
  VALUE obj;
  struct worksheet *ptr;

  obj = Data_Make_Struct(klass, struct worksheet, NULL, worksheet_free, ptr);

  ptr->worksheet = NULL;

  return obj;
}

/* :nodoc: */
VALUE
worksheet_init(int argc, VALUE *argv, VALUE self) {
  char *name = NULL;
  VALUE opts = Qnil;
  VALUE auto_width = Qtrue;
  struct workbook *wb_ptr;
  struct worksheet *ptr;

  Data_Get_Struct(self, struct worksheet, ptr);

  if (argc > 2 || argc < 1) {
    rb_raise(rb_eArgError, "wrong number of arguments");
  } else if (argc == 2) {
    switch (TYPE(argv[1])) {
    case T_HASH:
      opts = argv[1];
      break;
    case T_STRING:
    case T_SYMBOL:
      name = StringValueCStr(argv[1]);
      break;
    case T_NIL:
      break;
    default:
      rb_raise(rb_eArgError, "wrong type of name");
      break;
    }
  }

  if (!NIL_P(opts)) {
    VALUE val = rb_hash_aref(opts, ID2SYM(rb_intern("auto_width")));
    if (val == Qfalse) {
      auto_width = Qfalse;
    }
    val = rb_hash_aref(opts, ID2SYM(rb_intern("name")));
    if (!NIL_P(val) && !name) {
      name = StringValueCStr(val);
    }
  }

  rb_iv_set(self, "@workbook", argv[0]);
  rb_iv_set(self, "@use_auto_width", auto_width);
  rb_iv_set(self, "@col_auto_widths", rb_ary_new());

  Data_Get_Struct(argv[0], struct workbook, wb_ptr);
  ptr->worksheet = workbook_add_worksheet(wb_ptr->workbook, name);
  if (!ptr->worksheet)
    rb_raise(rb_eRuntimeError, "worksheet was not created");
  return self;
}

/* :nodoc: */
VALUE
worksheet_release(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);

  worksheet_free(ptr);
  return self;
}

void
worksheet_free(void *p) {
  struct worksheet *ptr = p;

  if (ptr->worksheet) {
    ptr->worksheet = NULL;
  }
}

/*  call-seq:
 *     ws.write_string(cell, value, format = nil) -> self
 *     ws.write_string(row, col, value, format = nil) -> self
 *
 *  Writes a string +value+ into a +cell+ with +format+.
 */
VALUE
worksheet_write_string_(int argc, VALUE *argv, VALUE self) {
  lxw_row_t row;
  lxw_col_t col;
  VALUE value = Qnil;
  VALUE format_key = Qnil;

  rb_check_arity(argc, 2, 4);
  int larg = extract_cell(argc, argv, &row, &col);

  if (larg < argc) {
    value = argv[larg];
    ++larg;
  }

  if (larg < argc) {
    format_key = argv[larg];
    ++larg;
  }

  const char *str = StringValueCStr(value);
  struct worksheet *ptr;
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_write_string(ptr->worksheet, row, col, str, format);
  return self;
}

/*  call-seq:
 *     ws.number_string(cell, value, format = nil) -> self
 *     ws.number_string(row, col, value, format = nil) -> self
 *
 *  Writes a number +value+ into a +cell+ with +format+.
 */
VALUE
worksheet_write_number_(int argc, VALUE *argv, VALUE self) {
  lxw_row_t row;
  lxw_col_t col;
  VALUE value = Qnil;
  VALUE format_key = Qnil;

  rb_check_arity(argc, 2, 4);
  int larg = extract_cell(argc, argv, &row, &col);

  if (larg < argc) {
    value = argv[larg];
    ++larg;
  }

  if (larg < argc) {
    format_key = argv[larg];
    ++larg;
  }

  const double num = NUM2DBL(value);
  struct worksheet *ptr;
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_write_number(ptr->worksheet, row, col, num, format);
  return self;
}

/*  call-seq:
 *     ws.write_formula(cell, formula, format = nil) -> self
 *     ws.write_formula(row, col, formula, format = nil) -> self
 *
 *  Writes a +formula+ into a +cell+ with +format+.
 */
VALUE
worksheet_write_formula_(int argc, VALUE *argv, VALUE self) {
  lxw_row_t row;
  lxw_col_t col;
  VALUE value = Qnil;
  VALUE format_key = Qnil;

  rb_check_arity(argc, 2, 4);
  int larg = extract_cell(argc, argv, &row, &col);

  if (larg < argc) {
    value = argv[larg];
    ++larg;
  }

  if (larg < argc) {
    format_key = argv[larg];
    ++larg;
  }

  const char *str = RSTRING_PTR(value);
  struct worksheet *ptr;
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_write_formula(ptr->worksheet, row, col, str, format);
  return self;
}

/*  call-seq:
 *     ws.write_array_formula(range, formula, format = nil) -> self
 *     ws.write_array_formula(cell_from, cell_to, formula, format = nil) -> self
 *     ws.write_array_formula(row_from, col_from, row_to, col_to, formula, format = nil) -> self
 *
 *  Writes an array +formula+ into a cell +range+ with +format+.
 */
VALUE worksheet_write_array_formula_(int argc, VALUE *argv, VALUE self) {
  lxw_row_t row_from, row_to;
  lxw_col_t col_from, col_to;
  const char *str = NULL;
  VALUE format_key = Qnil;

  rb_check_arity(argc, 2, 6);
  int larg = extract_range(argc, argv, &row_from, &col_from, &row_to, &col_to);

  if (larg < argc) {
    str = StringValueCStr(argv[larg]);
    ++larg;
  } else {
    rb_raise(rb_eArgError, "No formula specified");
  }

  if (larg < argc) {
    format_key = argv[larg];
    ++larg;
  }

  struct worksheet *ptr;
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_write_array_formula(ptr->worksheet, row_from, col_from, row_to, col_to, str, format);
  return self;
}

/*  call-seq:
 *     ws.write_datetime(cell, value, format = nil) -> self
 *     ws.write_datetime(row, col, value, format = nil) -> self
 *
 *  Writes a datetime +value+ into a +cell+ with +format+.
 */
VALUE
worksheet_write_datetime_(int argc, VALUE *argv, VALUE self) {
  lxw_row_t row;
  lxw_col_t col;
  VALUE value = Qnil;
  VALUE format_key = Qnil;

  rb_check_arity(argc, 2, 4);
  int larg = extract_cell(argc, argv, &row, &col);

  if (larg < argc) {
    value = argv[larg];
    ++larg;
  }

  if (larg < argc) {
    format_key = argv[larg];
    ++larg;
  }

  struct lxw_datetime datetime = value_to_lxw_datetime(value);
  struct worksheet *ptr;
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_write_datetime(ptr->worksheet, row, col, &datetime, format);
  return self;
}

/*  call-seq:
 *     ws.write_url(cell, url, format = nil, string: nil, tooltip: nil, format: nil) -> self
 *     ws.write_url(row, col, url, format = nil, string: nil, tooltip: nil, format: nil) -> self
 *
 *  Writes a +url+ into a +cell+ with +format+.
 */
VALUE
worksheet_write_url_(int argc, VALUE *argv, VALUE self) {
  lxw_row_t row;
  lxw_col_t col;
  VALUE url = Qnil; //argv[2];
  VALUE format_key = Qnil; // argv[3];
  VALUE opts = Qnil;

  rb_check_arity(argc, 2, 5);
  int larg = extract_cell(argc, argv, &row, &col);

  if (larg < argc) {
    url = argv[larg];
    ++larg;
  }

  const char *url_str = RSTRING_PTR(url);
  char *string = NULL;
  char *tooltip = NULL;
  while (larg < argc) {
    switch(TYPE(argv[larg])) {
    case T_HASH:
      opts = argv[larg];
      break;
    case T_SYMBOL: case T_STRING: case T_NIL:
      format_key = argv[larg];
      break;
    default:
      rb_raise(rb_eTypeError, "Expected Hash, symbol or string but got %"PRIsVALUE, rb_obj_class(argv[larg]));
    }
    ++larg;
  }

  if (!NIL_P(opts)) {
    VALUE val;
    val = rb_hash_aref(opts, ID2SYM(rb_intern("string")));
    if (!NIL_P(val)) {
      string = StringValueCStr(val);
    }
    val = rb_hash_aref(opts, ID2SYM(rb_intern("tooltip")));
    if (!NIL_P(val)) {
      tooltip = StringValueCStr(val);
    }
    if (NIL_P(format_key)) {
      format_key = rb_hash_aref(opts, ID2SYM(rb_intern("format")));
    }
  }
  struct worksheet *ptr;
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  Data_Get_Struct(self, struct worksheet, ptr);
  if (string || tooltip) {
    worksheet_write_url_opt(ptr->worksheet, row, col, url_str, format, string, tooltip);
  } else {
    worksheet_write_url(ptr->worksheet, row, col, url_str, format);
  }
  return self;
}

/*  call-seq:
 *     ws.write_boolean(cell, value, format = nil) -> self
 *     ws.write_boolean(row, col, value, format = nil) -> self
 *
 *  Writes a boolean +value+ into a +cell+ with +format+.
 */
VALUE
worksheet_write_boolean_(int argc, VALUE *argv, VALUE self) {
  lxw_row_t row;
  lxw_col_t col;
  int bool_value = 0;
  VALUE format_key = Qnil;

  rb_check_arity(argc, 2, 4);
  int larg = extract_cell(argc, argv, &row, &col);

  if (larg < argc) {
    bool_value = argv[larg] && !NIL_P(argv[larg]);
    ++larg;
  }

  if (larg < argc) {
    format_key = argv[larg];
    ++larg;
  }

  struct worksheet *ptr;
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_write_boolean(ptr->worksheet, row, col, bool_value, format);
  return self;
}

/*  call-seq:
 *     ws.write_blank(cell, format = nil) -> self
 *     ws.write_blank(row, col, format = nil) -> self
 *
 *  Make a +cell+ blank with +format+.
 */
VALUE
worksheet_write_blank_(int argc, VALUE *argv, VALUE self) {
  lxw_row_t row;
  lxw_col_t col;
  VALUE format_key = Qnil;

  rb_check_arity(argc, 1, 3);
  int larg = extract_cell(argc, argv, &row, &col);

  if (larg < argc) {
    format_key = argv[larg];
    ++larg;
  }

  struct worksheet *ptr;
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_write_blank(ptr->worksheet, row, col, format);
  return self;
}

/*  call-seq:
 *     ws.write_formula_num(cell, formula, value, format = nil) -> self
 *     ws.write_formula_num(row, col, formula, value, format = nil) -> self
 *
 *  Writes a +formula+ with +value+ into a +cell+ with +format+.
 */
VALUE worksheet_write_formula_num_(int argc, VALUE *argv, VALUE self) {
  lxw_row_t row;
  lxw_col_t col;
  VALUE formula = Qnil;
  VALUE value = Qnil;
  VALUE format_key = Qnil;

  rb_check_arity(argc, 3, 5);
  int larg = extract_cell(argc, argv, &row, &col);

  if (larg < argc) {
    formula = argv[larg];
    ++larg;
  }

  if (larg < argc) {
    value = argv[larg];
    ++larg;
  }

  if (larg < argc) {
    format_key = argv[larg];
    ++larg;
  }

  const char *str = RSTRING_PTR(formula);
  struct worksheet *ptr;
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_write_formula_num(ptr->worksheet, row, col, str, format, NUM2DBL(value));
  return self;
}

/*  call-seq: ws.set_row(row, height: nil, format: nil, hide: false) -> self
 *
 *  Set properties of the row.
 */
VALUE
worksheet_set_row_(VALUE self, VALUE row, VALUE opts) {
  double height = LXW_DEF_ROW_HEIGHT;
  lxw_format *format = NULL;
  lxw_row_col_options options = {
    .hidden = 0,
    .collapsed = 0,
    .level = 0
  };
  VALUE val;
  VALUE workbook = rb_iv_get(self, "@workbook");

  val = rb_hash_aref(opts, ID2SYM(rb_intern("height")));
  if (val != Qnil)
    height = NUM2DBL(val);

  val = rb_hash_aref(opts, ID2SYM(rb_intern("format")));
  format = workbook_get_format(workbook, val);

  val = rb_hash_aref(opts, ID2SYM(rb_intern("hide")));
  if (val != Qnil && val)
    options.hidden = 1;

  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  if (options.hidden) {
    worksheet_set_row_opt(ptr->worksheet, NUM2INT(row), height, format, &options);
  } else {
    worksheet_set_row(ptr->worksheet, NUM2INT(row), height, format);
  }
  return self;
}

/*  call-seq: ws.set_column(col_from, col_to, width: nil, format: nil, hide: nil) -> self
 *
 *  Set properties of the cols range.
 */
VALUE
worksheet_set_column_(VALUE self, VALUE col_from, VALUE col_to, VALUE opts) {
  double width = LXW_DEF_COL_WIDTH;
  lxw_format *format = NULL;
  lxw_row_col_options options = {
    .hidden = 0,
    .collapsed = 0,
    .level = 0
  };
  VALUE val;
  lxw_col_t col1 = value_to_col(col_from);
  lxw_col_t col2 = value_to_col(col_to);
  VALUE workbook = rb_iv_get(self, "@workbook");

  val = rb_hash_aref(opts, ID2SYM(rb_intern("width")));
  if (val != Qnil)
    width = NUM2DBL(val);

  val = rb_hash_aref(opts, ID2SYM(rb_intern("format")));
  format = workbook_get_format(workbook, val);

  val = rb_hash_aref(opts, ID2SYM(rb_intern("hide")));
  if (val != Qnil && val)
    options.hidden = 1;

  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  if (options.hidden) {
    worksheet_set_column_opt(ptr->worksheet, col1, col2, width, format, &options);
  } else {
    worksheet_set_column(ptr->worksheet, col1, col2, width, format);
  }
  return self;
}

/*  call-seq:
 *     ws.insert_image(cell, fname, opts = {}) -> self
 *     ws.insert_image(row, col, fname, opts = {}) -> self
 *
 *  Inserts image from +fname+ into the worksheet (at +cell+).
 *
 *  Available +opts+
 *  :offset, :x_offset, :y_offset:: Image offset (+:offset+ for both; Integer).
 *  :scale, :x_scale, :y_scale:: Image scale (+:scale+ for both; Numeric).
 */
VALUE
worksheet_insert_image_(int argc, VALUE *argv, VALUE self) {
  lxw_row_t row;
  lxw_col_t col;
  VALUE fname = Qnil;
  VALUE opts = Qnil;
  lxw_image_options options;
  char with_options = '\0';

  rb_check_arity(argc, 2, 4);
  int larg = extract_cell(argc, argv, &row, &col);

  if (larg < argc) {
    fname = argv[larg];
    ++larg;
  }

  if (larg < argc) {
    opts = argv[larg];
    ++larg;
  }

  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);

  if (!NIL_P(opts)) {
    options = val_to_lxw_image_options(opts, &with_options);
  }

  if (with_options) {
    worksheet_insert_image_opt(ptr->worksheet, row, col, StringValueCStr(fname), &options);
  } else {
    worksheet_insert_image(ptr->worksheet, row, col, StringValueCStr(fname));
  }

  return self;
}

/*  call-seq:
 *     ws.insert_chart(cell, chart, opts = {}) -> self
 *     ws.insert_chart(row, col, chart, opts = {}) -> self
 *
 *  Inserts chart from +fname+ into the worksheet (at +cell+).
 *
 *  For available +opts+ see #insert_image.
 */
VALUE
worksheet_insert_chart_(int argc, VALUE *argv, VALUE self) {
  lxw_row_t row;
  lxw_col_t col;
  VALUE chart, opts = Qnil;
  lxw_image_options options;
  char with_options = '\0';

  rb_check_arity(argc, 2, 4);
  int larg = extract_cell(argc, argv, &row, &col);

  if (larg < argc) {
    chart = argv[larg];
    ++larg;
  } else {
    rb_raise(rb_eArgError, "No chart specified");
  }

  if (larg < argc) {
    opts = argv[larg];
    ++larg;
  }

  if (!NIL_P(opts)) {
    options = val_to_lxw_image_options(opts, &with_options);
  }

  struct worksheet *ptr;
  struct chart *chart_ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  Data_Get_Struct(chart, struct chart, chart_ptr);

  if (with_options) {
    worksheet_insert_chart_opt(ptr->worksheet, row, col, chart_ptr->chart, &options);
  } else {
    worksheet_insert_chart(ptr->worksheet, row, col, chart_ptr->chart);
  }

  return self;
}

/*  call-seq:
 *     ws.merge_range(range, value = "", format = nil) -> self
 *     ws.merge_range(cell_from, cell_to, value = "", format = nil) -> self
 *     ws.merge_range(row_from, col_from, row_to, col_to, value = "", format = nil) -> self
 *
 *  Merges +range+, setting string +value+ with +format+.
 *
 *     ws.merge_range('A1:D5')
 *     ws.merge_range('A1', 'D5', 'Merged cells')
 *     ws.merge_range(0, 0, 4, 3)
 */
VALUE
worksheet_merge_range_(int argc, VALUE *argv, VALUE self) {
  char *str = "";
  lxw_format *format = NULL;
  lxw_col_t col1, col2;
  lxw_row_t row1, row2;
  VALUE workbook = rb_iv_get(self, "@workbook");

  rb_check_arity(argc, 1, 6);
  int larg = extract_range(argc, argv, &row1, &col1, &row2, &col2);

  if (larg < argc) {
    str = StringValueCStr(argv[larg]);
    ++larg;
  }

  if (larg < argc) {
    format = workbook_get_format(workbook, argv[larg]);
    ++larg;
  }

  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);

  worksheet_merge_range(ptr->worksheet, row1, col1, row2, col2, str, format);
  return self;
}

/*  call-seq:
 *     ws.autofilter(range) -> self
 *     ws.autofilter(cell_from, cell_to) -> self
 *     ws.autofilter(row_from, col_from, row_to, col_to) -> self
 *
 *  Applies autofilter to the +range+.
 *
 *     ws.autofilter('A1:D5')
 *     ws.autofilter('A1', 'D5')
 *     ws.autofilter(0, 0, 4, 3)
 */
VALUE
worksheet_autofilter_(int argc, VALUE *argv, VALUE self) {
  lxw_row_t row_from, row_to;
  lxw_col_t col_from, col_to;

  rb_check_arity(argc, 1, 4);
  extract_range(argc, argv, &row_from, &col_from, &row_to, &col_to);

  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);

  worksheet_autofilter(ptr->worksheet, row_from, col_from, row_to, col_to);
  return self;
}

/*  call-seq: ws.activate -> self
 *
 *  Set the worksheet to be active on opening the workbook.
 */
VALUE worksheet_activate_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_activate(ptr->worksheet);
  return self;
}

/*  call-seq: ws.select -> self
 *
 *  Set the worksheet to be selected on openong the workbook.
 */
VALUE worksheet_select_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_select(ptr->worksheet);
  return self;
}

/*  call-seq: ws.hide -> self
 *
 *  Hide the worksheet.
 */
VALUE
worksheet_hide_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_hide(ptr->worksheet);
  return self;
}

/*  call-seq: ws.set_fitst_sheet -> self
 *
 *  Set the sheet to be the first visible in the sheets list (which is placed
 *  under the work area in Excel).
 */
VALUE
worksheet_set_first_sheet_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_first_sheet(ptr->worksheet);
  return self;
}

/*  call-seq:
 *     ws.freeze_panes(cell) -> self
 *     ws.freeze_panes(row, col) -> self
 *
 *  Divides the worksheet into horizontal or vertical panes and freezes them.
 *
 *  The specified +cell+ is the top left in the right bottom pane.
 *
 *  In order to freeze only rows/cols pass 0 as +row+ or +col+.
 *
 *  Advanced usage (with 2nd cell and type) is not documented yet.
 */
VALUE
worksheet_freeze_panes_(int argc, VALUE *argv, VALUE self) {
  lxw_row_t row, row2;
  lxw_col_t col, col2;
  uint8_t type = 0;
  rb_check_arity(argc, 1, 5);
  int larg = extract_cell(argc, argv, &row, &col);
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  if (larg >= argc) {
    worksheet_freeze_panes(ptr->worksheet, row, col);
  } else {
    larg += extract_cell(argc - larg, argv + larg, &row2, &col2);
    if (larg < argc) {
      type = NUM2INT(argv[larg]);
    }
    worksheet_freeze_panes_opt(ptr->worksheet, row, col, row2, col2, type);
  }
  return self;
}

/*  call-seq: ws.split_panes(vertical, horizontal) -> self
 *
 *  Divides the worksheet int vertical and horizontal panes with respective
 *  positions from parameters (Numeric).
 *
 *  If only one split is desired pass 0 for other.
 */
VALUE
worksheet_split_panes_(VALUE self, VALUE vertical, VALUE horizontal) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_split_panes(ptr->worksheet, NUM2DBL(vertical), NUM2DBL(horizontal));
  return self;
}

/*  call-seq:
 *     ws.set_selection(range) -> self
 *     ws.set_selection(cell_from, cell_to) -> self
 *     ws.set_selection(row_from, col_from, row_to, col_to) -> self
 *
 *  Selects the specified +range+ on the worksheet.
 *
 *     ws.set_selection('A1:G5')
 *     ws.set_selection('A1', 'G5')
 *     ws.set_selection(0, 0, 4, 6)
 */
VALUE
worksheet_set_selection_(int argc, VALUE *argv, VALUE self) {
  lxw_row_t row_from, row_to;
  lxw_col_t col_from, col_to;

  rb_check_arity(argc, 1, 4);
  extract_range(argc, argv, &row_from, &col_from, &row_to, &col_to);

  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);

  worksheet_set_selection(ptr->worksheet, row_from, col_from, row_to, col_to);
  return self;
}

/*  call-seq: ws.set_landscape -> self
 *
 *  Sets print orientation of the worksheet to landscape.
 */
VALUE
worksheet_set_landscape_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_landscape(ptr->worksheet);
  return self;
}

/*  call-seq: ws.set_portrait -> self
 *
 *  Sets print orientation of the worksheet to portrait.
 */
VALUE
worksheet_set_portrait_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_portrait(ptr->worksheet);
  return self;
}

/*  call-seq: ws.set_page_view -> self
 *
 *  Sets worksheet displa mode to "Page View/Layout".
 */
VALUE
worksheet_set_page_view_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_page_view(ptr->worksheet);
  return self;
}

/*  call-seq: ws.paper=(type) -> type
 *
 *  Sets the paper +type+ for printing. See {libxlsxwriter doc}[https://libxlsxwriter.github.io/worksheet_8h.html#a9f8af12017797b10c5ee68e04149032f] for options.
 */
VALUE
worksheet_set_paper_(VALUE self, VALUE paper_type) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_paper(ptr->worksheet, NUM2INT(paper_type));
  return self;
}

/* Sets the worksheet margins (Numeric) for the printed page. */
VALUE
worksheet_set_margins_(VALUE self, VALUE left, VALUE right, VALUE top, VALUE bottom) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_margins(ptr->worksheet, NUM2DBL(left), NUM2DBL(right), NUM2DBL(top), NUM2DBL(bottom));
  return self;
}

/*  call-seq: ws.set_header(text, opts) -> self
 *
 *  See {libxlsxwriter doc}[https://libxlsxwriter.github.io/worksheet_8h.html#a4070c24ed5107f33e94f30a1bf865ba9]
 *  for the +text+ control characters.
 *
 *  Currently the only supported option is +:margin+ (Numeric).
 */
VALUE
worksheet_set_header_(VALUE self, VALUE val, VALUE opts) {
  const char *str = StringValueCStr(val);
  lxw_header_footer_options options = { .margin = 0.0 };
  char with_options = '\0';
  if (!NIL_P(opts)) {
    VALUE margin = rb_hash_aref(opts, ID2SYM(rb_intern("margin")));
    if (!NIL_P(margin)) {
      with_options = '\1';
      options.margin = NUM2DBL(margin);
    }
  }
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  if (with_options) {
    worksheet_set_header(ptr->worksheet, str);
  } else {
    worksheet_set_header_opt(ptr->worksheet, str, &options);
  }
  return self;
}

/*  call-seq: ws.set_footer(text, opts)
 *
 *  See #set_header for params description.
 */
VALUE
worksheet_set_footer_(VALUE self, VALUE val, VALUE opts) {
  const char *str = StringValueCStr(val);
  lxw_header_footer_options options = { .margin = 0.0 };
  char with_options = '\0';
  if (!NIL_P(opts)) {
    VALUE margin = rb_hash_aref(opts, ID2SYM(rb_intern("margin")));
    if (!NIL_P(margin)) {
      with_options = '\1';
      options.margin = NUM2DBL(margin);
    }
  }
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  if (with_options) {
    worksheet_set_footer(ptr->worksheet, str);
  } else {
    worksheet_set_footer_opt(ptr->worksheet, str, &options);
  }
  return self;
}

/*  call-seq: wb.h_pagebreaks=(breaks) -> breaks
 *
 *  Adds horizontal page +breaks+ to the worksheet.
 *
 *     wb.h_pagebreaks = [20, 40, 80]
 */
VALUE
worksheet_set_h_pagebreaks_(VALUE self, VALUE val) {
  const size_t len = RARRAY_LEN(val);
  lxw_row_t rows[len+1];
  for (size_t i = 0; i<len; ++i) {
    rows[i] = NUM2INT(rb_ary_entry(val, i));
  }
  rows[len] = 0;
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_h_pagebreaks(ptr->worksheet, rows);
  return val;
}

/*  call-seq: wb.h_pagebreaks=(breaks) -> breaks
 *
 *  Adds vertical page +breaks+ to the worksheet.
 *
 *     wb.v_pagebreaks = [20, 40, 80]
 */
VALUE
worksheet_set_v_pagebreaks_(VALUE self, VALUE val) {
  const size_t len = RARRAY_LEN(val);
  lxw_col_t cols[len+1];
  for (size_t i = 0; i<len; ++i) {
    cols[i] = value_to_col(rb_ary_entry(val, i));
  }
  cols[len] = 0;
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_v_pagebreaks(ptr->worksheet, cols);
  return val;
}

/* Changes the default print direction */
VALUE
worksheet_print_across_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_print_across(ptr->worksheet);
  return self;
}

/*  call-seq: ws.zoom=(zoom) -> zoom
 *
 *  Sets the worksheet zoom factor in the range 10 <= +zoom+ <= 400.
 */
VALUE
worksheet_set_zoom_(VALUE self, VALUE val) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_zoom(ptr->worksheet, NUM2INT(val));
  return self;
}

/*  call-seq: ws.gridlines=(option) -> option
 *
 *  Display or hide screen and print gridlines using one of the values
 *  XlsxWriter::Worksheet::{GRIDLINES_HIDE_ALL, GRIDLINES_SHOW_SCREEN,
 *  GRIDLINES_SHOW_PRINT, GRIDLINES_SHOW_ALL}.
 */
VALUE
worksheet_gridlines_(VALUE self, VALUE value) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);

  worksheet_gridlines(ptr->worksheet, NUM2INT(value));

  return value;
}

/* Center the worksheet data horizontally between the margins on the printed page */
VALUE
worksheet_center_horizontally_(VALUE self){
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_center_horizontally(ptr->worksheet);
  return self;
}

/* Center the worksheet data vertically between the margins on the printed page */
VALUE
worksheet_center_vertically_(VALUE self) {
    struct worksheet *ptr;
    Data_Get_Struct(self, struct worksheet, ptr);
    worksheet_center_vertically(ptr->worksheet);
    return self;
}

/* Print rows and column header (wich is disabled by default). */
VALUE
worksheet_print_row_col_headers_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_print_row_col_headers(ptr->worksheet);
  return self;
}

/*  call-seq: ws.repeat_rows(row_from, row_to)
 *
 *  Sets rows to be repeatedly printed out on all pages.
 */
VALUE
worksheet_repeat_rows_(VALUE self, VALUE row_from, VALUE row_to) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_repeat_rows(ptr->worksheet, NUM2INT(row_from), NUM2INT(row_to));
  return self;
}

/*  call-seq: ws.repeat_columns(col_from, col_to)
 *
 *  Sets columns to be repeatedly printed out on all pages.
 */
VALUE
worksheet_repeat_columns_(VALUE self, VALUE col_from, VALUE col_to) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_repeat_columns(ptr->worksheet, value_to_col(col_from), value_to_col(col_to));
  return self;
}

/*  call-seq:
 *     ws.print_area(range)
 *     ws.print_area(cell_from, cell_to)
 *     ws.print_area(row_from, col_from, row_to, col_to)
 *
 *  Specifies area of the worksheet to be printed.
 */
VALUE
worksheet_print_area_(int argc, VALUE *argv, VALUE self) {
  lxw_row_t row_from, row_to;
  lxw_col_t col_from, col_to;

  rb_check_arity(argc, 1, 4);
  extract_range(argc, argv, &row_from, &col_from, &row_to, &col_to);

  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);

  worksheet_print_area(ptr->worksheet, row_from, col_from, row_to, col_to);
  return self;
}

/*  call-seq: ws.fit_to_pages(width, height) -> self
 *
 *  Fits the printed area to a specific number of pages both vertically and
 *  horizontally.
 */
VALUE
worksheet_fit_to_pages_(VALUE self, VALUE width, VALUE height) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_fit_to_pages(ptr->worksheet, NUM2INT(width), NUM2INT(height));
  return self;
}

/*  call-seq: ws.start_page=(page) -> page
 *
 *  Set the number of the first printed page.
 */
VALUE
worksheet_set_start_page_(VALUE self, VALUE start_page) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_start_page(ptr->worksheet, NUM2INT(start_page));
  return start_page;
}

/*  call-seq: ws.print_scale=(scale) -> scale
 *
 *  Sets the +scale+ factor of the printed page (10 <= +scale+ <= 400).
 */
VALUE
worksheet_set_print_scale_(VALUE self, VALUE scale) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_print_scale(ptr->worksheet, NUM2INT(scale));
  return scale;
}

/* Sets text direction to rtl (e.g. for worksheets on Hebrew or Arabic). */
VALUE
worksheet_right_to_left_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_right_to_left(ptr->worksheet);
  return self;
}

/* Hides all zero values. */
VALUE
worksheet_hide_zero_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_hide_zero(ptr->worksheet);
  return self;
}

/*  call-seq: ws.tab_color=(color) -> color
 *
 *  Set the tab color (from RGB integer).
 *
 *     ws.tab_color = 0xF0F0F0
 */
VALUE
worksheet_set_tab_color_(VALUE self, VALUE color) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_tab_color(ptr->worksheet, NUM2INT(color));
  return color;
}

/*  call-seq: ws.protect(password[, opts]) -> self
 *
 *  Protects the worksheet elements from modification.
 *  See {libxlsxwriter doc}[https://libxlsxwriter.github.io/worksheet_8h.html#a1b49e135d4debcdb25941f2f40f04cb0]
 *  for options.
 */
VALUE
worksheet_protect_(int argc, VALUE *argv, VALUE self) {
  rb_check_arity(argc, 0, 2);
  uint8_t with_options = '\0';
  VALUE val;
  VALUE opts = Qnil;
  // All options are off by default
  lxw_protection options = {};
  const char *password = NULL;
  if (argc > 0 && !NIL_P(argv[0])) {
    switch (TYPE(argv[0])) {
    case T_STRING:
      password = StringValueCStr(argv[0]);
      break;
    case T_HASH:
      opts = argv[0];
      break;
    default:
      rb_raise(rb_eArgError, "Wrong argument %"PRIsVALUE", expected a String or Hash", rb_obj_class(argv[0]));
    }
  }

  if (argc > 1) {
    if (TYPE(argv[1]) == T_HASH) {
      opts = argv[1];
    } else {
      rb_raise(rb_eArgError, "Expected a Hash, but got %"PRIsVALUE, rb_obj_class(argv[1]));
    }
  }

  if (!NIL_P(opts)) {
#define PR_OPT(field) {                                    \
      val = rb_hash_aref(opts, ID2SYM(rb_intern(#field))); \
      if (!NIL_P(val) && val) {                            \
        options.field = 1;                                 \
        with_options = 1;                                  \
      }                                                    \
    }
    PR_OPT(no_select_locked_cells);
    PR_OPT(no_select_unlocked_cells);
    PR_OPT(format_cells);
    PR_OPT(format_columns);
    PR_OPT(format_rows);
    PR_OPT(insert_columns);
    PR_OPT(insert_rows);
    PR_OPT(insert_hyperlinks);
    PR_OPT(delete_columns);
    PR_OPT(delete_rows);
    PR_OPT(sort);
    PR_OPT(autofilter);
    PR_OPT(pivot_tables);
    PR_OPT(scenarios);
    PR_OPT(objects);
#undef PR_OPT
  }
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  if (with_options) {
    worksheet_protect(ptr->worksheet, password, &options);
  } else {
    worksheet_protect(ptr->worksheet, password, NULL);
  }
  return self;
}

/*  call-seq: ws.set_default_row(height, hide_unuser_rows) -> self
 *
 *  Sets the default row properties for the worksheet.
 */
VALUE
worksheet_set_default_row_(VALUE self, VALUE height, VALUE hide_unused_rows) {
  struct worksheet *ptr;
  uint8_t hide_ur = !NIL_P(hide_unused_rows) && hide_unused_rows != Qfalse ? 1 : 0;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_default_row(ptr->worksheet, NUM2DBL(height), hide_ur);
  return self;
}

/*  call-seq: ws.vertical_dpi -> int
 *
 *  Returns the vertical dpi.
 */
VALUE
worksheet_get_vertical_dpi_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  return INT2NUM(ptr->worksheet->vertical_dpi);
}

/*  call-seq: ws.vertical_dpi=(dpi) -> dpi
 *
 *  Sets the vertical dpi.
 */
VALUE
worksheet_set_vertical_dpi_(VALUE self, VALUE val) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  ptr->worksheet->vertical_dpi = NUM2INT(val);
  return val;
}


lxw_col_t
value_to_col(VALUE value) {
  switch (TYPE(value)) {
  case T_FIXNUM:
    return NUM2INT(value);
  case T_STRING:
    return lxw_name_to_col(RSTRING_PTR(value));
  default:
    rb_raise(rb_eTypeError, "Wrong type for col %"PRIsVALUE, rb_obj_class(value));
    return -1;
  }
}

int
extract_cell(int argc, VALUE *argv, lxw_row_t *row, lxw_col_t *col) {
  char *str;
  switch (TYPE(argv[0])) {
  case T_STRING:
    str = RSTRING_PTR(argv[0]);
    if ((str[0] >= 'A' && str[0] <= 'Z') ||
        (str[0] >= 'a' && str[0] <= 'z')) {
      // Column in also in argv[0]
      (*row) = lxw_name_to_row(str);
      (*col) = lxw_name_to_col(str);
      return 1;
    } else {
      if (argc > 1) {
        (*row) = atoi(str);
        switch(TYPE(argv[1])) {
        case T_STRING:
          str = RSTRING_PTR(argv[1]);
          if ((str[0] >= 'A' && str[0] <= 'Z') ||
              (str[0] >= 'a' && str[0] <= 'z')) {
            (*col) = lxw_name_to_col(str);
          } else {
            (*col) = atoi(str);
          }
          return 2;
        case T_FIXNUM:
          (*col) = NUM2INT(argv[1]);
          return 2;
        default:
          rb_raise(rb_eArgError, "Cannot extract column info from %"PRIsVALUE,
                   rb_inspect(argv[1]));
          return 0;
        }
      } else {
        rb_raise(rb_eArgError, "Cannot extract column info from %"PRIsVALUE,
                 rb_inspect(argv[0]));
        return 0;
      }
    }
    break;
  case T_FIXNUM:
    if (argc > 1) {
      (*row) = NUM2INT(argv[0]);
      (*col) = value_to_col(argv[1]);
      return 2;
    } else {
      rb_raise(rb_eArgError, "Cannot extract column, not enough arguments");
      return 0;
    }
    break;
  default:
    rb_raise(rb_eTypeError, "Expected string or number, got %"PRIsVALUE, rb_obj_class(argv[0]));
    return 0;
  }
}

int extract_range(int argc, VALUE *argv, lxw_row_t *row_from, lxw_col_t *col_from,
                                         lxw_row_t *row_to,   lxw_col_t *col_to) {
  char *str;
  if (TYPE(argv[0]) == T_STRING) {
    str = RSTRING_PTR(argv[0]);
    if (strstr(str, ":")) {
      (*row_from) = lxw_name_to_row(str);
      (*col_from) = lxw_name_to_col(str);
      (*row_to) = lxw_name_to_row_2(str);
      (*col_to) = lxw_name_to_col_2(str);
      return 1;
    }
  }

  int larg = extract_cell(argc, argv, row_from, col_from);
  larg += extract_cell(argc - larg, argv + larg, row_to, col_to);
  return larg;
}

#define SET_IMG_OPT(key, setter) {                    \
    val = rb_hash_aref(opts, ID2SYM(rb_intern(key))); \
    if (!NIL_P(val)) {                                \
      (*with_options) = '\1';                         \
      setter;                                         \
    }                                                 \
  }
lxw_image_options
val_to_lxw_image_options(VALUE opts, char *with_options) {
  VALUE val;
  lxw_image_options options = {
    .x_offset = 0,
    .y_offset = 0,
    .x_scale = 1.0,
    .y_scale = 1.0
  };
  SET_IMG_OPT("offset",   options.x_offset = options.y_offset = NUM2INT(val));
  SET_IMG_OPT("x_offset", options.x_offset =                    NUM2INT(val));
  SET_IMG_OPT("y_offset",                    options.y_offset = NUM2INT(val));
  SET_IMG_OPT("scale",    options.x_scale  = options.y_scale  = NUM2DBL(val));
  SET_IMG_OPT("x_scale",  options.x_scale  =                    NUM2DBL(val));
  SET_IMG_OPT("y_scale",                     options.y_scale  = NUM2DBL(val));
  return options;
}
#undef SET_IMG_OPT


void
init_xlsxwriter_worksheet() {
  /* Xlsx worksheet */
  cWorksheet = rb_define_class_under(mXlsxWriter, "Worksheet", rb_cObject);

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

#define MAP_LXW_WH_CONST(name, val_name) rb_define_const(cWorksheet, #name, INT2NUM(LXW_##val_name))
  MAP_LXW_WH_CONST(DEF_COL_WIDTH, DEF_COL_WIDTH);
  MAP_LXW_WH_CONST(DEF_ROW_HEIGHT, DEF_ROW_HEIGHT);

  MAP_LXW_WH_CONST(GRIDLINES_HIDE_ALL, HIDE_ALL_GRIDLINES);
  MAP_LXW_WH_CONST(GRIDLINES_SHOW_SCREEN, SHOW_SCREEN_GRIDLINES);
  MAP_LXW_WH_CONST(GRIDLINES_SHOW_PRINT, SHOW_PRINT_GRIDLINES);
  MAP_LXW_WH_CONST(GRIDLINES_SHOW_ALL, SHOW_ALL_GRIDLINES);
#undef MAP_LXW_WH_CONST
}
