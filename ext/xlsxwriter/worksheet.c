#include "chart.h"
#include "common.h"
#include "rich_string.h"
#include "worksheet.h"
#include "workbook.h"
#include "xlsxwriter_ext.h"

VALUE cWorksheet;

void worksheet_clear(void *p);
void worksheet_free(void *p);

size_t worksheet_size(const void *data) {
  return sizeof(struct worksheet);
}

const rb_data_type_t worksheet_type = {
    .wrap_struct_name = "worksheet",
    .function =
        {
            .dmark = NULL,
            .dfree = worksheet_free,
            .dsize = worksheet_size,
        },
    .data = NULL,
    .flags = RUBY_TYPED_FREE_IMMEDIATELY,
};


VALUE
worksheet_alloc(VALUE klass) {
  VALUE obj;
  struct worksheet *ptr;

  obj = TypedData_Make_Struct(klass, struct worksheet, &worksheet_type, ptr);

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

  TypedData_Get_Struct(self, struct worksheet, &worksheet_type, ptr);

  rb_check_arity(argc, 1, 2);
  if (argc == 2) {
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

  TypedData_Get_Struct(argv[0], struct workbook, &workbook_type, wb_ptr);
  ptr->worksheet = workbook_add_worksheet(wb_ptr->workbook, name);
  if (!ptr->worksheet)
    rb_raise(eXlsxWriterError, "worksheet was not created");
  return self;
}

/* :nodoc: */
VALUE
worksheet_release(VALUE self) {
  struct worksheet *ptr;
  TypedData_Get_Struct(self, struct worksheet, &worksheet_type, ptr);

  worksheet_clear(ptr);
  return self;
}

void
worksheet_clear(void *p) {
  struct worksheet *ptr = p;

  if (ptr->worksheet) {
    ptr->worksheet = NULL;
  }
}

void
worksheet_free(void *p) {
  worksheet_clear(p);
  ruby_xfree(p);
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
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  LXW_NO_RESULT_CALL(worksheet, write_string, row, col, str, format);
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
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  LXW_NO_RESULT_CALL(worksheet, write_number, row, col, num, format);
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
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  LXW_NO_RESULT_CALL(worksheet, write_formula, row, col, str, format);
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

  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  LXW_NO_RESULT_CALL(worksheet, write_array_formula, row_from, col_from, row_to, col_to, str, format);
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
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  LXW_NO_RESULT_CALL(worksheet, write_datetime, row, col, &datetime, format);
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
  TypedData_Get_Struct(self, struct worksheet, &worksheet_type, ptr);
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

  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  LXW_NO_RESULT_CALL(worksheet, write_boolean, row, col, bool_value, format);
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

  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  LXW_NO_RESULT_CALL(worksheet, write_blank, row, col, format);
  return self;
}

/*  call-seq:
 *     ws.write_formula_num(cell, formula, value, format = nil) -> self
 *     ws.write_formula_num(row, col, formula, value, format = nil) -> self
 *
 *  Writes a +formula+ with +value+ into a +cell+ with +format+.
 */
VALUE
worksheet_write_formula_num_(int argc, VALUE *argv, VALUE self) {
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
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  LXW_NO_RESULT_CALL(worksheet, write_formula_num, row, col, str, format, NUM2DBL(value));
  return self;
}

/*  call-seq:
 *     ws.write_rich_string(cell, value) -> self
 *     ws.write_rich_string(row, col, value) -> self
 *
 *  Writes a +rich_string+ value into a +cell+.
 */
VALUE
worksheet_write_rich_string_(int argc, VALUE *argv, VALUE self) {
  lxw_row_t row;
  lxw_col_t col;
  VALUE value = Qnil;
  VALUE format_key = Qnil;

  rb_check_arity(argc, 2, 4);
  int larg = extract_cell(argc, argv, &row, &col);
  VALUE workbook = rb_iv_get(self, "@workbook");

  if (larg < argc) {
    value = argv[larg];
    if (TYPE(value) == T_ARRAY)
      value = rb_funcall(cRichString, rb_intern("new"), 2, workbook, value);
    else if (rb_class_of(value) != cRichString) {
      rb_raise(rb_eArgError, "Wrong type of value: %"PRIsVALUE, rb_class_of(value));
    }
    ++larg;
  }

  if (larg < argc) {
    format_key = argv[larg];
    ++larg;
  }

  if (NIL_P(value)) {
    rb_raise(rb_eArgError, "No value specified");
  }

  lxw_format *format = workbook_get_format(workbook, format_key);
  lxw_rich_string_tuple **rich_string = serialize_rich_string(value);
  if (rich_string) {
    LXW_ERR_RESULT_CALL(worksheet, write_rich_string, row, col, rich_string, format);
  }
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

  val = rb_hash_aref(opts, ID2SYM(rb_intern("collapse")));
  if (val != Qnil && val)
    options.collapsed = 1;

  val = rb_hash_aref(opts, ID2SYM(rb_intern("level")));
  if (val != Qnil)
    options.level = NUM2INT(val);

  struct worksheet *ptr;
  TypedData_Get_Struct(self, struct worksheet, &worksheet_type, ptr);
  if (options.hidden || options.collapsed || options.level) {
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

  val = rb_hash_aref(opts, ID2SYM(rb_intern("collapse")));
  if (val != Qnil && val)
    options.collapsed = 1;

  val = rb_hash_aref(opts, ID2SYM(rb_intern("level")));
  if (val != Qnil)
    options.level = NUM2INT(val);

  struct worksheet *ptr;
  TypedData_Get_Struct(self, struct worksheet, &worksheet_type, ptr);
  if (options.hidden || options.collapsed || options.level) {
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
  TypedData_Get_Struct(self, struct worksheet, &worksheet_type, ptr);

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
 *     ws.insert_image_buffer(cell, data, options) -> self
 *     ws.insert_image_buffer(row, col, data, options) -> self
 *
 *  Adds data validation or limits user input to a list of values.
 */
VALUE
worksheet_insert_image_buffer_(int argc, VALUE *argv, VALUE self) {
  lxw_row_t row;
  lxw_col_t col;
  VALUE data = Qnil;
  VALUE opts = Qnil;
  lxw_image_options options;
  char with_options = '\0';

  rb_check_arity(argc, 2, 4);
  int larg = extract_cell(argc, argv, &row, &col);

  if (larg < argc) {
    data = argv[larg];
    ++larg;
  } else {
    rb_raise(rb_eArgError, "Cannot embed image without data");
  }

  if (larg < argc) {
    opts = argv[larg];
    ++larg;
  }
  struct worksheet *ptr;
  TypedData_Get_Struct(self, struct worksheet, &worksheet_type, ptr);

  if (!NIL_P(opts)) {
    options = val_to_lxw_image_options(opts, &with_options);
  }

  lxw_error error;
  if (with_options) {
    error = worksheet_insert_image_buffer_opt(ptr->worksheet, row, col, (const unsigned char *) RSTRING_PTR(data), RSTRING_LEN(data), &options);
  } else {
    error = worksheet_insert_image_buffer(ptr->worksheet, row, col, (const unsigned char *) RSTRING_PTR(data), RSTRING_LEN(data));
  }
  handle_lxw_error(error);

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
  lxw_chart_options options;
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
    options = val_to_lxw_chart_options(opts, &with_options);
  }

  struct worksheet *ptr;
  struct chart *chart_ptr;
  TypedData_Get_Struct(self, struct worksheet, &worksheet_type, ptr);
  TypedData_Get_Struct(chart, struct chart, &chart_type, chart_ptr);

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

  LXW_NO_RESULT_CALL(worksheet, merge_range, row1, col1, row2, col2, str, format);
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

  LXW_NO_RESULT_CALL(worksheet, autofilter, row_from, col_from, row_to, col_to);
  return self;
}

/*  call-seq: ws.activate -> self
 *
 *  Set the worksheet to be active on opening the workbook.
 */
VALUE
worksheet_activate_(VALUE self) {
  LXW_NO_RESULT_CALL(worksheet, activate);
  return self;
}

/*  call-seq: ws.select -> self
 *
 *  Set the worksheet to be selected on opening the workbook.
 */
VALUE
worksheet_select_(VALUE self) {
  LXW_NO_RESULT_CALL(worksheet, select);
  return self;
}

/*  call-seq: ws.hide -> self
 *
 *  Hide the worksheet.
 */
VALUE
worksheet_hide_(VALUE self) {
  LXW_NO_RESULT_CALL(worksheet, hide);
  return self;
}

/*  call-seq: ws.set_fitst_sheet -> self
 *
 *  Set the sheet to be the first visible in the sheets list (which is placed
 *  under the work area in Excel).
 */
VALUE
worksheet_set_first_sheet_(VALUE self) {
  LXW_NO_RESULT_CALL(worksheet, set_first_sheet);
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
  TypedData_Get_Struct(self, struct worksheet, &worksheet_type, ptr);
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
  LXW_NO_RESULT_CALL(worksheet, split_panes, NUM2DBL(vertical), NUM2DBL(horizontal));
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

  LXW_NO_RESULT_CALL(worksheet, set_selection, row_from, col_from, row_to, col_to);
  return self;
}

/*  call-seq: ws.set_landscape -> self
 *
 *  Sets print orientation of the worksheet to landscape.
 */
VALUE
worksheet_set_landscape_(VALUE self) {
  LXW_NO_RESULT_CALL(worksheet, set_landscape);
  return self;
}

/*  call-seq: ws.set_portrait -> self
 *
 *  Sets print orientation of the worksheet to portrait.
 */
VALUE
worksheet_set_portrait_(VALUE self) {
  LXW_NO_RESULT_CALL(worksheet, set_portrait);
  return self;
}

/*  call-seq: ws.set_page_view -> self
 *
 *  Sets worksheet displa mode to "Page View/Layout".
 */
VALUE
worksheet_set_page_view_(VALUE self) {
  LXW_NO_RESULT_CALL(worksheet, set_page_view);
  return self;
}

/*  call-seq: ws.paper=(type) -> type
 *
 *  Sets the paper +type+ for printing. See {libxlsxwriter doc}[https://libxlsxwriter.github.io/worksheet_8h.html#a9f8af12017797b10c5ee68e04149032f] for options.
 */
VALUE
worksheet_set_paper_(VALUE self, VALUE paper_type) {
  LXW_NO_RESULT_CALL(worksheet, set_paper, NUM2INT(paper_type));
  return self;
}

/* Sets the worksheet margins (Numeric) for the printed page. */
VALUE
worksheet_set_margins_(VALUE self, VALUE left, VALUE right, VALUE top, VALUE bottom) {
  LXW_NO_RESULT_CALL(worksheet, set_margins, NUM2DBL(left), NUM2DBL(right), NUM2DBL(top), NUM2DBL(bottom));
  return self;
}

/*  call-seq: ws.set_header(text[, opts]) -> self
 *
 *  See {libxlsxwriter doc}[https://libxlsxwriter.github.io/worksheet_8h.html#a4070c24ed5107f33e94f30a1bf865ba9]
 *  for the +text+ control characters.
 *
 *  Currently the only supported option is +:margin+ (Numeric).
 */
VALUE
worksheet_set_header_(int argc, VALUE *argv, VALUE self) {
  rb_check_arity(argc, 1, 2);
  struct _header_options ho = _extract_header_options(argc, argv);
  struct worksheet *ptr;
  TypedData_Get_Struct(self, struct worksheet, &worksheet_type, ptr);
  lxw_error err;
  if (!ho.with_options) {
    err = worksheet_set_header(ptr->worksheet, ho.str);
  } else {
    err = worksheet_set_header_opt(ptr->worksheet, ho.str, &ho.options);
  }
  handle_lxw_error(err);
  return self;
}

/*  call-seq: ws.set_footer(text, opts)
 *
 *  See #set_header for params description.
 */
VALUE
worksheet_set_footer_(int argc, VALUE *argv, VALUE self) {
  struct worksheet *ptr;
  rb_check_arity(argc, 1, 2);
  struct _header_options ho = _extract_header_options(argc, argv);
  TypedData_Get_Struct(self, struct worksheet, &worksheet_type, ptr);
  lxw_error err;
  if (!ho.with_options) {
    err = worksheet_set_footer(ptr->worksheet, ho.str);
  } else {
    err = worksheet_set_footer_opt(ptr->worksheet, ho.str, &ho.options);
  }
  handle_lxw_error(err);
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
  size_t i;
  for (i = 0; i<len; ++i) {
    rows[i] = NUM2INT(rb_ary_entry(val, i));
  }
  rows[len] = 0;
  LXW_NO_RESULT_CALL(worksheet, set_h_pagebreaks, rows);
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
  size_t i;
  for (i = 0; i<len; ++i) {
    cols[i] = value_to_col(rb_ary_entry(val, i));
  }
  cols[len] = 0;
  LXW_NO_RESULT_CALL(worksheet, set_v_pagebreaks, cols);
  return val;
}

/* Changes the default print direction */
VALUE
worksheet_print_across_(VALUE self) {
  LXW_NO_RESULT_CALL(worksheet, print_across);
  return self;
}

/*  call-seq: ws.zoom=(zoom) -> zoom
 *
 *  Sets the worksheet zoom factor in the range 10 <= +zoom+ <= 400.
 */
VALUE
worksheet_set_zoom_(VALUE self, VALUE val) {
  LXW_NO_RESULT_CALL(worksheet, set_zoom, NUM2INT(val));
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
  LXW_NO_RESULT_CALL(worksheet, gridlines, NUM2INT(value));
  return value;
}

/* Center the worksheet data horizontally between the margins on the printed page */
VALUE
worksheet_center_horizontally_(VALUE self){
  LXW_NO_RESULT_CALL(worksheet, center_horizontally);
  return self;
}

/* Center the worksheet data vertically between the margins on the printed page */
VALUE
worksheet_center_vertically_(VALUE self) {
  LXW_NO_RESULT_CALL(worksheet, center_vertically);
  return self;
}

/* Print rows and column header (wich is disabled by default). */
VALUE
worksheet_print_row_col_headers_(VALUE self) {
  LXW_NO_RESULT_CALL(worksheet, print_row_col_headers);
  return self;
}

/*  call-seq: ws.repeat_rows(row_from, row_to)
 *
 *  Sets rows to be repeatedly printed out on all pages.
 */
VALUE
worksheet_repeat_rows_(VALUE self, VALUE row_from, VALUE row_to) {
  LXW_NO_RESULT_CALL(worksheet, repeat_rows, NUM2INT(row_from), NUM2INT(row_to));
  return self;
}

/*  call-seq: ws.repeat_columns(col_from, col_to)
 *
 *  Sets columns to be repeatedly printed out on all pages.
 */
VALUE
worksheet_repeat_columns_(VALUE self, VALUE col_from, VALUE col_to) {
  LXW_NO_RESULT_CALL(worksheet, repeat_columns, value_to_col(col_from), value_to_col(col_to));
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

  LXW_NO_RESULT_CALL(worksheet, print_area, row_from, col_from, row_to, col_to);
  return self;
}

/*  call-seq: ws.fit_to_pages(width, height) -> self
 *
 *  Fits the printed area to a specific number of pages both vertically and
 *  horizontally.
 */
VALUE
worksheet_fit_to_pages_(VALUE self, VALUE width, VALUE height) {
  LXW_NO_RESULT_CALL(worksheet, fit_to_pages, NUM2INT(width), NUM2INT(height));
  return self;
}

/*  call-seq: ws.start_page=(page) -> page
 *
 *  Set the number of the first printed page.
 */
VALUE
worksheet_set_start_page_(VALUE self, VALUE start_page) {
  LXW_NO_RESULT_CALL(worksheet, set_start_page, NUM2INT(start_page));
  return start_page;
}

/*  call-seq: ws.print_scale=(scale) -> scale
 *
 *  Sets the +scale+ factor of the printed page (10 <= +scale+ <= 400).
 */
VALUE
worksheet_set_print_scale_(VALUE self, VALUE scale) {
  LXW_NO_RESULT_CALL(worksheet, set_print_scale, NUM2INT(scale));
  return scale;
}

/* Sets text direction to rtl (e.g. for worksheets on Hebrew or Arabic). */
VALUE
worksheet_right_to_left_(VALUE self) {
  LXW_NO_RESULT_CALL(worksheet, right_to_left);
  return self;
}

/* Hides all zero values. */
VALUE
worksheet_hide_zero_(VALUE self) {
  LXW_NO_RESULT_CALL(worksheet, hide_zero);
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
  LXW_NO_RESULT_CALL(worksheet, set_tab_color, NUM2INT(color));
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
  struct _protect_options po = _extract_protect_options(argc, argv);
  LXW_NO_RESULT_CALL(worksheet, protect, po.password, (po.with_options ? &po.options : NULL));
  return self;
}

/*  call-seq: ws.outline_settings = { visible: true, symbols_below: true, symbols_right: true, auto_style: false }
 *
 *  Sets the Outline and Grouping display properties.
 */
VALUE
worksheet_outline_settings_(VALUE self, VALUE opts) {
  VALUE value;
#define parse_param(name, default)                                      \
  value = rb_hash_aref(opts, ID2SYM(rb_intern(#name)));                 \
  uint8_t name = NIL_P(value) ? default : (value ? LXW_TRUE : LXW_FALSE)
  parse_param(visible, LXW_TRUE);
  parse_param(symbols_below, LXW_TRUE);
  parse_param(symbols_right, LXW_TRUE);
  parse_param(auto_style, LXW_FALSE);
#undef parse_param
  LXW_NO_RESULT_CALL(worksheet, outline_settings, visible, symbols_below, symbols_right, auto_style);
  return self;
}

/*  call-seq: ws.set_default_row(height, hide_unuser_rows) -> self
 *
 *  Sets the default row properties for the worksheet.
 */
VALUE
worksheet_set_default_row_(VALUE self, VALUE height, VALUE hide_unused_rows) {
  uint8_t hide_ur = !NIL_P(hide_unused_rows) && hide_unused_rows != Qfalse ? 1 : 0;
  LXW_NO_RESULT_CALL(worksheet, set_default_row, NUM2DBL(height), hide_ur);
  return self;
}

/*  call-seq: ws.horizontal_dpi -> int
 *
 *  Returns the horizontal dpi.
 */
VALUE
worksheet_get_horizontal_dpi_(VALUE self) {
  struct worksheet *ptr;
  TypedData_Get_Struct(self, struct worksheet, &worksheet_type, ptr);
  return INT2NUM(ptr->worksheet->horizontal_dpi);
}

/*  call-seq: ws.horizontal_dpi=(dpi) -> dpi
 *
 *  Sets the horizontal dpi.
 */
VALUE
worksheet_set_horizontal_dpi_(VALUE self, VALUE val) {
  struct worksheet *ptr;
  TypedData_Get_Struct(self, struct worksheet, &worksheet_type, ptr);
  ptr->worksheet->horizontal_dpi = NUM2INT(val);
  return val;
}

/*  call-seq: ws.vertical_dpi -> int
 *
 *  Returns the vertical dpi.
 */
VALUE
worksheet_get_vertical_dpi_(VALUE self) {
  struct worksheet *ptr;
  TypedData_Get_Struct(self, struct worksheet, &worksheet_type, ptr);
  return INT2NUM(ptr->worksheet->vertical_dpi);
}

/*  call-seq: ws.vertical_dpi=(dpi) -> dpi
 *
 *  Sets the vertical dpi.
 */
VALUE
worksheet_set_vertical_dpi_(VALUE self, VALUE val) {
  struct worksheet *ptr;
  TypedData_Get_Struct(self, struct worksheet, &worksheet_type, ptr);
  ptr->worksheet->vertical_dpi = NUM2INT(val);
  return val;
}

/*  call-seq:
 *     ws.add_data_validation(cell, options) -> self
 *     ws.add_data_validation(range, options) -> self
 *     ws.add_data_validation(row, col, options) -> self
 *     ws.add_data_validation(row_from, col_from, row_to, col_to, options) -> self
 *
 *  Adds data validation or limits user input to a list of values.
 */
VALUE
worksheet_data_validation_(int argc, VALUE *argv, VALUE self) {
  lxw_data_validation data_validation = {0};
  rb_check_arity(argc, 2, 5);
  char range = 0;
  if (argc > 3) {
    range = 1;
  } else if (TYPE(argv[0]) == T_STRING) {
    char *str = RSTRING_PTR(argv[0]);
    if (strstr(str, ":")) {
      range = 1;
    } else if ((str[0] >= 'A' && str[0] <= 'Z') ||
               (str[0] >= 'a' && str[0] <= 'z'))
      range = argc > 2;
  }

  lxw_row_t row, row2;
  lxw_col_t col, col2;
  int larg;
  if (range) {
    larg = extract_range(argc, argv, &row, &col, &row2, &col2);
  } else {
    larg = extract_cell(argc, argv, &row, &col);
  }

  if (larg != argc - 1) {
    rb_raise(rb_eArgError, "Wrong number of arguments");
  }
  VALUE opts = argv[larg];
  if (TYPE(opts) != T_HASH) {
    rb_raise(rb_eTypeError, "Wrong type for options %"PRIsVALUE", Hash expected", rb_obj_class(opts));
  }

  VALUE val;
  val = rb_hash_aref(opts, ID2SYM(rb_intern("validate")));
  if (!NIL_P(val)) {
    data_validation.validate = NUM2INT(val);
  }

  val = rb_hash_aref(opts, ID2SYM(rb_intern("criteria")));
  if (!NIL_P(val)) {
    data_validation.criteria = NUM2INT(val);
  }

  val = rb_hash_aref(opts, ID2SYM(rb_intern("ignore_blank")));
  if (!NIL_P(val) && val) {
    data_validation.ignore_blank = LXW_VALIDATION_ON;
  }

  val = rb_hash_aref(opts, ID2SYM(rb_intern("show_input")));
  if (!NIL_P(val) && val) {
    data_validation.show_input = LXW_VALIDATION_ON;
  }

  val = rb_hash_aref(opts, ID2SYM(rb_intern("show_error")));
  if (!NIL_P(val) && val) {
    data_validation.show_error = LXW_VALIDATION_ON;
  }

  val = rb_hash_aref(opts, ID2SYM(rb_intern("error_type")));
  if (!NIL_P(val)) {
    data_validation.error_type = NUM2INT(val);
  }

  val = rb_hash_aref(opts, ID2SYM(rb_intern("dropdown")));
  if (!NIL_P(val) && val) {
    data_validation.dropdown = LXW_VALIDATION_ON;
  }

  val = rb_hash_aref(opts, ID2SYM(rb_intern("input_title")));
  if (!NIL_P(val)) {
    data_validation.input_title = RSTRING_PTR(val);
  }

  val = rb_hash_aref(opts, ID2SYM(rb_intern("input_message")));
  if (!NIL_P(val)) {
    data_validation.input_message = RSTRING_PTR(val);
  }

  val = rb_hash_aref(opts, ID2SYM(rb_intern("error_title")));
  if (!NIL_P(val)) {
    data_validation.error_title = RSTRING_PTR(val);
  }

  val = rb_hash_aref(opts, ID2SYM(rb_intern("error_message")));
  if (!NIL_P(val)) {
    data_validation.error_message = RSTRING_PTR(val);
  }

  VALUE is_datetime;
#define parse_array_0(x)
#define parse_array_1(prefix)                                           \
    if (TYPE(val) == T_ARRAY) {                                         \
      int len = RARRAY_LEN(val);                                        \
      int i;                                                            \
      data_validation.prefix##_list = malloc(sizeof(char *) * (len + 1)); \
      data_validation.prefix##_list[len] = NULL;                        \
      for (i = 0; i < len; ++i) {                                       \
        data_validation.prefix##_list[i] = RSTRING_PTR(rb_ary_entry(val, i)); \
      }                                                                 \
    } else

#define parse_value(prefix, key, handle_array)                          \
  val = rb_hash_aref(opts, ID2SYM(rb_intern(key)));                     \
  if (!NIL_P(val)) {                                                    \
    switch(TYPE(val)) {                                                 \
    case T_FLOAT: case T_FIXNUM: case T_BIGNUM:                         \
      data_validation.prefix##_number = NUM2DBL(val);                   \
      break;                                                            \
    case T_STRING:                                                      \
      data_validation.prefix##_formula = RSTRING_PTR(val);              \
      if (data_validation.prefix##_formula[0] != '=') {                 \
        data_validation.prefix##_formula = 0;                           \
        rb_raise(rb_eArgError, "Is not a formula: %"PRIsVALUE, val);    \
      }                                                                 \
      break;                                                            \
    default:                                                            \
      is_datetime = rb_funcall(val, rb_intern("is_a?"), 1, rb_cTime);   \
      parse_array_##handle_array(prefix)                                \
      if (is_datetime || rb_respond_to(val, rb_intern("to_time"))) {    \
        if (!is_datetime)                                               \
          val = rb_funcall(val, rb_intern("to_time"), 0);               \
        data_validation.prefix##_datetime = value_to_lxw_datetime(val); \
      } else {                                                          \
        rb_raise(rb_eTypeError, "Cannot handle " key " type %"PRIsVALUE, rb_class_of(val)); \
      }                                                                 \
    }                                                                   \
  }

  parse_value(minimum, "min", 0);
  parse_value(maximum, "max", 0);
  parse_value(value, "value", 1);

#undef parse_array_0
#undef parse_array_1
#undef parse_value

  struct worksheet *ptr;
  TypedData_Get_Struct(self, struct worksheet, &worksheet_type, ptr);
  lxw_error err;
  if (range) {
    err = worksheet_data_validation_range(ptr->worksheet, row, col, row2, col2, &data_validation);
  } else {
    err = worksheet_data_validation_cell(ptr->worksheet, row, col, &data_validation);
  }

  if (data_validation.value_list) {
    free(data_validation.value_list);
  }

  handle_lxw_error(err);

  return self;
}

/*  call-seq: wb.vba_name = name
 *
 *  Set the VBA name for the worksheet.
 */
VALUE
worksheet_set_vba_name_(VALUE self, VALUE name) {
  LXW_ERR_RESULT_CALL(worksheet, set_vba_name, StringValueCStr(name));
  return name;
}

// Helpers

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

VALUE
rb_extract_col(VALUE _self, VALUE col) {
  return INT2FIX(value_to_col(col));
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
lxw_chart_options
val_to_lxw_chart_options(VALUE opts, char *with_options) {
  VALUE val;
  lxw_chart_options options = {
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
  SET_IMG_OPT("object_position", options.object_position = NUM2INT(val));
  return options;
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
  SET_IMG_OPT("object_position", options.object_position = NUM2INT(val));
  SET_IMG_OPT("description", options.description = StringValueCStr(val));
  SET_IMG_OPT("url", options.url = StringValueCStr(val));
  SET_IMG_OPT("tip", options.tip = StringValueCStr(val));
  return options;
}
#undef SET_IMG_OPT


void
init_xlsxwriter_worksheet() {
  /* Xlsx worksheet */
  cWorksheet = rb_define_class_under(mXlsxWriter, "Worksheet", rb_cObject);

  rb_define_alloc_func(cWorksheet, worksheet_alloc);
  rb_define_attr(cWorksheet, "workbook", 1, 0);
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
  rb_define_method(cWorksheet, "write_rich_string", worksheet_write_rich_string_, -1);
  rb_define_method(cWorksheet, "set_row", worksheet_set_row_, 2);
  rb_define_method(cWorksheet, "set_column", worksheet_set_column_, 3);
  rb_define_method(cWorksheet, "insert_image", worksheet_insert_image_, -1);
  rb_define_method(cWorksheet, "insert_image_buffer", worksheet_insert_image_buffer_, -1);
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
  rb_define_method(cWorksheet, "set_header", worksheet_set_header_, -1);
  rb_define_method(cWorksheet, "set_footer", worksheet_set_footer_, -1);
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
  rb_define_method(cWorksheet, "outline_settings=", worksheet_outline_settings_, 1);
  rb_define_method(cWorksheet, "set_default_row", worksheet_set_default_row_, 2);

  rb_define_method(cWorksheet, "horizontal_dpi", worksheet_get_horizontal_dpi_, 0);
  rb_define_method(cWorksheet, "horizontal_dpi=", worksheet_set_horizontal_dpi_, 1);
  rb_define_method(cWorksheet, "vertical_dpi", worksheet_get_vertical_dpi_, 0);
  rb_define_method(cWorksheet, "vertical_dpi=", worksheet_set_vertical_dpi_, 1);

  rb_define_method(cWorksheet, "add_data_validation", worksheet_data_validation_, -1);
  rb_define_method(cWorksheet, "vba_name=", worksheet_set_vba_name_, 1);

  rb_define_private_method(cWorksheet, "extract_column", rb_extract_col, 1);

#define MAP_LXW_WH_CONST(name, val_name) rb_define_const(cWorksheet, #name, INT2NUM(LXW_##val_name))
#define MAP_LXW_WH_CONST1(name) MAP_LXW_WH_CONST(name, name)
  MAP_LXW_WH_CONST(DEF_COL_WIDTH, DEF_COL_WIDTH);
  MAP_LXW_WH_CONST(DEF_ROW_HEIGHT, DEF_ROW_HEIGHT);

  MAP_LXW_WH_CONST(GRIDLINES_HIDE_ALL, HIDE_ALL_GRIDLINES);
  MAP_LXW_WH_CONST(GRIDLINES_SHOW_SCREEN, SHOW_SCREEN_GRIDLINES);
  MAP_LXW_WH_CONST(GRIDLINES_SHOW_PRINT, SHOW_PRINT_GRIDLINES);
  MAP_LXW_WH_CONST(GRIDLINES_SHOW_ALL, SHOW_ALL_GRIDLINES);

  MAP_LXW_WH_CONST1(VALIDATION_TYPE_INTEGER);
  MAP_LXW_WH_CONST1(VALIDATION_TYPE_INTEGER_FORMULA);
  MAP_LXW_WH_CONST1(VALIDATION_TYPE_DECIMAL);
  MAP_LXW_WH_CONST1(VALIDATION_TYPE_DECIMAL_FORMULA);
  MAP_LXW_WH_CONST1(VALIDATION_TYPE_LIST);
  MAP_LXW_WH_CONST1(VALIDATION_TYPE_LIST_FORMULA);
  MAP_LXW_WH_CONST1(VALIDATION_TYPE_DATE);
  MAP_LXW_WH_CONST1(VALIDATION_TYPE_DATE_FORMULA);
  MAP_LXW_WH_CONST1(VALIDATION_TYPE_TIME);
  MAP_LXW_WH_CONST1(VALIDATION_TYPE_TIME_FORMULA);
  MAP_LXW_WH_CONST1(VALIDATION_TYPE_LENGTH);
  MAP_LXW_WH_CONST1(VALIDATION_TYPE_LENGTH_FORMULA);
  MAP_LXW_WH_CONST1(VALIDATION_TYPE_CUSTOM_FORMULA);
  MAP_LXW_WH_CONST1(VALIDATION_TYPE_ANY);

  MAP_LXW_WH_CONST1(VALIDATION_CRITERIA_BETWEEN);
  MAP_LXW_WH_CONST1(VALIDATION_CRITERIA_NOT_BETWEEN);
  MAP_LXW_WH_CONST1(VALIDATION_CRITERIA_EQUAL_TO);
  MAP_LXW_WH_CONST1(VALIDATION_CRITERIA_NOT_EQUAL_TO);
  MAP_LXW_WH_CONST1(VALIDATION_CRITERIA_GREATER_THAN);
  MAP_LXW_WH_CONST1(VALIDATION_CRITERIA_LESS_THAN);
  MAP_LXW_WH_CONST1(VALIDATION_CRITERIA_GREATER_THAN_OR_EQUAL_TO);
  MAP_LXW_WH_CONST1(VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO);

  MAP_LXW_WH_CONST1(OBJECT_POSITION_DEFAULT);
  MAP_LXW_WH_CONST1(OBJECT_MOVE_AND_SIZE);
  MAP_LXW_WH_CONST1(OBJECT_MOVE_DONT_SIZE);
  MAP_LXW_WH_CONST1(OBJECT_DONT_MOVE_DONT_SIZE);
  MAP_LXW_WH_CONST1(OBJECT_MOVE_AND_SIZE_AFTER);
#undef MAP_LXW_WH_CONST1
#undef MAP_LXW_WH_CONST
}
