#include "worksheet.h"
#include "workbook.h"

VALUE
worksheet_alloc(VALUE klass)
{
  VALUE obj;
  struct worksheet *ptr;

  obj = Data_Make_Struct(klass, struct worksheet, NULL, worksheet_free, ptr);

  ptr->worksheet = NULL;

  return obj;
}

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
  return self;
}

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

VALUE
worksheet_write_string_(VALUE self, VALUE row, VALUE col, VALUE value, VALUE format_key) {
  const char *str = StringValueCStr(value);
  struct worksheet *ptr;
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_write_string(ptr->worksheet, NUM2INT(row), value_to_col(col), str, format);
  return self;
}

VALUE
worksheet_write_number_(VALUE self, VALUE row, VALUE col, VALUE value, VALUE format_key) {
  const double num = NUM2DBL(value);
  struct worksheet *ptr;
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_write_number(ptr->worksheet, NUM2INT(row), value_to_col(col), num, format);
  return self;
}

VALUE
worksheet_write_formula_(VALUE self, VALUE row, VALUE col, VALUE value, VALUE format_key) {
  const char *str = RSTRING_PTR(value);
  struct worksheet *ptr;
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_write_formula(ptr->worksheet, NUM2INT(row), value_to_col(col), str, format);
  return self;
}

VALUE worksheet_write_array_formula_(VALUE self, VALUE row_from, VALUE col_from, VALUE row_to, VALUE col_to, VALUE value, VALUE format_key) {
  const char *str = StringValueCStr(value);
  struct worksheet *ptr;
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_write_array_formula(ptr->worksheet, NUM2INT(row_from), value_to_col(col_from), NUM2INT(row_to), value_to_col(col_to), str, format);
  return self;
}

VALUE
worksheet_write_datetime_(VALUE self, VALUE row, VALUE col, VALUE value, VALUE format_key) {
  struct lxw_datetime datetime = value_to_lxw_datetime(value);
  struct worksheet *ptr;
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_write_datetime(ptr->worksheet, NUM2INT(row), value_to_col(col), &datetime, format);
  return self;
}

VALUE
worksheet_write_url_(int argc, VALUE *argv, VALUE self) {
  rb_check_arity(argc, 3, 5);
  VALUE row = argv[0];
  VALUE col = argv[1];
  VALUE url = argv[2];
  VALUE format_key = Qnil; // argv[3];
  VALUE opts = Qnil;
  const char *url_str = RSTRING_PTR(url);
  char *string = NULL;
  char *tooltip = NULL;
  if (argc > 3) {
    switch(TYPE(argv[3])) {
    case T_HASH:
      opts = argv[3];
      break;
    case T_SYMBOL: case T_STRING: case T_NIL:
      format_key = argv[3];
      break;
    default:
      rb_raise(rb_eTypeError, "Expected Hash, symbol or string but got %"PRIsVALUE, rb_obj_class(argv[3]));
    }
  }
  if (argc == 5) {
    switch(TYPE(argv[4])) {
    case T_HASH: case T_NIL:
      opts = argv[4];
      break;
    default:
      rb_raise(rb_eTypeError, "Expected Hash, got %"PRIsVALUE, rb_obj_class(argv[3]));
    }
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
    worksheet_write_url_opt(ptr->worksheet, NUM2INT(row), value_to_col(col), url_str, format, string, tooltip);
  } else {
    worksheet_write_url(ptr->worksheet, NUM2INT(row), value_to_col(col), url_str, format);
  }
  return self;
}

VALUE
worksheet_write_boolean_(VALUE self, VALUE row, VALUE col, VALUE value, VALUE format_key) {
  int bool_value = value && (value != Qnil);
  struct worksheet *ptr;
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_write_boolean(ptr->worksheet, NUM2INT(row), value_to_col(col), bool_value, format);
  return self;
}

VALUE
worksheet_write_blank_(VALUE self, VALUE row, VALUE col, VALUE format_key) {
  struct worksheet *ptr;
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_write_blank(ptr->worksheet, NUM2INT(row), value_to_col(col), format);
  return self;
}

VALUE worksheet_write_formula_num_(VALUE self, VALUE row, VALUE col, VALUE formula, VALUE format_key, VALUE value) {
  const char *str = RSTRING_PTR(formula);
  struct worksheet *ptr;
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_write_formula_num(ptr->worksheet, NUM2INT(row), value_to_col(col), str, format, NUM2DBL(value));
  return self;
}

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

#define SET_IMG_OPT(key, setter) {                    \
    val = rb_hash_aref(opts, ID2SYM(rb_intern(key))); \
    if (!NIL_P(val)) {                                \
      with_options = '\1';                            \
      setter;                                         \
    }                                                 \
  }

VALUE
worksheet_insert_image_(VALUE self, VALUE row, VALUE col, VALUE fname, VALUE opts) {
  VALUE val;
  lxw_image_options options = {
    .x_offset = 0,
    .y_offset = 0,
    .x_scale = 1.0,
    .y_scale = 1.0
  };
  char with_options = '\0';

  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);

  SET_IMG_OPT("offset",   options.x_offset = options.y_offset = NUM2INT(val));
  SET_IMG_OPT("x_offset", options.x_offset =                    NUM2INT(val));
  SET_IMG_OPT("y_offset",                    options.y_offset = NUM2INT(val));
  SET_IMG_OPT("scale",    options.x_scale  = options.y_scale  = NUM2DBL(val));
  SET_IMG_OPT("x_scale",  options.x_scale  =                    NUM2DBL(val));
  SET_IMG_OPT("y_scale",                     options.y_scale  = NUM2DBL(val));

  if (with_options) {
    worksheet_insert_image_opt(ptr->worksheet, NUM2INT(row), value_to_col(col), StringValueCStr(fname), &options);
  } else {
    worksheet_insert_image(ptr->worksheet, NUM2INT(row), value_to_col(col), StringValueCStr(fname));
  }

  return self;
}

VALUE
worksheet_insert_chart_(VALUE self, VALUE row, VALUE col, VALUE chart, VALUE opts) {
  rb_raise(rb_eRuntimeError, "worksheet_insert_chart is not ported yet");
  return self;
}

#undef SET_IMG_OPT

VALUE
worksheet_merge_range_(VALUE self, VALUE row_from, VALUE col_from,
                       VALUE row_to, VALUE col_to, VALUE value, VALUE format_key) {
  lxw_format *format = NULL;
  lxw_col_t col1 = value_to_col(col_from);
  lxw_col_t col2 = value_to_col(col_to);
  VALUE workbook = rb_iv_get(self, "@workbook");

  format = workbook_get_format(workbook, format_key);

  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);

  worksheet_merge_range(ptr->worksheet, NUM2INT(row_from), col1,
                        NUM2INT(row_to), col2, StringValueCStr(value), format);
  return self;
}

VALUE
worksheet_autofilter_(VALUE self, VALUE row_from, VALUE col_from, VALUE row_to, VALUE col_to) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);

  worksheet_autofilter(ptr->worksheet, NUM2INT(row_from), value_to_col(col_from),
                                       NUM2INT(row_to),   value_to_col(col_to));
  return self;
}
VALUE worksheet_activate_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_activate(ptr->worksheet);
  return self;
}
VALUE worksheet_select_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_select(ptr->worksheet);
  return self;
}

VALUE
worksheet_hide_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_hide(ptr->worksheet);
  return self;
}

VALUE
worksheet_set_first_sheet_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_first_sheet(ptr->worksheet);
  return self;
}
VALUE
worksheet_freeze_panes_(int argc, VALUE *argv, VALUE self) {
  rb_check_arity(argc, 2, 5);
  VALUE row = argv[0];
  VALUE col = argv[1];
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  if (argc == 2) {
    worksheet_freeze_panes(ptr->worksheet, NUM2INT(row), value_to_col(col));
  } else {
    VALUE row2 = row;
    VALUE col2 = col;
    uint8_t type = 0;
    if (argc > 2)
      row2 = argv[2];
    if (argc > 3)
      col2 = argv[3];
    if (argc > 4)
      type = NUM2INT(argv[4]);
    worksheet_freeze_panes_opt(ptr->worksheet, NUM2INT(row), value_to_col(col),
                               NUM2INT(row2), value_to_col(col2), type);
  }
  return self;
}

VALUE
worksheet_split_panes_(VALUE self, VALUE vertical, VALUE horizontal) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_split_panes(ptr->worksheet, NUM2DBL(vertical), NUM2DBL(horizontal));
  return self;
}

VALUE
worksheet_set_selection_(VALUE self, VALUE row_from, VALUE col_from, VALUE row_to, VALUE col_to) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_selection(ptr->worksheet, NUM2INT(row_from), value_to_col(col_from),
                          NUM2INT(row_to), value_to_col(col_to));
  return self;
}

VALUE
worksheet_set_landscape_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_landscape(ptr->worksheet);
  return self;
}

VALUE
worksheet_set_portrait_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_portrait(ptr->worksheet);
  return self;
}

VALUE
worksheet_set_page_view_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_page_view(ptr->worksheet);
  return self;
}

VALUE
worksheet_set_paper_(VALUE self, VALUE paper_type) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_paper(ptr->worksheet, NUM2INT(paper_type));
  return self;
}

VALUE
worksheet_set_margins_(VALUE self, VALUE left, VALUE right, VALUE top, VALUE bottom) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_margins(ptr->worksheet, NUM2DBL(left), NUM2DBL(right), NUM2DBL(top), NUM2DBL(bottom));
  return self;
}

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

VALUE
worksheet_print_across_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_print_across(ptr->worksheet);
  return self;
}

VALUE
worksheet_set_zoom_(VALUE self, VALUE val) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_zoom(ptr->worksheet, NUM2INT(val));
  return self;
}

VALUE
worksheet_gridlines_(VALUE self, VALUE value) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);

  worksheet_gridlines(ptr->worksheet, NUM2INT(value));

  return value;
}

VALUE
worksheet_center_horizontally_(VALUE self){
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_center_horizontally(ptr->worksheet);
  return self;
}

VALUE
worksheet_center_vertically_(VALUE self) {
    struct worksheet *ptr;
    Data_Get_Struct(self, struct worksheet, ptr);
    worksheet_center_vertically(ptr->worksheet);
    return self;
}

VALUE
worksheet_print_row_col_headers_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_print_row_col_headers(ptr->worksheet);
  return self;
}

VALUE
worksheet_repeat_rows_(VALUE self, VALUE row_from, VALUE row_to) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_repeat_rows(ptr->worksheet, NUM2INT(row_from), NUM2INT(row_to));
  return self;
}

VALUE
worksheet_repeat_columns_(VALUE self, VALUE col_from, VALUE col_to) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_repeat_columns(ptr->worksheet, value_to_col(col_from), value_to_col(col_to));
  return self;
}

VALUE
worksheet_print_area_(VALUE self, VALUE row_from, VALUE col_from, VALUE row_to, VALUE col_to) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_print_area(ptr->worksheet, NUM2INT(row_from), value_to_col(col_from),
                                       NUM2INT(row_to),   value_to_col(col_to));
  return self;
}

VALUE
worksheet_fit_to_pages_(VALUE self, VALUE width, VALUE height) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_fit_to_pages(ptr->worksheet, NUM2INT(width), NUM2INT(height));
  return self;
}

VALUE
worksheet_set_start_page_(VALUE self, VALUE start_page) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_start_page(ptr->worksheet, NUM2INT(start_page));
  return start_page;
}

VALUE
worksheet_set_print_scale_(VALUE self, VALUE scale) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_print_scale(ptr->worksheet, NUM2INT(scale));
  return scale;
}

VALUE
worksheet_right_to_left_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_right_to_left(ptr->worksheet);
  return self;
}

VALUE
worksheet_hide_zero_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_hide_zero(ptr->worksheet);
  return self;
}

VALUE
worksheet_set_tab_color_(VALUE self, VALUE color) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_tab_color(ptr->worksheet, NUM2INT(color));
  return color;
}

VALUE
worksheet_protect_(int argc, VALUE *argv, VALUE self) {
  rb_check_arity(argc, 0, 2);
  uint8_t with_options = '\0';
  VALUE val;
  VALUE opts = Qnil;
  lxw_protection options;
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

  // All options are off by default
  memset(&options, 0, sizeof options);

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

VALUE
worksheet_set_default_row_(VALUE self, VALUE height, VALUE hide_unused_rows) {
  struct worksheet *ptr;
  uint8_t hide_ur = !NIL_P(hide_unused_rows) && hide_unused_rows != Qfalse ? 1 : 0;
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_set_default_row(ptr->worksheet, NUM2DBL(height), hide_ur);
  return self;
}


VALUE
worksheet_get_vertical_dpi_(VALUE self) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);
  return ptr->worksheet->vertical_dpi;
}

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
    rb_raise(rb_eArgError, "wrong type for col");
    return -1;
  }
}
