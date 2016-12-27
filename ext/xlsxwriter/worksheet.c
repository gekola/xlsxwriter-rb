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

VALUE
worksheet_write_datetime_(VALUE self, VALUE row, VALUE col, VALUE value, VALUE format_key) {
  struct lxw_datetime datetime = {
    .year = NUM2INT(rb_funcall(value, rb_intern("year"), 0)),
    .month = NUM2INT(rb_funcall(value, rb_intern("month"), 0)),
    .day = NUM2INT(rb_funcall(value, rb_intern("day"), 0)),
    .hour = NUM2INT(rb_funcall(value, rb_intern("hour"), 0)),
    .min = NUM2INT(rb_funcall(value, rb_intern("min"), 0)),
    .sec = NUM2DBL(rb_funcall(value, rb_intern("sec"), 0)) +
           NUM2DBL(rb_funcall(value, rb_intern("second_fraction"), 0))
  };
  struct worksheet *ptr;
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_write_datetime(ptr->worksheet, NUM2INT(row), value_to_col(col), &datetime, format);
  return self;
}

VALUE
worksheet_write_url_(VALUE self, VALUE row, VALUE col, VALUE value, VALUE format_key) {
  const char *str = RSTRING_PTR(value);
  struct worksheet *ptr;
  VALUE workbook = rb_iv_get(self, "@workbook");
  lxw_format *format = workbook_get_format(workbook, format_key);
  Data_Get_Struct(self, struct worksheet, ptr);
  worksheet_write_url(ptr->worksheet, NUM2INT(row), value_to_col(col), str, format);
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
worksheet_gridlines_(VALUE self, VALUE value) {
  struct worksheet *ptr;
  Data_Get_Struct(self, struct worksheet, ptr);

  worksheet_gridlines(ptr->worksheet, NUM2INT(value));

  return value;
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
