#include "chart.h"
#include <ruby/encoding.h>
#include "workbook.h"
#include "worksheet.h"

VALUE
chart_alloc(VALUE klass) {
  VALUE obj;
  struct chart *ptr;

  obj = Data_Make_Struct(klass, struct chart, NULL, NULL, ptr);

  ptr->chart = NULL;

  return obj;
}

VALUE
chart_init(VALUE self, VALUE workbook, VALUE type) {
  struct chart *ptr;
  struct workbook *wb_ptr;

  Data_Get_Struct(workbook, struct workbook, wb_ptr);
  Data_Get_Struct(self, struct chart, ptr);
  if (wb_ptr && wb_ptr->workbook) {
    ptr->chart = workbook_add_chart(wb_ptr->workbook, NUM2INT(type));
  }

  rb_iv_set(self, "@workbook", workbook);

  return self;
}

VALUE
chart_axis_alloc(VALUE klass) {
  VALUE obj;
  struct chart_axis *ptr;

  obj = Data_Make_Struct(klass, struct chart_axis, NULL, NULL, ptr);

  return obj;
}

VALUE
chart_axis_init(VALUE self, VALUE chart, VALUE type) {
  struct chart_axis *ptr;
  struct chart *c_ptr;

  Data_Get_Struct(chart, struct chart, c_ptr);
  Data_Get_Struct(self, struct chart_axis, ptr);
  rb_iv_set(self, "@chart", chart);
  if (c_ptr && c_ptr->chart) {
    ID axis = rb_check_id_cstr("x", 1, NULL);
    if (axis && axis == rb_check_id(&type)) {
      ptr->axis = c_ptr->chart->x_axis;
      return self;
    }

    axis = rb_check_id_cstr("y", 1, NULL);
    if (axis && axis == rb_check_id(&type)) {
      ptr->axis = c_ptr->chart->y_axis;
      return self;
    }

    rb_raise(rb_eArgError, "Unexpected axis type %"PRIsVALUE, rb_inspect(type));
  } else {
    rb_raise(rb_eRuntimeError, "Chart seems to be already closed.");
  }

  return self;
}

VALUE
chart_series_alloc(VALUE klass) {
  VALUE obj;
  struct chart_series *ptr;

  obj = Data_Make_Struct(klass, struct chart_series, NULL, NULL, ptr);

  ptr->series = NULL;

  return obj;
}

VALUE
chart_series_init(int argc, VALUE *argv, VALUE self) {
  struct chart_series *ptr;
  struct chart *c_ptr;
  char *cats = NULL, *vals = NULL;
  rb_check_arity(argc, 1, 3);

  Data_Get_Struct(argv[0], struct chart, c_ptr);
  Data_Get_Struct(self, struct chart_series, ptr);

  if (argc > 2) {
    cats = StringValueCStr(argv[1]);
    vals = StringValueCStr(argv[2]);
  } else if (argc > 1) {
    vals = StringValueCStr(argv[1]);
  }
  if (c_ptr && c_ptr->chart) {
    ptr->series = chart_add_series(c_ptr->chart, cats, vals);
  }

  return self;
}


VALUE
chart_add_series_(int argc, VALUE *argv, VALUE self) {
  rb_check_arity(argc, 0, 2);

  VALUE series = rb_funcall(cChartSeries, rb_intern("new"), argc+1, self, argv[0], argv[1]);
  return series;
}

VALUE
chart_x_axis_(VALUE self) {
  return rb_funcall(cChartAxis, rb_intern("new"), 2, self, ID2SYM(rb_intern("x")));
}

VALUE
chart_y_axis_(VALUE self) {
  return rb_funcall(cChartAxis, rb_intern("new"), 2, self, ID2SYM(rb_intern("y")));
}

VALUE
chart_title_set_name_(VALUE self, VALUE name) {
  struct chart *ptr;
  Data_Get_Struct(self, struct chart, ptr);

  if (RTEST(name)) {
    chart_title_set_name(ptr->chart, StringValueCStr(name));
  } else {
    chart_title_off(ptr->chart);
  }

  return name;
}

VALUE
chart_title_set_name_range_(int argc, VALUE *argv, VALUE self) {
  rb_check_arity(argc, 2, 3);
  const char *str = StringValueCStr(argv[0]);
  lxw_row_t row;
  lxw_col_t col;
  extract_cell(argc - 1, argv + 1, &row, &col);
  struct chart *ptr;
  Data_Get_Struct(self, struct chart, ptr);
  chart_title_set_name_range(ptr->chart, str, row, col);
  return self;
}

VALUE
chart_legend_set_position_(VALUE self, VALUE pos) {
  struct chart *ptr;
  Data_Get_Struct(self, struct chart, ptr);

  chart_legend_set_position(ptr->chart, NUM2UINT(rb_check_to_int(pos)));
  return pos;
}

VALUE
chart_legend_set_font_(VALUE self, VALUE opts) {
  struct chart *ptr;
  lxw_chart_font font = val_to_lxw_chart_font(opts);
  Data_Get_Struct(self, struct chart, ptr);

  chart_legend_set_font(ptr->chart, &font);
  return self;
}

VALUE
chart_legend_delete_series_(VALUE self, VALUE series) {
  series = rb_check_array_type(series);
  const size_t len = RARRAY_LEN(series);
  int16_t series_arr[len+1];
  for (size_t i = 0; i < len; ++i) {
    series_arr[i] = NUM2INT(rb_check_to_int(rb_ary_entry(series, i)));
  }
  series_arr[len] = -1;

  struct chart *ptr;
  Data_Get_Struct(self, struct chart, ptr);
  chart_legend_delete_series(ptr->chart, series_arr);
  return self;
}

VALUE
chart_set_style_(VALUE self, VALUE style) {
  struct chart *ptr;
  Data_Get_Struct(self, struct chart, ptr);

  chart_set_style(ptr->chart, NUM2INT(rb_check_to_int(style)));
  return style;
}

VALUE
chart_set_rotation_(VALUE self, VALUE rotation) {
  struct chart *ptr;
  Data_Get_Struct(self, struct chart, ptr);

  chart_set_rotation(ptr->chart, NUM2UINT(rb_check_to_int(rotation)));
  return rotation;
}

VALUE
chart_set_hole_size_(VALUE self, VALUE size) {
  struct chart *ptr;
  Data_Get_Struct(self, struct chart, ptr);

  chart_set_hole_size(ptr->chart, NUM2UINT(rb_check_to_int(size)));
  return size;
}


VALUE
chart_get_axis_id_1_(VALUE self) {
  struct chart *ptr;
  Data_Get_Struct(self, struct chart, ptr);
  return UINT2NUM(ptr->chart->axis_id_1);
}

VALUE
chart_set_axis_id_1_(VALUE self, VALUE val) {
  struct chart *ptr;
  Data_Get_Struct(self, struct chart, ptr);
  ptr->chart->axis_id_1 = NUM2UINT(val);
  return val;
}

VALUE
chart_get_axis_id_2_(VALUE self) {
  struct chart *ptr;
  Data_Get_Struct(self, struct chart, ptr);
  return UINT2NUM(ptr->chart->axis_id_2);
}

VALUE
chart_set_axis_id_2_(VALUE self, VALUE val) {
  struct chart *ptr;
  Data_Get_Struct(self, struct chart, ptr);
  ptr->chart->axis_id_2 = NUM2UINT(val);
  return val;
}


VALUE
chart_axis_set_name_(VALUE self, VALUE val) {
  struct chart_axis *ptr;
  Data_Get_Struct(self, struct chart_axis, ptr);

  chart_axis_set_name(ptr->axis, StringValueCStr(val));
  return val;
}

VALUE
chart_axis_set_name_range_(int argc, VALUE *argv, VALUE self) {
  rb_check_arity(argc, 2, 3);
  const char *str = StringValueCStr(argv[0]);
  lxw_row_t row;
  lxw_col_t col;
  extract_cell(argc - 1, argv + 1, &row, &col);
  struct chart_axis *ptr;
  Data_Get_Struct(self, struct chart_axis, ptr);
  chart_axis_set_name_range(ptr->axis, str, row, col);
  return self;
}

VALUE
chart_axis_set_name_font_(VALUE self, VALUE value) {
  struct chart_axis *ptr;
  lxw_chart_font font = val_to_lxw_chart_font(value);
  Data_Get_Struct(self, struct chart_axis, ptr);

  chart_axis_set_name_font(ptr->axis, &font);
  return self;
}

VALUE
chart_axis_set_num_font_(VALUE self, VALUE value) {
  struct chart_axis *ptr;
  lxw_chart_font font = val_to_lxw_chart_font(value);
  Data_Get_Struct(self, struct chart_axis, ptr);

  chart_axis_set_num_font(ptr->axis, &font);
  return self;
}

VALUE
chart_axis_set_line_(VALUE self, VALUE opts) {
  struct chart_axis *ptr;
  lxw_chart_line line = val_to_lxw_chart_line(opts);

  Data_Get_Struct(self, struct chart_axis, ptr);

  chart_axis_set_line(ptr->axis, &line);
  return self;
}

VALUE
chart_axis_set_fill_(VALUE self, VALUE opts) {
  struct chart_axis *ptr;
  lxw_chart_fill fill = val_to_lxw_chart_fill(opts);

  Data_Get_Struct(self, struct chart_axis, ptr);

  chart_axis_set_fill(ptr->axis, &fill);
  return self;
}

VALUE
chart_series_set_categories_(int argc, VALUE *argv, VALUE self) {
  rb_check_arity(argc, 2, 5);
  const char *str = StringValueCStr(argv[0]);
  lxw_row_t row_from, row_to;
  lxw_col_t col_from, col_to;
  extract_range(argc - 1, argv + 1, &row_from, &col_from, &row_to, &col_to);
  struct chart_series *ptr;
  Data_Get_Struct(self, struct chart_series, ptr);
  chart_series_set_categories(ptr->series, str, row_from, col_from, row_to, col_to);
  return self;
}

VALUE
chart_series_set_values_(int argc, VALUE *argv, VALUE self) {
  rb_check_arity(argc, 2, 5);
  const char *str = StringValueCStr(argv[0]);
  lxw_row_t row_from, row_to;
  lxw_col_t col_from, col_to;
  extract_range(argc - 1, argv + 1, &row_from, &col_from, &row_to, &col_to);
  struct chart_series *ptr;
  Data_Get_Struct(self, struct chart_series, ptr);
  chart_series_set_values(ptr->series, str, row_from, col_from, row_to, col_to);
  return self;
}


VALUE
chart_series_set_name_(VALUE self, VALUE name) {
  struct chart_series *ptr;
  Data_Get_Struct(self, struct chart_series, ptr);

  chart_series_set_name(ptr->series, StringValueCStr(name));
  return name;
}

VALUE
chart_series_set_name_range_(int argc, VALUE *argv, VALUE self) {
  rb_check_arity(argc, 2, 3);
  const char *str = StringValueCStr(argv[0]);
  lxw_row_t row;
  lxw_col_t col;
  extract_cell(argc - 1, argv + 1, &row, &col);
  struct chart_series *ptr;
  Data_Get_Struct(self, struct chart_series, ptr);
  chart_series_set_name_range(ptr->series, str, row, col);
  return self;
}


#define SET_PROP(prop, vres) {                          \
    val = rb_hash_aref(opts, ID2SYM(rb_intern(#prop))); \
    if (RTEST(val)) res.prop = vres;                    \
  }

lxw_chart_fill
val_to_lxw_chart_fill(VALUE opts) {
  lxw_chart_fill res = {}; // All zeros
  VALUE val;

  SET_PROP(color, NUM2UINT(val));
  SET_PROP(none, 1);
  SET_PROP(transparency, NUM2UINT(val));

  return res;
}

lxw_chart_font
val_to_lxw_chart_font(VALUE opts) {
  lxw_chart_font res = {}; // Zero-initialized
  VALUE val;

  SET_PROP(name, StringValueCStr(val));
  SET_PROP(size, NUM2UINT(val));
  SET_PROP(bold, 1);
  SET_PROP(italic, 1);
  SET_PROP(underline, 1);
  SET_PROP(rotation, NUM2INT(val));
  SET_PROP(color, NUM2ULONG(val));
  SET_PROP(baseline, NUM2INT(val));

  return res;
}

lxw_chart_line
val_to_lxw_chart_line(VALUE opts) {
  lxw_chart_line res = {}; // All zeros
  VALUE val;

  SET_PROP(color, NUM2UINT(val));
  SET_PROP(none, 1);
  SET_PROP(width, NUM2DBL(val));
  SET_PROP(dash_type, NUM2UINT(val));

  return res;
}

#undef SET_PROP
