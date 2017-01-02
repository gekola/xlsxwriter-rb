#include "chart.h"
#include "workbook.h"

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
  rb_check_arity(argc, 2, 3);

  Data_Get_Struct(argv[0], struct chart, c_ptr);
  Data_Get_Struct(self, struct chart_series, ptr);

  if (argc > 2) {
    cats = StringValueCStr(argv[1]);
    vals = StringValueCStr(argv[2]);
  } else {
    vals = StringValueCStr(argv[1]);
  }
  if (c_ptr && c_ptr->chart) {
    ptr->series = chart_add_series(c_ptr->chart, cats, vals);
  }

  return self;
}


VALUE
chart_add_series_(int argc, VALUE *argv, VALUE self) {
  rb_check_arity(argc, 1, 2);

  VALUE series = rb_funcall(cChartSeries, rb_intern("new"), argc+1, self, argv[0], argv[1]);
  return series;
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
