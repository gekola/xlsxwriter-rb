#include "chart.h"
#include <ruby/encoding.h>
#include "workbook.h"
#include "worksheet.h"

VALUE cChart;
VALUE cChartSeries;
VALUE cChartAxis;


VALUE
chart_alloc(VALUE klass) {
  VALUE obj;
  struct chart *ptr;

  obj = Data_Make_Struct(klass, struct chart, NULL, NULL, ptr);

  ptr->chart = NULL;

  return obj;
}

/* :nodoc: */
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

/* :nodoc: */
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

/* :nodoc: */
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

/*  call-seq: chart.add_series([categories,] values) -> self
 *
 *  Adds data series to the chart.
 */
VALUE
chart_add_series_(int argc, VALUE *argv, VALUE self) {
  rb_check_arity(argc, 0, 2);

  VALUE series = rb_funcall(cChartSeries, rb_intern("new"), argc+1, self, argv[0], argv[1]);
  return series;
}

/*  call-seq: chart.x_axis -> Chart::Axis
 *
 *  Returns x axis as Chart::Axis
 */
VALUE
chart_x_axis_(VALUE self) {
  return rb_funcall(cChartAxis, rb_intern("new"), 2, self, ID2SYM(rb_intern("x")));
}

/*  call-seq: chart.y_axis -> Chart::Axis
 *
 *  Returns y axis as Chart::Axis
 */
VALUE
chart_y_axis_(VALUE self) {
  return rb_funcall(cChartAxis, rb_intern("new"), 2, self, ID2SYM(rb_intern("y")));
}

/*  call-seq: chart.title=(title) -> title
 *
 *  Sets the chart title.
 */
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

/*  call-seq:
 *     chart.set_name_range(name, cell) -> self
 *     chart.set_name_range(name, row, col) -> self
 *
 *  Sets the chart title range (+cell+) and +name+.
 */
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

/*  call-seq: chart.legend_position=(position) -> position
 *
 *  Sets the chart legend position, one of <code>Chart::{LEGEND_NONE, LEGEND_RIGHT,
 *  LEGEND_LEFT, LEGEND_TOP, LEGEND_BOTTOM, LEGEND_OVERLAY_RIGHT,
 *  LEGEND_OVERLAY_LEFT}</code>.
 */
VALUE
chart_legend_set_position_(VALUE self, VALUE pos) {
  struct chart *ptr;
  Data_Get_Struct(self, struct chart, ptr);

  chart_legend_set_position(ptr->chart, NUM2UINT(rb_check_to_int(pos)));
  return pos;
}

/*  call-seq: chart.legend_set_font(options) -> self
 *
 *  Sets legend font from options (name, size, bold, italic, underline, rotation,
 *  color, baseline).
 */
VALUE
chart_legend_set_font_(VALUE self, VALUE opts) {
  struct chart *ptr;
  lxw_chart_font font = val_to_lxw_chart_font(opts);
  Data_Get_Struct(self, struct chart, ptr);

  chart_legend_set_font(ptr->chart, &font);
  return self;
}

/*  call-seq: chart.legend_delete_series(series) -> self
 *
 *  Deletes series by 0-based indexes from the chart legend.
 *
 *     chart.legend_delete_series([0, 2])
 */
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

/*  call-seq: chart.style=(style) -> style
 *
 *  Sets the chart +style+ (integer from 1 to 48, default is 2).
 */
VALUE
chart_set_style_(VALUE self, VALUE style) {
  struct chart *ptr;
  Data_Get_Struct(self, struct chart, ptr);

  chart_set_style(ptr->chart, NUM2INT(rb_check_to_int(style)));
  return style;
}

/* :nodoc: */
VALUE
chart_set_rotation_(VALUE self, VALUE rotation) {
  struct chart *ptr;
  Data_Get_Struct(self, struct chart, ptr);

  chart_set_rotation(ptr->chart, NUM2UINT(rb_check_to_int(rotation)));
  return rotation;
}

/* :nodoc: */
VALUE
chart_set_hole_size_(VALUE self, VALUE size) {
  struct chart *ptr;
  Data_Get_Struct(self, struct chart, ptr);

  chart_set_hole_size(ptr->chart, NUM2UINT(rb_check_to_int(size)));
  return size;
}


/* :nodoc: */
VALUE
chart_get_axis_id_1_(VALUE self) {
  struct chart *ptr;
  Data_Get_Struct(self, struct chart, ptr);
  return UINT2NUM(ptr->chart->axis_id_1);
}

/* :nodoc: */
VALUE
chart_set_axis_id_1_(VALUE self, VALUE val) {
  struct chart *ptr;
  Data_Get_Struct(self, struct chart, ptr);
  ptr->chart->axis_id_1 = NUM2UINT(val);
  return val;
}

/* :nodoc: */
VALUE
chart_get_axis_id_2_(VALUE self) {
  struct chart *ptr;
  Data_Get_Struct(self, struct chart, ptr);
  return UINT2NUM(ptr->chart->axis_id_2);
}

/* :nodoc: */
VALUE
chart_set_axis_id_2_(VALUE self, VALUE val) {
  struct chart *ptr;
  Data_Get_Struct(self, struct chart, ptr);
  ptr->chart->axis_id_2 = NUM2UINT(val);
  return val;
}


/*  call-seq: axis.name=(name) -> name
 *
 *  Sets the chart axis +name+.
 */
VALUE
chart_axis_set_name_(VALUE self, VALUE val) {
  struct chart_axis *ptr;
  Data_Get_Struct(self, struct chart_axis, ptr);

  chart_axis_set_name(ptr->axis, StringValueCStr(val));
  return val;
}

/*  call-seq:
 *     axis.set_name_range(name, cell) -> self
 *     axis.set_name_range(name, row, col) -> self
 *
 *  Sets the chart axis name range from +cell+, with value +name+.
 */
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

/*  call-seq: axis.set_name_font(options) -> self
 *
 *  Same as Chart#set_font, but for axis name.
 */
VALUE
chart_axis_set_name_font_(VALUE self, VALUE value) {
  struct chart_axis *ptr;
  lxw_chart_font font = val_to_lxw_chart_font(value);
  Data_Get_Struct(self, struct chart_axis, ptr);

  chart_axis_set_name_font(ptr->axis, &font);
  return self;
}

/*  call-seq: axis.set_num_font(options) -> self
 *
 *  Same as Chart#set_font, but for axis numbers.
 */
VALUE
chart_axis_set_num_font_(VALUE self, VALUE value) {
  struct chart_axis *ptr;
  lxw_chart_font font = val_to_lxw_chart_font(value);
  Data_Get_Struct(self, struct chart_axis, ptr);

  chart_axis_set_num_font(ptr->axis, &font);
  return self;
}

/*  call-seq: axis.set_line(options)
 *
 *  Sets axis line options. See {libxlsxwriter doc}[https://libxlsxwriter.github.io/structlxw__chart__line.html] for details.
 */
VALUE
chart_axis_set_line_(VALUE self, VALUE opts) {
  struct chart_axis *ptr;
  lxw_chart_line line = val_to_lxw_chart_line(opts);

  Data_Get_Struct(self, struct chart_axis, ptr);

  chart_axis_set_line(ptr->axis, &line);
  return self;
}

/*  call-seq: axis.set_fill(options)
 *
 *  Sets axis fill options. See {libxlsxwriter doc}[https://libxlsxwriter.github.io/structlxw__chart__fill.html] for details.
 */
VALUE
chart_axis_set_fill_(VALUE self, VALUE opts) {
  struct chart_axis *ptr;
  lxw_chart_fill fill = val_to_lxw_chart_fill(opts);

  Data_Get_Struct(self, struct chart_axis, ptr);

  chart_axis_set_fill(ptr->axis, &fill);
  return self;
}

/*  call-seq:
 *     series.set_categories(sheetname, range)
 *     series.set_categories(sheetname, cell_from, cell_to)
 *     series.set_categories(sheetname, row_from, col_from, row_to, col_to)
 *
 *  Sets chart series categories (alternative to first argument of
 *  Chart#add_series).
 */
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

/*  call-seq:
 *     series.set_values(sheetname, range)
 *     series.set_values(sheetname, cell_from, cell_to)
 *     series.set_values(sheetname, row_from, col_from, row_to, col_to)
 *
 *  Sets chart series values (alternative to second argument of Chart#add_series).
 */
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

/*  call-seq: series.name=(name) -> name
 *
 *  Set chart series name.
 */
VALUE
chart_series_set_name_(VALUE self, VALUE name) {
  struct chart_series *ptr;
  Data_Get_Struct(self, struct chart_series, ptr);

  chart_series_set_name(ptr->series, StringValueCStr(name));
  return name;
}

/*  call-seq:
 *     series.set_name_range(name, cell) -> self
 *     series.set_name_range(name, row, col) -> self
 *
 *  Sets the chart series name range from +cell+, with value +name+.
 */
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


void init_xlsxwriter_chart() {
#if 0
  // Making RDoc happy.
  mXlsxWriter = rb_define_module("XlsxWriter");
  cWorkbook = rb_define_class_under(mXlsxWriter, "Workbook", rb_cObject);
#endif
  /* Workbook chart */
  cChart       = rb_define_class_under(cWorkbook, "Chart",  rb_cObject);
  /* Chart data series */
  cChartSeries = rb_define_class_under(cChart,    "Series", rb_cObject);
  /* Chart axis */
  cChartAxis   = rb_define_class_under(cChart,    "Axis",   rb_cObject);


  rb_define_alloc_func(cChart, chart_alloc);
  rb_define_method(cChart, "initialize", chart_init, 2);
  rb_define_method(cChart, "add_series", chart_add_series_, -1);
  rb_define_method(cChart, "x_axis", chart_x_axis_, 0);
  rb_define_method(cChart, "y_axis", chart_y_axis_, 0);
  rb_define_method(cChart, "title=", chart_title_set_name_, 1);
  rb_define_method(cChart, "legend_position=", chart_legend_set_position_, 1);
  rb_define_method(cChart, "legend_set_font", chart_legend_set_font_, 1);
  rb_define_method(cChart, "legend_delete_series", chart_legend_delete_series_, 1);
  rb_define_method(cChart, "style=", chart_set_style_, 1);
  rb_define_method(cChart, "rotation=", chart_set_rotation_, 1);
  rb_define_method(cChart, "hole_size=", chart_set_hole_size_, 1);

  rb_define_method(cChart, "axis_id_1",  chart_get_axis_id_1_, 0);
  rb_define_method(cChart, "axis_id_1=", chart_set_axis_id_1_, 1);
  rb_define_method(cChart, "axis_id_2",  chart_get_axis_id_2_, 0);
  rb_define_method(cChart, "axis_id_2=", chart_set_axis_id_2_, 1);

  rb_define_alloc_func(cChartAxis, chart_axis_alloc);
  rb_define_method(cChartAxis, "initialize", chart_axis_init, 2);

  rb_define_method(cChartAxis, "name=", chart_axis_set_name_, 1);
  rb_define_method(cChartAxis, "set_name_range", chart_axis_set_name_range_, -1);
  rb_define_method(cChartAxis, "set_name_font", chart_axis_set_name_font_, 1);
  rb_define_method(cChartAxis, "set_num_font", chart_axis_set_num_font_, 1);
  rb_define_method(cChartAxis, "set_line", chart_axis_set_line_, 1);
  rb_define_method(cChartAxis, "set_fill", chart_axis_set_fill_, 1);

  rb_define_alloc_func(cChartSeries, chart_series_alloc);
  rb_define_method(cChartSeries, "initialize", chart_series_init, -1);

  rb_define_method(cChartSeries, "set_categories", chart_series_set_categories_, -1);
  rb_define_method(cChartSeries, "set_values", chart_series_set_values_, -1);
  rb_define_method(cChartSeries, "name=", chart_series_set_name_, 1);
  rb_define_method(cChartSeries, "set_name_range", chart_series_set_name_range_, -1);

#define MAP_CHART_CONST(name) rb_define_const(cChart, #name, INT2NUM(LXW_CHART_##name))
  MAP_CHART_CONST(NONE);
  MAP_CHART_CONST(AREA);
  MAP_CHART_CONST(AREA_STACKED);
  MAP_CHART_CONST(AREA_STACKED_PERCENT);
  MAP_CHART_CONST(BAR);
  MAP_CHART_CONST(BAR_STACKED);
  MAP_CHART_CONST(BAR_STACKED_PERCENT);
  MAP_CHART_CONST(COLUMN);
  MAP_CHART_CONST(COLUMN_STACKED);
  MAP_CHART_CONST(COLUMN_STACKED_PERCENT);
  MAP_CHART_CONST(DOUGHNUT);
  MAP_CHART_CONST(LINE);
  MAP_CHART_CONST(PIE);
  MAP_CHART_CONST(SCATTER);
  MAP_CHART_CONST(SCATTER_STRAIGHT);
  MAP_CHART_CONST(SCATTER_STRAIGHT_WITH_MARKERS);
  MAP_CHART_CONST(SCATTER_SMOOTH);
  MAP_CHART_CONST(SCATTER_SMOOTH_WITH_MARKERS);
  MAP_CHART_CONST(RADAR);
  MAP_CHART_CONST(RADAR_WITH_MARKERS);
  MAP_CHART_CONST(RADAR_FILLED);

  MAP_CHART_CONST(LEGEND_NONE);
  MAP_CHART_CONST(LEGEND_RIGHT);
  MAP_CHART_CONST(LEGEND_LEFT);
  MAP_CHART_CONST(LEGEND_TOP);
  MAP_CHART_CONST(LEGEND_BOTTOM);
  MAP_CHART_CONST(LEGEND_TOP_RIGHT);
  MAP_CHART_CONST(LEGEND_OVERLAY_RIGHT);
  MAP_CHART_CONST(LEGEND_OVERLAY_LEFT);
  MAP_CHART_CONST(LEGEND_OVERLAY_TOP_RIGHT);

  MAP_CHART_CONST(AXIS_LABEL_ALIGN_CENTER);
  MAP_CHART_CONST(AXIS_LABEL_ALIGN_LEFT);
  MAP_CHART_CONST(AXIS_LABEL_ALIGN_RIGHT);


#undef MAP_CHART_CONST
}
