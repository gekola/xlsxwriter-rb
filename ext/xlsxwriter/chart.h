#include <ruby.h>
#include "xlsxwriter.h"

#ifndef __CHART__
#define __CHART__

struct chart {
  lxw_chart *chart;
};

struct chart_series {
  lxw_chart_series *series;
};

VALUE chart_alloc(VALUE klass);
VALUE chart_init(VALUE self, VALUE worksheet, VALUE type);
VALUE chart_series_alloc(VALUE klass);
VALUE chart_series_init(int argc, VALUE *argv, VALUE self);

VALUE chart_add_series_(int argc, VALUE *argv, VALUE self);

VALUE chart_get_axis_id_1_(VALUE self);
VALUE chart_set_axis_id_1_(VALUE self, VALUE val);
VALUE chart_get_axis_id_2_(VALUE self);
VALUE chart_set_axis_id_2_(VALUE self, VALUE val);

#endif // __CHART__
