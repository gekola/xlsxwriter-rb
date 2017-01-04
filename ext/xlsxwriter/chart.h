#include <ruby.h>
#include "xlsxwriter.h"

#ifndef __CHART__
#define __CHART__

struct chart {
  lxw_chart *chart;
};

struct chart_axis {
  lxw_chart_axis *axis;
};

struct chart_series {
  lxw_chart_series *series;
};

VALUE chart_alloc(VALUE klass);
VALUE chart_init(VALUE self, VALUE worksheet, VALUE type);
VALUE chart_axis_alloc(VALUE klass);
VALUE chart_axis_init(VALUE self, VALUE chart, VALUE type);
VALUE chart_series_alloc(VALUE klass);
VALUE chart_series_init(int argc, VALUE *argv, VALUE self);

VALUE chart_add_series_(int argc, VALUE *argv, VALUE self);
VALUE chart_x_axis_(VALUE self);
VALUE chart_y_axis_(VALUE self);
VALUE chart_title_set_name_(VALUE self, VALUE name);
VALUE chart_title_set_name_range_(int argc, VALUE *argv, VALUE self);
VALUE chart_legend_set_position_(VALUE self, VALUE pos);
VALUE chart_legend_set_font_(VALUE self, VALUE opts);
VALUE chart_legend_delete_series_(VALUE self, VALUE series);
VALUE chart_set_style_(VALUE self, VALUE style);
VALUE chart_set_rotation_(VALUE self, VALUE rotation);
VALUE chart_set_hole_size_(VALUE self, VALUE size);

VALUE chart_get_axis_id_1_(VALUE self);
VALUE chart_set_axis_id_1_(VALUE self, VALUE val);
VALUE chart_get_axis_id_2_(VALUE self);
VALUE chart_set_axis_id_2_(VALUE self, VALUE val);

VALUE chart_axis_set_name_(VALUE self, VALUE val);
VALUE chart_axis_set_name_range_(int argc, VALUE *argv, VALUE self);
VALUE chart_axis_set_name_font_(VALUE self, VALUE val);
VALUE chart_axis_set_num_font_(VALUE self, VALUE val);
VALUE chart_axis_set_line_(VALUE self, VALUE opts);
VALUE chart_axis_set_fill_(VALUE self, VALUE opts);

VALUE chart_series_set_categories_(int argc, VALUE *argv, VALUE self);
VALUE chart_series_set_values_(int argc, VALUE *argv, VALUE self);
VALUE chart_series_set_name_(VALUE self, VALUE name);
VALUE chart_series_set_name_range_(int argc, VALUE *argv, VALUE self);


lxw_chart_fill val_to_lxw_chart_fill(VALUE opts);
lxw_chart_font val_to_lxw_chart_font(VALUE opts);
lxw_chart_line val_to_lxw_chart_line(VALUE opts);

#endif // __CHART__
