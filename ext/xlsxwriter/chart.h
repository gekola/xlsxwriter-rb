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

lxw_chart_fill val_to_lxw_chart_fill(VALUE opts);
lxw_chart_font val_to_lxw_chart_font(VALUE opts);
lxw_chart_line val_to_lxw_chart_line(VALUE opts);

void init_xlsxwriter_chart();

extern VALUE cChart;
extern VALUE cChartSeries;
extern VALUE cChartAxis;

#endif // __CHART__
