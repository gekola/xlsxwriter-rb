#include <ruby.h>
#include <xlsxwriter.h>

#ifndef __CHARTSHEET__
#define __CHARTSHEET__

struct chartsheet {
  lxw_chartsheet *chartsheet;
};

void init_xlsxwriter_chartsheet();

extern VALUE cChartsheet;

#endif // __CHARTSHEET__
