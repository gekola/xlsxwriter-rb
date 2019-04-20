#include <ruby.h>
#include <xlsxwriter.h>

#ifndef __FORMAT__
#define __FORMAT__

void format_apply_opts(lxw_format *format, VALUE opts);

void init_xlsxwriter_format();

extern VALUE mXlsxFormat;

#endif // __FORMAT__
