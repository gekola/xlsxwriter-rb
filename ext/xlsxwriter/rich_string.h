#include <ruby.h>
#include <xlsxwriter.h>

#ifndef __RICH_STRING__
#define __RICH_STRING__

void init_xlsxwriter_rich_string();
lxw_rich_string_tuple **serialize_rich_string(VALUE rs);

extern VALUE cRichString;

#endif /* RICH_STRING_H */
