#include <ruby.h>

#ifndef __WORKBOOK__
#define __WORKBOOK__

struct workbook {
  char *path;
  lxw_workbook *workbook;
  lxw_doc_properties *properties;
  st_table *formats;
};

lxw_format *workbook_get_format(VALUE self, VALUE key);
lxw_datetime value_to_lxw_datetime(VALUE val);

void init_xlsxwriter_workbook();

extern VALUE mXlsxWriter;
extern VALUE cWorkbook;
extern const rb_data_type_t workbook_type;

#endif // __WORKBOOK__
