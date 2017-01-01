#include <ruby.h>

#ifndef __WORKBOOK__
#define __WORKBOOK__

struct workbook {
  char *path;
  lxw_workbook *workbook;
  lxw_doc_properties *properties;
  st_table *formats;
};


VALUE workbook_alloc(VALUE klass);
VALUE workbook_new_(int argc, VALUE *argv, VALUE self);
VALUE workbook_init(int argc, VALUE *argv, VALUE self);
VALUE workbook_release(VALUE self);
void workbook_free(void *);
VALUE workbook_add_worksheet_(int argc, VALUE *argv, VALUE self);
VALUE workbook_add_format_(VALUE self, VALUE key, VALUE opts);
VALUE workbook_set_default_xf_indices_(VALUE self);
VALUE workbook_properties_(VALUE self);
VALUE workbook_define_name_(VALUE self, VALUE name, VALUE formula);
VALUE workbook_validate_worksheet_name_(VALUE self, VALUE name);


lxw_format *workbook_get_format(VALUE self, VALUE key);
lxw_datetime value_to_lxw_datetime(VALUE val);
void handle_lxw_error(lxw_error err);

extern VALUE mXlsxWriter;
extern VALUE cWorkbook;
extern VALUE cWorksheet;
extern VALUE mXlsxFormat;
extern VALUE cWorkbookProperties;

#endif // __WORKBOOK__
