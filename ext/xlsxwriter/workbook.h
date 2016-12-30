#include <ruby.h>

#ifndef __WORKBOOK__
#define __WORKBOOK__

struct workbook {
  char *path;
  lxw_workbook *workbook;
  st_table *formats;
};


VALUE workbook_alloc(VALUE klass);
VALUE workbook_init(int argc, VALUE *argv, VALUE self);
VALUE workbook_release(VALUE self);
void workbook_free(void *);
VALUE workbook_add_worksheet_(int argc, VALUE *argv, VALUE self);
VALUE workbook_add_format_(VALUE self, VALUE key, VALUE opts);
VALUE workbook_set_default_xf_indices_(VALUE self);


lxw_format *workbook_get_format(VALUE self, VALUE key);

#endif // __WORKBOOK__
