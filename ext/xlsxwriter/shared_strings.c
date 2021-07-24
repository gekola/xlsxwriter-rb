#include <ruby.h>
#include <xlsxwriter.h>

#include "xlsxwriter_ext.h"
#include "shared_strings.h"

struct sst {
  lxw_sst *sst;
};

VALUE cSharedStringsTable;

void sst_free(void *p);

/* :nodoc: */
VALUE
sst_alloc(VALUE klass) {
  VALUE obj;
  struct sst *sst_ptr;
  obj = Data_Make_Struct(klass, struct sst, NULL, sst_free, sst_ptr);
  sst_ptr->sst = NULL;

  return obj;
}

VALUE
alloc_shared_strings_table_by_ref(struct lxw_sst *sst) {
  if (!sst)
    return Qnil;

  VALUE object = rb_obj_alloc(cSharedStringsTable);
  struct sst *sst_ptr;
  Data_Get_Struct(object, struct sst, sst_ptr);
  sst_ptr->sst = sst;
  return object;
}

void
sst_free(void *p) {
}

VALUE
sst_string_count_get_(VALUE self) {
  struct sst *sst_ptr;
  Data_Get_Struct(self, struct sst, sst_ptr);
  return INT2NUM(sst_ptr->sst->string_count);
}

VALUE
sst_string_count_set_(VALUE self, VALUE val) {
  struct sst *sst_ptr;
  Data_Get_Struct(self, struct sst, sst_ptr);

  sst_ptr->sst->string_count = NUM2INT(val);
  return val;
}

void
init_xlsxwriter_shared_strings_table() {
  cSharedStringsTable = rb_define_class_under(mXlsxWriter, "SharedStringsTable", rb_cObject);

  rb_define_alloc_func(cSharedStringsTable, sst_alloc);
  rb_define_method(cSharedStringsTable, "string_count", sst_string_count_get_, 0);
  rb_define_method(cSharedStringsTable, "string_count=", sst_string_count_set_, 1);
}
