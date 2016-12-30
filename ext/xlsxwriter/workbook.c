#include <string.h>
#include "xlsxwriter.h"
#include "format.h"
#include "workbook.h"

VALUE
workbook_alloc(VALUE klass)
{
  VALUE obj;
  struct workbook *ptr;

  obj = Data_Make_Struct(klass, struct workbook, NULL, workbook_free, ptr);

  ptr->path = NULL;
  ptr->workbook = NULL;
  ptr->formats = NULL;

  return obj;
}

VALUE
workbook_init(int argc, VALUE *argv, VALUE self) {
  struct workbook *ptr;
  lxw_workbook_options options = {
    .constant_memory = 0,
    .tmpdir = NULL
  };

  if (argc < 1 || argc > 2) {
    rb_raise(rb_eArgError, "wrong number of arguments");
    return self;
  } else if (argc == 2) {
    VALUE const_mem = rb_hash_aref(argv[1], ID2SYM(rb_intern("constant_memory")));
    if (!NIL_P(const_mem) && const_mem) {
      options.constant_memory = 1;
      VALUE tmpdir = rb_hash_aref(argv[1], ID2SYM(rb_intern("tmpdir")));
      if (!NIL_P(tmpdir))
        options.tmpdir = RSTRING_PTR(tmpdir);
    }
  }

  Data_Get_Struct(self, struct workbook, ptr);

  size_t len = RSTRING_LEN(argv[0]);
  ptr->path = malloc(len + 1);
  strncpy(ptr->path, RSTRING_PTR(argv[0]), len + 1);
  if (options.constant_memory) {
    ptr->workbook = workbook_new_opt(ptr->path, &options);
  } else {
    ptr->workbook = workbook_new(ptr->path);
  }
  rb_iv_set(self, "@font_sizes", rb_hash_new());

  return self;
}

VALUE
workbook_release(VALUE self) {
  struct workbook *ptr;
  Data_Get_Struct(self, struct workbook, ptr);

  workbook_free(ptr);
  return self;
}

void
workbook_free(void *p) {
  struct workbook *ptr = p;

  if (ptr->workbook) {
    workbook_close(ptr->workbook);
    ptr->workbook = NULL;
  }
  if (ptr->path) {
    free(ptr->path);
    ptr->path = NULL;
  }
  if (ptr->formats) {
    st_free_table(ptr->formats);
    ptr->formats = NULL;
  };
}

VALUE
workbook_add_worksheet_(int argc, VALUE *argv, VALUE self) {
  VALUE worksheet = Qnil;

  if (argc > 1) {
    rb_raise(rb_eArgError, "wrong number of arguments");
    return self;
  }

  struct workbook *ptr;
  Data_Get_Struct(self, struct workbook, ptr);
  if (ptr->workbook) {
    VALUE mXlsxWriter = rb_const_get(rb_cObject, rb_intern("XlsxWriter"));
    VALUE cWorksheet = rb_const_get(mXlsxWriter, rb_intern("Worksheet"));
    worksheet = rb_funcall(cWorksheet, rb_intern("new"), argc + 1, self, argv[0]);
  }

  return worksheet;
}

VALUE
workbook_add_format_(VALUE self, VALUE key, VALUE opts) {
  struct workbook *ptr;
  lxw_format *format;
  Data_Get_Struct(self, struct workbook, ptr);

  if (!ptr->formats) {
    ptr->formats = st_init_numtable();
  }

  format = workbook_add_format(ptr->workbook);
  st_insert(ptr->formats, rb_to_id(key), (st_data_t)format);
  format_apply_opts(format, opts);

  VALUE font_size = rb_hash_aref(opts, ID2SYM(rb_intern("font_size")));
  if (!NIL_P(font_size)) {
    VALUE bold = rb_hash_aref(opts, ID2SYM(rb_intern("bold")));
    if (!NIL_P(bold) && bold) {
      rb_hash_aset(rb_iv_get(self, "@font_sizes"), key, rb_float_new(NUM2DBL(font_size) * 1.5));
    } else {
      rb_hash_aset(rb_iv_get(self, "@font_sizes"), key, font_size);
    }
  }

  return self;
}


VALUE workbook_set_default_xf_indices_(VALUE self) {
  struct workbook *ptr;
  Data_Get_Struct(self, struct workbook, ptr);
  lxw_workbook_set_default_xf_indices(ptr->workbook);
  return self;
}


lxw_format *
workbook_get_format(VALUE self, VALUE key) {
  struct workbook *ptr;
  lxw_format *format = NULL;

  if (NIL_P(key))
    return NULL;

  Data_Get_Struct(self, struct workbook, ptr);

  if (!ptr->formats)
    return NULL;

  st_lookup(ptr->formats, rb_to_id(key), (st_data_t *) &format);

  return format;
}
