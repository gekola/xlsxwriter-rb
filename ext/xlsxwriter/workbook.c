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
  ptr->properties = NULL;

  return obj;
}

VALUE
workbook_new_(int argc, VALUE *argv, VALUE self) {
  VALUE workbook = rb_call_super(argc, argv);
  if (rb_block_given_p()) {
    rb_yield(workbook);
    workbook_release(workbook);
    return Qnil;
  }
  return workbook;
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
  ptr->properties = NULL;
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
    if (ptr->properties) {
      workbook_set_properties(ptr->workbook, ptr->properties);
    }
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

  if (ptr->properties) {
#define FREE_PROP(prop) {              \
      if (ptr->properties->prop) {     \
        xfree(ptr->properties->prop);  \
        ptr->properties->prop = NULL;  \
      }                                \
    }
    FREE_PROP(title);
    FREE_PROP(subject);
    FREE_PROP(author);
    FREE_PROP(manager);
    FREE_PROP(company);
    FREE_PROP(category);
    FREE_PROP(keywords);
    FREE_PROP(comments);
    FREE_PROP(status);
    FREE_PROP(hyperlink_base);
#undef FREE_PROP
    free(ptr->properties);
    ptr->properties = NULL;
  }
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

  if (rb_block_given_p()) {
    VALUE res = rb_yield(worksheet);
    return res;
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

VALUE
workbook_set_default_xf_indices_(VALUE self) {
  struct workbook *ptr;
  Data_Get_Struct(self, struct workbook, ptr);
  lxw_workbook_set_default_xf_indices(ptr->workbook);
  return self;
}

VALUE
workbook_properties_(VALUE self) {
  VALUE props = rb_obj_alloc(cWorkbookProperties);
  rb_obj_call_init(props, 1, &self);
  return props;
}

VALUE
workbook_define_name_(VALUE self, VALUE name, VALUE formula) {
  struct workbook *ptr;
  Data_Get_Struct(self, struct workbook, ptr);
  workbook_define_name(ptr->workbook, StringValueCStr(name), StringValueCStr(formula));
  return self;
}

VALUE
workbook_validate_worksheet_name_(VALUE self, VALUE name) {
  struct workbook *ptr;
  lxw_error err;
  Data_Get_Struct(self, struct workbook, ptr);
  err = workbook_validate_worksheet_name(ptr->workbook, StringValueCStr(name));
  handle_lxw_error(err);
  return Qtrue;
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

lxw_datetime
value_to_lxw_datetime(VALUE val) {
  const ID i_to_time = rb_intern("to_time");
  if (rb_respond_to(val, i_to_time)) {
    val = rb_funcall(val, i_to_time, 0);
  }
  lxw_datetime res = {
    .year  = NUM2INT(rb_funcall(val, rb_intern("year"), 0)),
    .month = NUM2INT(rb_funcall(val, rb_intern("month"), 0)),
    .day   = NUM2INT(rb_funcall(val, rb_intern("day"), 0)),
    .hour  = NUM2INT(rb_funcall(val, rb_intern("hour"), 0)),
    .min   = NUM2INT(rb_funcall(val, rb_intern("min"), 0)),
    .sec   = NUM2DBL(rb_funcall(val, rb_intern("sec"), 0)) +
    NUM2DBL(rb_funcall(val, rb_intern("subsec"), 0))
  };
  return res;
}

void
handle_lxw_error(lxw_error err) {
  
}
