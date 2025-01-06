#include <ruby.h>
#include <ruby/intern.h>
#include "xlsxwriter_ext.h"

#ifndef __COMMON__
#define __COMMON__

#define LXW_NO_RESULT_CALL(cls, method, ...) {  \
  struct cls *ptr;                              \
  TypedData_Get_Struct(self, struct cls, &(cls ## _type), ptr); \
  cls ## _ ## method(ptr->cls, ##__VA_ARGS__); }

#define LXW_ERR_RESULT_CALL(cls, method, ...) {      \
    struct cls *ptr;                                \
    TypedData_Get_Struct(self, struct cls, &(cls ## _type), ptr);             \
    lxw_error err = cls ## _ ## method(ptr->cls, ##__VA_ARGS__); \
    handle_lxw_error(err); }


struct _protect_options {
  uint8_t with_options;
  const char *password;
  lxw_protection options;
};

static inline struct _protect_options
_extract_protect_options(int argc, VALUE *argv) {
  VALUE val;
  VALUE opts = Qnil;
  struct _protect_options po = {0};
  // All options are off by default
  if (argc > 0 && !NIL_P(argv[0])) {
    switch (TYPE(argv[0])) {
    case T_STRING:
      po.password = StringValueCStr(argv[0]);
      break;
    case T_HASH:
      opts = argv[0];
      break;
    default:
      rb_raise(rb_eArgError, "Wrong argument %" PRIsVALUE ", expected a String or Hash", rb_obj_class(argv[0]));
    }
  }

  if (argc > 1) {
    if (TYPE(argv[1]) == T_HASH) {
      opts = argv[1];
    } else {
      rb_raise(rb_eArgError, "Expected a Hash, but got %" PRIsVALUE, rb_obj_class(argv[1]));
    }
  }

  if (!NIL_P(opts)) {
#define PR_OPT(field) {                                    \
      val = rb_hash_aref(opts, ID2SYM(rb_intern(#field))); \
      if (!NIL_P(val) && val) {                            \
        po.options.field = 1;                              \
        po.with_options = 1;                               \
      }                                                    \
    }
    PR_OPT(no_select_locked_cells);
    PR_OPT(no_select_unlocked_cells);
    PR_OPT(format_cells);
    PR_OPT(format_columns);
    PR_OPT(format_rows);
    PR_OPT(insert_columns);
    PR_OPT(insert_rows);
    PR_OPT(insert_hyperlinks);
    PR_OPT(delete_columns);
    PR_OPT(delete_rows);
    PR_OPT(sort);
    PR_OPT(autofilter);
    PR_OPT(pivot_tables);
    PR_OPT(scenarios);
    PR_OPT(objects);
#undef PR_OPT
  }

  return po;
}


struct _header_options {
  char with_options;
  const char *str;
  lxw_header_footer_options options;
};

static inline struct _header_options
_extract_header_options(int argc, VALUE *argv) {
  struct _header_options ho = { 0 };
  ho.str = StringValueCStr(argv[0]);
  if (argc > 1 && !NIL_P(argv[1])) {
    VALUE margin = rb_hash_aref(argv[1], ID2SYM(rb_intern("margin")));
    if (!NIL_P(margin)) {
      ho.with_options = '\1';
      ho.options.margin = NUM2DBL(margin);
    }
  }
  return ho;
}


#endif // __COMMON__
