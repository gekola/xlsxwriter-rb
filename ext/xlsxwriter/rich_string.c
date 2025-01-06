#include <ruby/util.h>
#include "rich_string.h"
#include "workbook.h"
#include "xlsxwriter_ext.h"

VALUE cRichString;

void rich_string_clear(void *p);
void rich_string_free(void *p);

size_t
rich_string_size(const void *data) {
  return sizeof(lxw_rich_string_tuple *);
}

const rb_data_type_t rich_string_type = {
    .wrap_struct_name = "lxw_rich_string_tuple",
    .function =
        {
            .dmark = NULL,
            .dfree = rich_string_free,
            .dsize = rich_string_size,
        },
    .data = NULL,
    .flags = RUBY_TYPED_FREE_IMMEDIATELY,
};

VALUE
rich_string_alloc(VALUE klass) {
  VALUE obj;
  lxw_rich_string_tuple **ptr;

  obj = TypedData_Make_Struct(klass, lxw_rich_string_tuple *, &rich_string_type, ptr);
  *ptr = NULL;

  return obj;
}

void
rich_string_clear(void *p) {
  lxw_rich_string_tuple **ptr = p;
  if (*ptr) {
    ruby_xfree(*ptr);
    *ptr = NULL;
  }
}

void
rich_string_free(void *p) {
  rich_string_clear(p);
  ruby_xfree(p);
}

VALUE
rich_string_cached_p(VALUE self) {
  lxw_rich_string_tuple **ptr;
  TypedData_Get_Struct(self, lxw_rich_string_tuple *, &rich_string_type, ptr);

  return (*ptr ? Qtrue : Qfalse);
}

lxw_rich_string_tuple **
serialize_rich_string(VALUE rs) {
  VALUE arr = rb_funcall(rs, rb_intern("parts"), 0);
  VALUE wb = rb_funcall(rs, rb_intern("workbook"), 0);

  int len = RARRAY_LEN(arr);

  lxw_rich_string_tuple **ptr;
  TypedData_Get_Struct(rs, lxw_rich_string_tuple *, &rich_string_type, ptr);

  if (*ptr) { // cached
    return (lxw_rich_string_tuple **) *ptr;
  }

  rb_obj_freeze(arr);
  *ptr = ruby_xmalloc(sizeof(lxw_rich_string_tuple) * len + sizeof(lxw_rich_string_tuple *) * (len + 1));
  lxw_rich_string_tuple **res = *(lxw_rich_string_tuple ***)ptr;

  lxw_rich_string_tuple *base_ptr = ((void *)*ptr + sizeof(lxw_rich_string_tuple *) * (len + 1));
  int i;
  for (i = 0; i < len; ++i) {
    VALUE part = rb_ary_entry(arr, i);
    VALUE text = rb_ary_entry(part, 0);
    VALUE format = rb_ary_entry(part, 1);

    base_ptr[i].format = workbook_get_format(wb, format);
    base_ptr[i].string = StringValueCStr(text);
    res[i] = &base_ptr[i];
  }
  res[len] = NULL;

  return res;
}


void init_xlsxwriter_rich_string() {
#if 0
  // Making RDoc happy.
  mXlsxWriter = rb_define_module("XlsxWriter");
#endif
  cRichString = rb_define_class_under(mXlsxWriter, "RichString", rb_cObject);
  rb_define_alloc_func(cRichString, rich_string_alloc);
  rb_define_private_method(cRichString, "cached?", rich_string_cached_p, 0);
}
