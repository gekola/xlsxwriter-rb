#include <string.h>
#include <ruby/util.h>
#include "xlsxwriter.h"
#include "workbook.h"
#include "workbook_properties.h"

VALUE
workbook_properties_init_(VALUE self, VALUE workbook) {
  struct workbook *wb_ptr;
  rb_iv_set(self, "@workbook", workbook);
  Data_Get_Struct(workbook, struct workbook, wb_ptr);
  if (!wb_ptr->properties) {
    wb_ptr->properties = calloc(1, sizeof(lxw_doc_properties));
  }
  if (rb_block_given_p()) {
    rb_yield(self);
  }
  return self;
}

VALUE
workbook_properties_set_dir_(VALUE self, VALUE value) {
  VALUE key = rb_id2str(rb_frame_callee());
  char *key_str = ruby_strdup(RSTRING_PTR(key));
  size_t last_pos = RSTRING_LEN(key) - 1;
  if (last_pos > 0 && key_str[last_pos] == '=') {
    key_str[last_pos] = '\0';
  }
  key = rb_str_new_cstr(key_str);
  workbook_properties_set_(self, key, value);
  xfree(key_str);
  return value;
}

VALUE
workbook_properties_set_(VALUE self, VALUE key, VALUE value) {
  char *key_cstr = NULL;
  switch (TYPE(key)) {
  case T_STRING:
    key_cstr = StringValueCStr(key);
    break;
  case T_SYMBOL:
    key = rb_sym2str(key);
    key_cstr = StringValueCStr(key);
    break;
  default:
    rb_raise(rb_eArgError, "Wrong type of key: %"PRIsVALUE, rb_obj_class(key));
  }

  struct workbook *wb_ptr;
  Data_Get_Struct(rb_iv_get(self, "@workbook"), struct workbook, wb_ptr);
  lxw_doc_properties *props = wb_ptr->properties;
  if (!props) {
    rb_raise(rb_eRuntimeError, "Workbook properties are already freed.");
  }

#define HANDLE_PROP(prop) {                              \
    if (!strcmp(#prop, key_cstr)) {                      \
      props->prop = ruby_strdup(StringValueCStr(value)); \
      return value;                                      \
    }                                                    \
  }
  HANDLE_PROP(title);
  HANDLE_PROP(subject);
  HANDLE_PROP(author);
  HANDLE_PROP(manager);
  HANDLE_PROP(company);
  HANDLE_PROP(category);
  HANDLE_PROP(keywords);
  HANDLE_PROP(comments);
  HANDLE_PROP(status);
  HANDLE_PROP(hyperlink_base);
#undef HANDLE_PROP

  // Not a standard property.
  switch (TYPE(value)) {
  case T_NIL:
    break;
  case T_STRING:
    workbook_set_custom_property_string(wb_ptr->workbook, key_cstr, StringValueCStr(value));
    break;
  case T_FLOAT: case T_RATIONAL:
    workbook_set_custom_property_number(wb_ptr->workbook, key_cstr, NUM2DBL(value));
    break;
  case T_FIXNUM: case T_BIGNUM:
    workbook_set_custom_property_integer(wb_ptr->workbook, key_cstr, NUM2INT(value));
    break;
  case T_TRUE: case T_FALSE:
    workbook_set_custom_property_boolean(wb_ptr->workbook, key_cstr, value != Qfalse);
    break;
  case T_DATA:
    if (rb_obj_class(value) == rb_cTime) {
      lxw_datetime datetime = value_to_lxw_datetime(value);
      workbook_set_custom_property_datetime(wb_ptr->workbook, key_cstr, &datetime);
    } else {
      rb_raise(rb_eTypeError, "wrong argument type %"PRIsVALUE
               " (must be string, numeric, true, false, nil or time)",
               rb_obj_class(value));
    }
    break;
  }
  return value;
}
