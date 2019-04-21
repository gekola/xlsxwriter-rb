#include <ruby.h>

#ifndef XLSXWRITER_H
#define XLSXWRITER_H

extern VALUE mXlsxWriter;
extern VALUE eXlsxWriterError;

static inline void handle_lxw_error(lxw_error err) {
  if (err) {
    rb_exc_raise(rb_funcall(eXlsxWriterError, rb_intern("new"), 2, rb_str_new_cstr(lxw_strerror(err)), INT2NUM(err)));
  }
}

#endif /* XLSXWRITER_H */
