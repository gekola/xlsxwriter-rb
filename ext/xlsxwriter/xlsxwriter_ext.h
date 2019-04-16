#include <ruby.h>

#ifndef XLSXWRITER_H
#define XLSXWRITER_H

extern VALUE mXlsxWriter;
extern VALUE eXlsxWriterError;

inline void handle_xlsxwriter_error(lxw_error err) {
  if (err) {
    rb_raise(eXlsxWriterError, lxw_strerror(err));
  }
}

#endif /* XLSXWRITER_H */
