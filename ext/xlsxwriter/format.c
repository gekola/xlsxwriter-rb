#include "format.h"

#define STR2SYM(x) ID2SYM(rb_intern(x))

#define MAP_FORMAT(fmt_name, ...) {                 \
    value = rb_hash_aref(opts, STR2SYM(#fmt_name)); \
    if (!NIL_P(value))                              \
      format_set_##fmt_name(format, __VA_ARGS__);   \
  }

#define MAP_FORMATB(fmt_name) {                     \
    value = rb_hash_aref(opts, STR2SYM(#fmt_name)); \
    if (!NIL_P(value) && value)                     \
      format_set_##fmt_name(format);                \
}

void format_apply_opts(lxw_format *format, VALUE opts) {
  VALUE value;

  MAP_FORMAT(font_name, StringValueCStr(value));
  MAP_FORMAT(font_size, NUM2INT(value))
  MAP_FORMAT(font_color, NUM2LONG(value));
  MAP_FORMATB(bold);
  MAP_FORMATB(italic);
  MAP_FORMAT(underline, NUM2LONG(value));
  MAP_FORMATB(font_strikeout);
  MAP_FORMAT(font_script, NUM2INT(value));
  MAP_FORMAT(num_format, StringValueCStr(value));
  MAP_FORMAT(num_format_index, NUM2INT(value));
  MAP_FORMATB(unlocked);
  MAP_FORMATB(hidden);
  MAP_FORMAT(align, NUM2INT(value));

  value = rb_hash_aref(opts, STR2SYM("vertical_align"));
  if (value != Qnil)
    format_set_align(format, NUM2INT(value));

  MAP_FORMATB(text_wrap);
  MAP_FORMAT(rotation, NUM2INT(value));
  MAP_FORMAT(indent, NUM2INT(value));
  MAP_FORMATB(shrink);
  MAP_FORMAT(pattern, NUM2INT(value));
  MAP_FORMAT(bg_color, NUM2LONG(value));
  MAP_FORMAT(fg_color, NUM2LONG(value));
  MAP_FORMAT(border, NUM2INT(value));
  MAP_FORMAT(bottom, NUM2INT(value));
  MAP_FORMAT(top, NUM2INT(value));
  MAP_FORMAT(left, NUM2INT(value));
  MAP_FORMAT(right, NUM2INT(value));
  MAP_FORMAT(border_color, NUM2LONG(value));
  MAP_FORMAT(bottom_color, NUM2LONG(value));
  MAP_FORMAT(top_color, NUM2LONG(value));
  MAP_FORMAT(left_color, NUM2LONG(value));
  MAP_FORMAT(right_color, NUM2LONG(value));
  MAP_FORMAT(diag_type, NUM2INT(value));
  MAP_FORMAT(diag_color, NUM2LONG(value));
  MAP_FORMAT(diag_border, NUM2INT(value));
  MAP_FORMATB(font_outline);
  MAP_FORMATB(font_shadow);
  MAP_FORMAT(font_family, NUM2INT(value));
  MAP_FORMAT(font_charset, NUM2INT(value));
  MAP_FORMAT(font_scheme, StringValueCStr(value));
  MAP_FORMATB(font_condense);
  MAP_FORMATB(font_extend);
  MAP_FORMAT(reading_order, NUM2INT(value));
  MAP_FORMAT(theme, NUM2INT(value));
}
