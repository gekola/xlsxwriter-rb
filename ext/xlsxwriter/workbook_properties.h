#include <ruby.h>

#ifndef __WORKBOOK_PROPERTIES__
#define __WORKBOOK_PROPIRTIES__

VALUE workbook_properties_init_(VALUE, VALUE);
VALUE workbook_properties_set_dir_(VALUE self, VALUE value);
VALUE workbook_properties_set_(VALUE self, VALUE key, VALUE value);

#endif // __WORKBOOK_PROPIRTIES__
