#include <ruby.h>
#include <xlsxwriter.h>

#ifndef __SHARED_STRINGS__
#define __SHARED_STRINGS__

void
init_xlsxwriter_shared_strings_table();

extern VALUE cSharedStringsTable;

VALUE
alloc_shared_strings_table_by_ref(struct lxw_sst *);

#endif // __SHARED_STRINGS__
