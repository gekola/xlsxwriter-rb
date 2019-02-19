#include <ruby.h>
#include "xlsxwriter.h"
#include "chart.h"
#include "format.h"
#include "workbook.h"
#include "workbook_properties.h"
#include "worksheet.h"

VALUE mXlsxWriter;
VALUE rbLibVersion;

/*  Document-module: XlsxWriter
 *
 *  XlsxWriter is a ruby interface to libxlsxwriter.
 *
 *  It provides a couple of useful shorthands (like being able to pass cells and
 *  ranges as both numbers, cell strings and range strings.
 *
 *  It also has column authowidth functionality partially taken from Axlsx gem
 *  enabled by default.
 *
 *  Simple example of using the XlsxWriter to generate an xlsx file containing
 *  'Hello' string in the first rows of column 'A':
 *
 *     XlsxWriter::Workbook.open('/tmp/text.xlsx') do |wb|
 *       ws.add_worksheet do |ws|
 *         10.times { |i| ws.add_row ['Hello!'] }
 *       end
 *     end
 */
void Init_xlsxwriter() {
  mXlsxWriter = rb_define_module("XlsxWriter");
  rbLibVersion = rb_str_new_cstr(lxw_version());
  rb_define_const(mXlsxWriter, "LIBRARY_VERSION", rbLibVersion);

  init_xlsxwriter_workbook();
  init_xlsxwriter_workbook_properties();
  init_xlsxwriter_format();
  init_xlsxwriter_worksheet();
  init_xlsxwriter_chart();
}
