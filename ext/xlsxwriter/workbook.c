#include <string.h>
#include <ruby.h>
#include <ruby/thread.h>
#include "xlsxwriter.h"
#include "chart.h"
#include "format.h"
#include "workbook.h"
#include "workbook_properties.h"
#include "worksheet.h"

VALUE cWorkbook;

void workbook_free(void *p);
VALUE workbook_release(VALUE self);


VALUE
workbook_alloc(VALUE klass) {
  VALUE obj;
  struct workbook *ptr;

  obj = Data_Make_Struct(klass, struct workbook, NULL, workbook_free, ptr);

  ptr->path = NULL;
  ptr->workbook = NULL;
  ptr->formats = NULL;
  ptr->properties = NULL;

  return obj;
}

/*  call-seq:
 *     XlsxWriter::Workbook.new(path, constant_memory: false, tmpdir: nil) -> workbook
 *     XlsxWriter::Workbook.new(path, constant_memory: false, tmpdir: nil) { |wb| block } -> nil
 *     XlsxWriter::Workbook.open(path, constant_memory: false, tmpdir: nil) { |wb| block } -> nil
 *
 *  Creates a new Xlsx workbook in file +path+ and returns a new Workbook object.
 *
 *  If +constant_memory+ is set to true workbook data is stored in temporary files
 *  in +tmpdir+, considerably reducing memory consumption for large documents.
 *
 *    XlsxWriter::Workbook.open('/tmp/test.xlsx', constant_memory: true) do |wb|
 *      # ... populate the workbook with data ...
 *    end
 */
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

/* :nodoc: */
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

/*  call-seq:
 *     wb.close -> nil
 *
 *  Dumps the workbook content to the file and closes the worksheet.
 *  To be used only for workbooks opened with <code>XlsxWriter::Workbook.new</code>
 *  without block.
 *
 *  No methods should be called on the worksheet after it is colsed.
 *
 *    wb = XlsxWriter::Workbook.new('/tmp/test.xlsx')
 *    wb.close
 */
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
    rb_thread_call_without_gvl((void *(*)(void *)) workbook_close, ptr->workbook, RUBY_UBF_IO, NULL);
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

/* call-seq:
 *    wb.add_worksheet([name]) -> ws
 *    wb.add_worksheet([name]) { |ws| block } -> obj
 *
 * Adds a worksheet named +name+ to the workbook.
 *
 * If a block is passed, the last statement is returned.
 *
 *    wb.add_worksheet('Sheet1') do |ws|
 *      ws.add_row(['test'])
 *    end
 */
VALUE
workbook_add_worksheet_(int argc, VALUE *argv, VALUE self) {
  VALUE worksheet = Qnil;

  rb_check_arity(argc, 0, 1);

  struct workbook *ptr;
  Data_Get_Struct(self, struct workbook, ptr);
  if (ptr->workbook) {
    worksheet = rb_funcall(cWorksheet, rb_intern("new"), argc + 1, self, argv[0]);
  }

  if (rb_block_given_p()) {
    VALUE res = rb_yield(worksheet);
    return res;
  }

  return worksheet;
}

/* call-seq:
 *    wb.add_format(key, definition) -> wb
 *
 * Adds a format identified as +key+ with parameters set from +definition+ to the
 * workbook.
 *
 * +definition+ should be an object and may contain the following options:
 *
 * :font_name:: Font family to be used to display the cell content (like Arial, Dejavu or Helvetica).
 * :font_size:: Font size.
 * :font_color:: Text color.
 * :bold, :italic, underline :: Bold, italic, underlined text.
 * :font_strikeout:: Striked out text.
 * :font_script:: Superscript (XlsxWrtiter::Format::FONT_SUPERSCRIPT) or subscript (XlsxWriter::Format::FONT_SUBSCRIPT).
 * :num_format:: Defines numerical format with mask, like <code>'d mmm yyyy'</code>
 *               or <code>'#,##0.00'</code>.
 * :num_format_index:: Defines numerical format from special {pre-defined set}[https://libxlsxwriter.github.io/format_8h.html#a688aa42bcc703d17e125d9a34c721872].
 * :unlocked:: Allows modifications of protected cells.
 * :hidden::
 * :align, :vertical_align::
 * :text_wrap::
 * :rotation::
 * :indent::
 * :shrink::
 * :pattern::
 * :bg_color::
 * :fg_color::
 * :border::
 * :bottom, :top, :left, :right::
 * :border_color, :bottom_color, :top_color, :left_color, :right_color::
 */
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

/*  call-seq:
 *     wb.add_chart(type) { |chart| block } -> obj
 *     wb.add_chert(type) -> chart
 *
 *  Adds a chart of type +type+ to the workbook.
 *
 *  +type+ is expected to be one of <code>XlsxWriter::Workbook::Chart::{NONE, AREA,
 *  AREA_STACKED, AREA_STACKED_PERCENT, BAR, BAR_STACKED, BAR_STACKED_PERCENT,
 *  COLUMN, COLUMN_STACKED, COLUMN_STACKED_PERCENT, DOUGHNUT, LINE, PIE, SCATTER,
 *  SCATTER_STRAIGHT, SCATTER_STRAIGHT_WOTH_MARKERS, SCATTER_SMOOTH,
 *  SCATTER_SMOOTH_WITH_MARKERS, RADAR, RADAR_WITH_MARKERS, RADAR_FILLED}</code>.
 *
 *    wb.add_chart(XlsxWriter::Workbook::Chart::PIE) do |chart|
 *      chart.add_series 'A1:A10', 'B1:B10'
 *      ws.insert_chart('D2', chart)
 *    end
 */
VALUE
workbook_add_chart_(VALUE self, VALUE type) {
  VALUE chart = rb_funcall(cChart, rb_intern("new"), 2, self, type);
  if (rb_block_given_p()) {
    VALUE res = rb_yield(chart);
    return res;
  }
  return chart;
}

/* :nodoc: */
VALUE
workbook_set_default_xf_indices_(VALUE self) {
  struct workbook *ptr;
  Data_Get_Struct(self, struct workbook, ptr);
  lxw_workbook_set_default_xf_indices(ptr->workbook);
  return self;
}

/*  call-seq: wb.properties -> wb_properties
 *
 *  Returns worbook properties accessor object.
 *
 *    wb.properties.title = 'My awesome sheet'
 */
VALUE
workbook_properties_(VALUE self) {
  VALUE props = rb_obj_alloc(cWorkbookProperties);
  rb_obj_call_init(props, 1, &self);
  return props;
}

/*  call-seq: wb.define_name(name, formula) -> wb
 *
 *  Create a defined +name+ in the workbook to use as a variable defined in +formula+.
 *
 *    wb.define_name 'Sales', '=Sheet1!$G$1:$H$10'
 */
VALUE
workbook_define_name_(VALUE self, VALUE name, VALUE formula) {
  struct workbook *ptr;
  Data_Get_Struct(self, struct workbook, ptr);
  workbook_define_name(ptr->workbook, StringValueCStr(name), StringValueCStr(formula));
  return self;
}

/*  call-seq:
 *    wb.validate_sheet_name(name) -> true
 *    wb.validate_worksheet_name(name) -> true
 *
 *  Validates a worksheet +name+. Returns +true+ or raises an exception (not
 *  implemented yet).
 *
 */
VALUE
workbook_validate_sheet_name_(VALUE self, VALUE name) {
  struct workbook *ptr;
  lxw_error err;
  Data_Get_Struct(self, struct workbook, ptr);
  err = workbook_validate_sheet_name(ptr->workbook, StringValueCStr(name));
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
  VALUE time_a = rb_funcall(val, rb_intern("to_a"), 0);
  static const VALUE idxs[6] = { INT2FIX(0), INT2FIX(1), INT2FIX(2), INT2FIX(3), INT2FIX(4), INT2FIX(5) };
  lxw_datetime res = {
    .year  = NUM2INT(rb_ary_aref(1, idxs+5, time_a)),
    .month = NUM2INT(rb_ary_aref(1, idxs+4, time_a)),
    .day   = NUM2INT(rb_ary_aref(1, idxs+3, time_a)),
    .hour  = NUM2INT(rb_ary_aref(1, idxs+2, time_a)),
    .min   = NUM2INT(rb_ary_aref(1, idxs+1, time_a)),
    .sec   = NUM2DBL(rb_ary_aref(1, idxs+0, time_a)) +
             NUM2DBL(rb_funcall(val, rb_intern("subsec"), 0))
  };

  return res;
}

void
handle_lxw_error(lxw_error err) {
  return;
}


/*  Document-class: XlsxWriter::Workbook
 *
 *  +Workbook+ is the main class exposed by +XlsxWriter+. It represents the
 *  workbook (.xlsx) file.
 */
void
init_xlsxwriter_workbook() {
  cWorkbook = rb_define_class_under(mXlsxWriter, "Workbook", rb_cObject);

  rb_define_alloc_func(cWorkbook, workbook_alloc);
  rb_define_singleton_method(cWorkbook, "new", workbook_new_, -1);
  rb_define_alias(rb_singleton_class(cWorkbook), "open", "new");
  rb_define_method(cWorkbook, "initialize", workbook_init, -1);
  rb_define_method(cWorkbook, "close", workbook_release, 0);
  rb_define_method(cWorkbook, "add_worksheet", workbook_add_worksheet_, -1);
  rb_define_method(cWorkbook, "add_format", workbook_add_format_, 2);
  rb_define_method(cWorkbook, "add_chart", workbook_add_chart_, 1);
  rb_define_method(cWorkbook, "set_default_xf_indices", workbook_set_default_xf_indices_, 0);
  rb_define_method(cWorkbook, "properties", workbook_properties_, 0);
  rb_define_method(cWorkbook, "define_name", workbook_define_name_, 2);
  rb_define_method(cWorkbook, "validate_sheet_name", workbook_validate_sheet_name_, 1);
  rb_define_method(cWorkbook, "validate_worksheet_name", workbook_validate_sheet_name_, 1);

  /*
   * This attribute contains effective font widths used for automatic column
   * widths of workbook columns.
   */
  rb_define_attr(cWorkbook, "font_sizes", 1, 0);
}
