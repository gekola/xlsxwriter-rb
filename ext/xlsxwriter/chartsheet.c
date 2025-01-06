#include "chart.h"
#include "chartsheet.h"
#include "common.h"
#include "workbook.h"
#include "xlsxwriter_ext.h"

VALUE cChartsheet;

void chartsheet_clear(void *p);
void chartsheet_free(void *p);

size_t
chartsheet_size(const void *data) {
  return sizeof(struct chartsheet);
}

const rb_data_type_t chartsheet_type = {
	.wrap_struct_name = "chartsheet",
	.function = {
		.dmark = NULL,
		.dfree = chartsheet_free,
		.dsize = chartsheet_size,
	},
	.data = NULL,
	.flags = RUBY_TYPED_FREE_IMMEDIATELY,
};

VALUE
chartsheet_alloc(VALUE klass) {
  VALUE obj;
  struct chartsheet *ptr;

  obj = TypedData_Make_Struct(klass, struct chartsheet, &chartsheet_type, ptr);

  ptr->chartsheet = NULL;

  return obj;
}

/* :nodoc: */
VALUE
chartsheet_init(int argc, VALUE *argv, VALUE self) {
  char *name = NULL;
  struct workbook *wb_ptr;
  struct chartsheet *ptr;

  TypedData_Get_Struct(self, struct chartsheet, &chartsheet_type, ptr);

  if (argc == 2) {
    name = StringValueCStr(argv[1]);
  }

  rb_iv_set(self, "@workbook", argv[0]);

  TypedData_Get_Struct(argv[0], struct workbook, &workbook_type, wb_ptr);
  ptr->chartsheet = workbook_add_chartsheet(wb_ptr->workbook, name);
  if (!ptr->chartsheet)
    rb_raise(eXlsxWriterError, "chartsheet was not created");
  return self;
}

/* :nodoc: */
VALUE
chartsheet_release(VALUE self) {
  struct chartsheet *ptr;
  TypedData_Get_Struct(self, struct chartsheet, &chartsheet_type, ptr);

  chartsheet_clear(ptr);
  return self;
}

void
chartsheet_clear(void *p) {
  struct chartsheet *ptr = p;

  if (ptr->chartsheet)
    ptr->chartsheet = NULL;
}

void
chartsheet_free(void *p) {
  chartsheet_clear(p);
  ruby_xfree(p);
}

/*  call-seq: cs.activate -> self
 *
 *  Set the chartsheet to be active on opening the workbook.
 */
VALUE
chartsheet_activate_(VALUE self) {
  LXW_NO_RESULT_CALL(chartsheet, activate);
  return self;
}

/*  call-seq: cs.chart = chart
 *
 *  Set the chart chartsheet contains.
 */
VALUE
chartsheet_set_chart_(VALUE self, VALUE chart) {
  struct chart *chart_ptr;
  TypedData_Get_Struct(chart, struct chart, &chart_type, chart_ptr);

  LXW_ERR_RESULT_CALL(chartsheet, set_chart, chart_ptr->chart);
  return self;
}

/*  call-seq: cs.hide -> self
 *
 *  Hide the chartsheet.
 */
VALUE
chartsheet_hide_(VALUE self) {
  LXW_NO_RESULT_CALL(chartsheet, hide);
  return self;
}

/*  call-seq: cs.paper=(type) -> type
 *
 *  Sets the paper +type+ for printing. See {libxlsxwriter doc}[https://libxlsxwriter.github.io/worksheet_8h.html#a9f8af12017797b10c5ee68e04149032f] for options.
 */
VALUE
chartsheet_set_paper_(VALUE self, VALUE paper_type) {
  LXW_NO_RESULT_CALL(chartsheet, set_paper, NUM2INT(paper_type));
  return self;
}

/*  call-seq: cs.protect(password[, opts]) -> self
 *
 *  Protects the worksheet elements from modification.
 *  See {libxlsxwriter doc}[https://libxlsxwriter.github.io/worksheet_8h.html#a1b49e135d4debcdb25941f2f40f04cb0]
 *  for options.
 */
VALUE
chartsheet_protect_(int argc, VALUE *argv, VALUE self) {
  rb_check_arity(argc, 0, 2);
  struct _protect_options po = _extract_protect_options(argc, argv);
  LXW_NO_RESULT_CALL(chartsheet, protect, po.password, (po.with_options ? &po.options : NULL));
  return self;
}

/*  call-seq: cs.set_fitst_sheet -> self
 *
 *  Set the sheet to be the first visible in the sheets list (which is placed
 *  under the work area in Excel).
 */
VALUE
chartsheet_set_first_sheet_(VALUE self) {
  LXW_NO_RESULT_CALL(chartsheet, set_first_sheet);
  return self;
}

/*  call-seq: cs.set_footer(text[, opts])
 *
 *  See Worksheet#set_header for params description.
 */
VALUE
chartsheet_set_footer_(int argc, VALUE *argv, VALUE self) {
  rb_check_arity(argc, 1, 2);
  struct _header_options ho = _extract_header_options(argc, argv);
  struct chartsheet *ptr;
  TypedData_Get_Struct(self, struct chartsheet, &chartsheet_type, ptr);
  lxw_error err;
  if (!ho.with_options) {
    err = chartsheet_set_footer(ptr->chartsheet, ho.str);
  } else {
    err = chartsheet_set_footer_opt(ptr->chartsheet, ho.str, &ho.options);
  }
  handle_lxw_error(err);
  return self;
}

/*  call-seq: cs.select -> self
 *
 *  Set the chartsheet to be selected on opening the workbook.
 */
VALUE
chartsheet_select_(VALUE self) {
  LXW_NO_RESULT_CALL(chartsheet, select);
  return self;
}

/*  call-seq: cs.set_header(text[, opts])
 *
 *  See Worksheet#set_header for params description.
 */
VALUE
chartsheet_set_header_(int argc, VALUE *argv, VALUE self) {
  rb_check_arity(argc, 1, 2);
  struct _header_options ho = _extract_header_options(argc, argv);
  struct chartsheet *ptr;
  TypedData_Get_Struct(self, struct chartsheet, &chartsheet_type, ptr);
  lxw_error err;
  if (!ho.with_options) {
    err = chartsheet_set_header(ptr->chartsheet, ho.str);
  } else {
    err = chartsheet_set_header_opt(ptr->chartsheet, ho.str, &ho.options);
  }
  handle_lxw_error(err);
  return self;
}

/*  call-seq: cs.set_landscape -> self
 *
 *  Sets print orientation of the chartsheet to landscape.
 */
VALUE
chartsheet_set_landscape_(VALUE self) {
  LXW_NO_RESULT_CALL(chartsheet, set_landscape);
  return self;
}

/* Sets the chartsheet margins (Numeric) for the printed page. */
VALUE
chartsheet_set_margins_(VALUE self, VALUE left, VALUE right, VALUE top, VALUE bottom) {
  LXW_NO_RESULT_CALL(chartsheet, set_margins, NUM2DBL(left), NUM2DBL(right), NUM2DBL(top), NUM2DBL(bottom));
  return self;
}

/*  call-seq: cs.set_portrait -> self
 *
 *  Sets print orientation of the chartsheet to portrait.
 */
VALUE
chartsheet_set_portrait_(VALUE self) {
  LXW_NO_RESULT_CALL(chartsheet, set_portrait);
  return self;
}

/*  call-seq: cs.tab_color=(color) -> color
 *
 *  Set the tab color (from RGB integer).
 *
 *     ws.tab_color = 0xF0F0F0
 */
VALUE
chartsheet_set_tab_color_(VALUE self, VALUE color) {
  LXW_NO_RESULT_CALL(chartsheet, set_tab_color, NUM2INT(color));
  return color;
}

/*  call-seq: cs.zoom=(zoom) -> zoom
 *
 *  Sets the worksheet zoom factor in the range 10 <= +zoom+ <= 400.
 */
VALUE
chartsheet_set_zoom_(VALUE self, VALUE val) {
  LXW_NO_RESULT_CALL(chartsheet, set_zoom, NUM2INT(val));
  return self;
}

/*  call-seq: cs.horizontal_dpi -> int
 *
 *  Returns the horizontal dpi.
 */
VALUE
chartsheet_get_horizontal_dpi_(VALUE self) {
  struct chartsheet *ptr;
  TypedData_Get_Struct(self, struct chartsheet, &chartsheet_type, ptr);
  return INT2NUM(ptr->chartsheet->worksheet->horizontal_dpi);
}

/*  call-seq: cs.horizontal_dpi=(dpi) -> dpi
 *
 *  Sets the horizontal dpi.
 */
VALUE
chartsheet_set_horizontal_dpi_(VALUE self, VALUE val) {
  struct chartsheet *ptr;
  TypedData_Get_Struct(self, struct chartsheet, &chartsheet_type, ptr);
  ptr->chartsheet->worksheet->horizontal_dpi = NUM2INT(val);
  return val;
}

/*  call-seq: cs.vertical_dpi -> int
 *
 *  Returns the vertical dpi.
 */
VALUE
chartsheet_get_vertical_dpi_(VALUE self) {
  struct chartsheet *ptr;
  TypedData_Get_Struct(self, struct chartsheet, &chartsheet_type, ptr);
  return INT2NUM(ptr->chartsheet->worksheet->vertical_dpi);
}

/*  call-seq: cs.vertical_dpi=(dpi) -> dpi
 *
 *  Sets the vertical dpi.
 */
VALUE
chartsheet_set_vertical_dpi_(VALUE self, VALUE val) {
  struct chartsheet *ptr;
  TypedData_Get_Struct(self, struct chartsheet, &chartsheet_type, ptr);
  ptr->chartsheet->worksheet->vertical_dpi = NUM2INT(val);
  return val;
}



void
init_xlsxwriter_chartsheet() {
  /* Xlsx chartsheet */
  cChartsheet = rb_define_class_under(mXlsxWriter, "Chartsheet", rb_cObject);

  rb_define_alloc_func(cChartsheet, chartsheet_alloc);
  rb_define_method(cChartsheet, "initialize", chartsheet_init, -1);
  rb_define_method(cChartsheet, "free", chartsheet_release, 0);
  rb_define_method(cChartsheet, "activate", chartsheet_activate_, 0);
  rb_define_method(cChartsheet, "chart=", chartsheet_set_chart_, 1);
  rb_define_method(cChartsheet, "hide", chartsheet_hide_, 0);
  rb_define_method(cChartsheet, "paper=", chartsheet_set_paper_, 1);
  rb_define_method(cChartsheet, "protect", chartsheet_protect_, -1);
  rb_define_method(cChartsheet, "select", chartsheet_select_, 0);
  rb_define_method(cChartsheet, "set_first_sheet", chartsheet_set_first_sheet_, 0);
  rb_define_method(cChartsheet, "set_footer", chartsheet_set_footer_, -1);
  rb_define_method(cChartsheet, "set_header", chartsheet_set_header_, -1);
  rb_define_method(cChartsheet, "set_landscape", chartsheet_set_landscape_, 0);
  rb_define_method(cChartsheet, "set_margins", chartsheet_set_margins_, 4);
  rb_define_method(cChartsheet, "set_portrait", chartsheet_set_portrait_, 0);
  rb_define_method(cChartsheet, "tab_color=", chartsheet_set_tab_color_, 1);
  rb_define_method(cChartsheet, "zoom=", chartsheet_set_zoom_, 1);

  rb_define_method(cChartsheet, "horizontal_dpi", chartsheet_get_horizontal_dpi_, 0);
  rb_define_method(cChartsheet, "horizontal_dpi=", chartsheet_set_horizontal_dpi_, 1);
  rb_define_method(cChartsheet, "vertical_dpi", chartsheet_get_vertical_dpi_, 0);
  rb_define_method(cChartsheet, "vertical_dpi=", chartsheet_set_vertical_dpi_, 1);
}
