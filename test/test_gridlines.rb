# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestGridlines < XlsxWriterTestCase
  test 'gridlines01' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => ['<pageMargins'] }
    ws = wb.add_worksheet

    ws.paper = 9
    ws.vertical_dpi = 200

    ws.gridlines = XlsxWriter::Worksheet::GRIDLINES_HIDE_ALL

    ws.write_string(0, 'A', 'Foo', nil)
  end
end
