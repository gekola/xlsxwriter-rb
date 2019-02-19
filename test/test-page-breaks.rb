# frozen_string_literal: true

require_relative './xlsx-func-testcase'

class TestpageBreaks < XlsxWriterTestCase
  test 'page_breaks06' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins' ] }
    ws = wb.add_worksheet
    ws.paper = 9
    ws.vertical_dpi = 200
    ws.h_pagebreaks = [1, 5, 8, 13]
    ws.v_pagebreaks = [1, 3, 8]
    ws.write_string(0, 'A', 'Foo', nil)
  end
end
