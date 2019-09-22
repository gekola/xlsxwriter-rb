# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestSetStartPage < XlsxWriterTestCase
  test 'set_start_page02' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => ['<pageMargins'] }
    wb.add_worksheet do |ws|
      ws.start_page = 2
      ws.paper = 9
      ws.vertical_dpi = 200
      ws.write_string(0, 'A', 'Foo', nil)
    end
  end
end
