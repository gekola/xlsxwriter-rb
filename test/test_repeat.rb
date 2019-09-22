# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestRepeat < XlsxWriterTestCase
  test 'repeat05' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => ['<pageMargins'],
                          'xl/worksheets/sheet3.xml' => ['<pageMargins'] }
    wb.add_worksheet do |ws|
      ws.paper = 9
      ws.vertical_dpi = 200
      ws.repeat_rows(0, 0)
      ws.write_string(0, 'A', 'Foo', nil)
    end

    wb.add_worksheet

    wb.add_worksheet do |ws|
      ws.paper = 9
      ws.vertical_dpi = 200
      ws.repeat_rows(2, 3)
      ws.repeat_columns(1, 5)
    end
  end
end
