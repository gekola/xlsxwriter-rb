require_relative './xlsx-func-testcase'

class TestPrintArea < XlsxWriterTestCase
  test 'print_area06' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins' ] }
    ws = wb.add_worksheet
    ws.paper = 9
    ws.vertical_dpi = 200
    ws.print_area(0, 'A', 8, 'F')
    ws.write_string(0, 'A', 'Foo', nil)
  end

  test 'print_area07' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins' ] }
    ws = wb.add_worksheet
    ws.print_area('A1:XFD1048576')
    ws.write_string('A1', 'Foo', nil)
  end
end
