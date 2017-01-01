require_relative './xlsx-func-testcase'

class TestPrintScale < XlsxWriterTestCase
  test 'print_scale01' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins' ] }
    ws = wb.add_worksheet
    ws.print_scale = 75
    ws.paper = 9
    ws.vertical_dpi = 200
    ws.write_string(0, 'A', 'Foo', nil)
  end
end
