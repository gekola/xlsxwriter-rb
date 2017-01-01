require_relative './xlsx-func-testcase'

class TestPageSetup < XlsxWriterTestCase
  test 'page_view01' do |wb|
    ws = wb.add_worksheet
    ws.set_page_view
    ws.write_string(0, 'A', 'Foo', nil)
    ws.paper = 9
    ws.vertical_dpi = 200
  end

  test 'landscape01' do |wb|
    ws = wb.add_worksheet
    ws.write_string(0, 'A', 'Foo', nil)
    ws.set_landscape
    ws.paper = 9
    ws.vertical_dpi = 200
  end

  test 'print_across01' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins'] }
    ws = wb.add_worksheet
    ws.print_across
    ws.paper = 9
    ws.vertical_dpi = 200
    ws.write_string 0, 0, 'Foo', nil
  end
end
