require_relative './xlsx-func-testcase'

class TestPrintOptions < XlsxWriterTestCase
  test 'print_options01' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins' ] }
    ws = wb.add_worksheet
    ws.paper = 9
    ws.vertical_dpi = 200
    ws.gridlines = XlsxWriter::Worksheet::GRIDLINES_SHOW_PRINT
    ws.write_string(0, 'A', 'Foo', nil)
  end

  test 'print_options02' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins' ] }
    ws = wb.add_worksheet
    ws.paper = 9
    ws.vertical_dpi = 200
    ws.center_horizontally
    ws.write_string(0, 'A', 'Foo', nil)
  end

  test 'print_options03' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins' ] }
    ws = wb.add_worksheet
    ws.paper = 9
    ws.vertical_dpi = 200
    ws.center_vertically
    ws.write_string(0, 'A', 'Foo', nil)
  end

  test 'print_options04' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins' ] }
    ws = wb.add_worksheet
    ws.paper = 9
    ws.vertical_dpi = 200
    ws.print_row_col_headers
    ws.write_string(0, 'A', 'Foo', nil)
  end

  test 'print_options05' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins' ] }
    ws = wb.add_worksheet
    ws.paper = 9
    ws.vertical_dpi = 200
    ws.gridlines = XlsxWriter::Worksheet::GRIDLINES_SHOW_PRINT
    ws.center_horizontally
    ws.center_vertically
    ws.print_row_col_headers
    ws.write_string(0, 'A', 'Foo', nil)
  end

  test 'print_options06' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins' ] }
    ws = wb.add_worksheet
    ws.paper = 9
    ws.vertical_dpi = 200
    ws.print_area(0, 'A', 19, 'G')
    ws.repeat_rows(0, 0)
    ws.write_string(0, 'A', 'Foo', nil)
  end
end
