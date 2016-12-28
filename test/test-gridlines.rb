require_relative './xlsx-func-testcase'

class TestGridlines < XlsxWriterTestCase
  test 'gridlines01' do |wb|
    ws = wb.add_worksheet

    ws.set_paper(9)
    worksheet.vertical_dpi = 200

    worksheet_gridlines = XlsxWriter::Worksheet::LXW_HIDE_ALL_GRIDLINES

    worksheet_write_string(worksheet, 0, "A", "Foo" , nil)
  end
end
