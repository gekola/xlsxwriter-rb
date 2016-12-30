require_relative './xlsx-func-testcase'

class TestRowColFormat < XlsxWriterTestCase
  test 'row_col_format02' do |wb|
    wb.add_format(:b, bold: true)
    wb.add_worksheet
      .set_row(0, height: 15, format: :b)
      .write_string(0, 0, 'Foo', nil)
  end

  test 'row_col_format09' do |wb|
    wb
      .add_format(:bold, bold: true)
      .add_format(:mixed, bold: true, italic: true)
      .add_format(:italic, italic: true)
      .set_default_xf_indices
      .add_worksheet
      .set_row(4, height: 15, format: :bold)
      .set_column(2, 2, width: 8.43, format: :italic)
      .write_string(0, 2, 'Foo', nil)
      .write_string(4, 0, 'Foo', nil)
      .write_string(4, 2, 'Foo', :mixed)
  end

  test 'row_col_format14' do |wb|
    wb
      .add_format(:bold, bold: true)
      .add_worksheet
      .set_column(1, 3, width: 5)
      .set_column(5, 5, width: 8)
      .set_column(7, 7, format: :bold)
      .set_column(9, 9, width: 2)
      .set_column(11, 11, hide: true)
  end
end
