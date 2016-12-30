require_relative './xlsx-func-testcase'

class TestMergeRange < XlsxWriterTestCase
  test 'merge_range03' do |wb|
    wb.add_format(:f, align: XlsxWriter::Format::ALIGN_CENTER)
    wb.add_worksheet
      .merge_range(1, 1, 1, 2, 'Foo', :f)
      .merge_range(1, 3, 1, 4, 'Foo', :f)
      .merge_range(1, 5, 1, 6, 'Foo', :f)
  end

  test 'merge_range05' do |wb|
    wb.add_format(:f, align: XlsxWriter::Format::ALIGN_CENTER)
    wb.add_worksheet
      .merge_range(1, 1, 1, 3, '', :f)
      .write_number(1, 1, 123, :f)
  end
end
