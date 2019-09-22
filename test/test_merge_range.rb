# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestMergeRange < XlsxWriterTestCase
  test 'merge_range03' do |wb|
    wb.add_format(:f, align: XlsxWriter::Format::ALIGN_CENTER)
    wb.add_worksheet
      .merge_range(1, 1, 1, 2, 'Foo', :f)
      .merge_range('D2:E2',    'Foo', :f)
      .merge_range('F2', 'G2', 'Foo', :f)
  end

  test 'merge_range05' do |wb|
    wb.add_format(:f, align: XlsxWriter::Format::ALIGN_CENTER)
    wb.add_worksheet
      .merge_range('B2:D2', '', :f)
      .write_number(1, 1, 123, :f)
  end
end
