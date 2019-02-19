# frozen_string_literal: true

require_relative './xlsx-func-testcase'

class TestMisc < XlsxWriterTestCase
  test 'tab_color01' do |wb|
    wb.add_worksheet
      .write_string(0, 0, 'Foo', nil)
      .tab_color = XlsxWriter::Format::COLOR_RED
  end

  test 'firstsheet01' do |wb|
    wss = 20.times.map { wb.add_worksheet }
    wss[7].set_first_sheet
    wss[19].activate
  end

  test 'hide01' do |wb|
    wb.add_worksheet
    wb.add_worksheet.hide
    wb.add_worksheet
  end

  test 'gh42_01' do |wb|
    wb.add_worksheet.write_string(0, 0, "\xe5\x9b\xbe\x14\xe5\x9b\xbe", nil)
  end

  test 'gh42_02' do |wb|
    wb.add_worksheet.write_string(0, 0, "\xe5\x9b\xbe\x20\xe5\x9b\xbe", nil)
  end
end
