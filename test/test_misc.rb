# frozen_string_literal: true

require_relative './xlsx_func_testcase'

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

  test 'shared_strings01' do |wb|
    ws = wb.add_worksheet
    ws.write_string(0, 0, '_x0000_')
    (1...127).each_with_object(String.new('')) do |i, s|
      s[0] = i.chr
      ws.write_string(i, 0, s) unless i == 34
    end
  end

  test 'gh42_01' do |wb|
    wb.add_worksheet.write_string(0, 0, "\xe5\x9b\xbe\x14\xe5\x9b\xbe", nil)
  end

  test 'gh42_02' do |wb|
    wb.add_worksheet.write_string(0, 0, "\xe5\x9b\xbe\x20\xe5\x9b\xbe", nil)
  end
end
