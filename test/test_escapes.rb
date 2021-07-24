# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestEscapes < XlsxWriterTestCase
  test 'escapes03' do |wb|
    wb
      .add_format(:bold, bold: true)
      .add_format(:italic, italic: true)
      .add_worksheet
      .write_string('A1', 'Foo', :bold)
      .write_string('A2', 'Bar', :italic)
      .write_rich_string('A3', [['a'], ["b\"<>'c", :bold], ['defg']])
  end

  test 'escapes04' do |wb|
    wb.unset_default_url_format
    wb.add_worksheet
      .write_url(0, 'A', 'http://www.perl.com/?a=1&b=2')
  end

  test 'escapes05' do |wb|
    wb.unset_default_url_format
    wb.add_worksheet('Start')
      .write_url(0, 'A', 'internal:\'A & B\'!A1', string: 'Jump to A & B')
    wb.add_worksheet('A & B')
  end

  test 'escapes06' do |wb|
    ws = wb.add_worksheet
    ws.set_column(0, 0, width: 14)
    wb.add_format :f1, num_format: '[Red]0.0%\\ "a"'
    ws.write_number(0, 'A', 123, :f1)
  end

  test 'escapes07' do |wb|
    wb.unset_default_url_format
    wb.add_worksheet
      .write_url(0, 'A', 'http://example.com/!"$%&\'( )*+,-./0123456789:;<=>?@' \
                         'ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`'                    \
                         'abcdefghijklmnopqrstuvwxyz{|}~', nil)
  end

  test 'escapes08' do |wb|
    wb.unset_default_url_format
    wb.add_worksheet
      .write_url 0, 'A', 'http://example.com/%5b0%5d', string: 'http://example.com/[0]'
  end
end
