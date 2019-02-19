# frozen_string_literal: true

require_relative './xlsx-func-testcase'

class TestEscapes < XlsxWriterTestCase
  test 'escapes04' do |wb|
    wb.add_worksheet
      .write_url(0, 'A', 'http://www.perl.com/?a=1&b=2')
  end

  test 'escapes05' do |wb|
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
    wb.add_worksheet
      .write_url(0, 'A', 'http://example.com/!"$%&\'( )*+,-./0123456789:;<=>?@' \
                         'ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`'                    \
                         'abcdefghijklmnopqrstuvwxyz{|}~', nil)
  end

  test 'escapes08' do |wb|
    wb.add_worksheet
      .write_url 0, 'A', 'http://example.com/%5b0%5d', string: 'http://example.com/[0]'
  end
end
