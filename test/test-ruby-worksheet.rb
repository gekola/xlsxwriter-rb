# frozen_string_literal: true

require_relative './support/with_xlsx_file'

class TestRubyWorksheet < Test::Unit::TestCase
  include WithXlsxFile

  def test_auto_width
    with_xlsx_file do |wb|
      ws = wb.add_worksheet

      ws.add_row ['asd', 'asdf', 'asdfg', 'asdfgh', 'asdfghj', 'd'*250]
      assert_in_epsilon(4.4, ws.get_column_auto_width('A'))
      assert_in_epsilon(4.4, ws.get_column_auto_width('B'))
      assert_in_epsilon(5.5, ws.get_column_auto_width('C'))
      assert_in_epsilon(6.6, ws.get_column_auto_width('D'))
      assert_in_epsilon(6.6, ws.get_column_auto_width('E'))
      assert_in_epsilon(255, ws.get_column_auto_width('F'))

      wb.add_format :f, font_size: 14
      ws.add_row ['asd', 'asdf', 'asdfg', 'asdfgh', 'asdfghj'], style: :f
      assert_in_epsilon(5.6, ws.get_column_auto_width('A'))
      assert_in_epsilon(5.6, ws.get_column_auto_width('B'))
      assert_in_epsilon(7.0, ws.get_column_auto_width('C'))
      assert_in_epsilon(8.4, ws.get_column_auto_width('D'))
      assert_in_epsilon(8.4, ws.get_column_auto_width('E'))

      assert_raise(XlsxWriter::Error.new('Function parameter validation error.')) do
        ws.write_rich_string('A2', [])
      end
    end
  end
end
