# frozen_string_literal: true

require 'test/unit'

class TestErrors < Test::Unit::TestCase
  def test_long_sheet_name
    file_path = 'tmp/error_long_sheet_name.xlsx'
    was_in_block = false
    XlsxWriter::Workbook.open(file_path) do |wb|
      assert_raise(RuntimeError.new('worksheet was not created')) do
        wb.add_worksheet('a' * 32) do |ws|
          was_in_block = true
          ws.add_row(['test'])
        end
      end
    end
    assert_equal(was_in_block, false)
  ensure
    File.delete file_path if file_path && File.exist?(file_path)
  end
end
