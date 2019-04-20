# frozen_string_literal: true

require 'test/unit'
require 'support/with_xlsx_file'

class TestErrors < Test::Unit::TestCase
  include WithXlsxFile

  def test_long_sheet_name
    with_xlsx_file('tmp/error_long_sheet_name.xlsx') do |wb|
      was_in_block = false
      assert_raise(XlsxWriter::Error.new('worksheet was not created')) do
        wb.add_worksheet('a' * 32) do |ws|
          was_in_block = true
          ws.add_row(['test'])
        end
      end
      assert_equal(was_in_block, false)
    end
  end

  def test_written_rich_string
    with_xlsx_file('tmp/error_written_rich_string.xlsx') do |wb|
      ws = wb.add_worksheet
      rs = XlsxWriter::RichString.new(wb, ['1'])
      rs << ' 2'
      ws.add_row([rs])
      assert_raise(XlsxWriter::Error.new('Cannot modify as the RichString has already been written to a workbook.')) do
        rs << ' 3'
      end
    end
  end
end
