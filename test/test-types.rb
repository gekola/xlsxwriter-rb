# frozen_string_literal: true

require_relative './xlsx-func-testcase'

class TestTypes < XlsxWriterTestCase
  test 'types02' do |wb|
    wb.add_worksheet
      .write_boolean(0, 'A', true,  nil)
      .write_boolean(1, 'A', false, nil)
  end

  test 'types08' do |wb|
    wb.add_format(:bold, bold: true)
      .add_format(:ita, italic: true)
      .add_worksheet
      .write_boolean(0, 'A', true, :bold)
      .write_boolean(1, 'A', false, :ita)
  end
end
